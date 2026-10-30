# -*- coding: utf-8 -*-
"""
ENVIO DE ÁLBUM PARA APROVAÇÃO NO MAISBOOK

Para cada pessoa com STATUS DIAGRAMADO na planilha ALBUNS 25.2:
  1. Procura no servidor os JPG do álbum e a capa (sem copiar para o lote)
  2. Abre o MaisBook no Chrome, pesquisa pelo SGE e abre o Drive do cliente
  3. Cria a pasta ALBUM APROVAÇAO
     - um álbum: os JPG e a capa entram direto nessa pasta
     - dois álbuns (24x34, que também leva o 20x20): cria 20x20 e 24x34
  4. Espera o Drive terminar de carregar as fotos
  5. Copia o link da pasta ALBUM APROVAÇAO
  6. No cartão do cliente, abre Editar, cola o link em Álbum pronto e salva
  7. Só então escolhe o template Aprovação de álbum e envia o e-mail
  8. Só marca como enviado depois que o MaisBook confirmar

Uso:
  python ENVIO_APROVACAO_MAISBOOK.PY                 envia de verdade
  python ENVIO_APROVACAO_MAISBOOK.PY --sge 93441     uma pessoa
  python ENVIO_APROVACAO_MAISBOOK.PY --limite 1      só a primeira
  python ENVIO_APROVACAO_MAISBOOK.PY --simular       procura e abre o cliente, sem pasta, sem arquivo e sem e-mail
  python ENVIO_APROVACAO_MAISBOOK.PY --continuar     retoma quem ainda não está em envio_aprovacao_feitos.json
  python ENVIO_APROVACAO_MAISBOOK.PY --listar       só lista quem está DIAGRAMADO

Janela: guia Enviar aprovação em organizador_ui.pyw
"""

import argparse
import csv
import json
import re
import sys
import time
from datetime import datetime
from importlib.machinery import SourceFileLoader
from pathlib import Path
from urllib.parse import quote

PASTA = Path(__file__).resolve().parent
PASTA_MB = PASTA.parent / "cpf_maisbook"
proc = SourceFileLoader("procura_servidor", str(PASTA / "PROCUARAÇAO_SERVIDOR.PY")).load_module()
mb = SourceFileLoader("cpf_maisbook", str(PASTA_MB / "CPF_MAISBOOK.PY")).load_module()
base = mb.base
nomes_batem = mb.nomes_batem

PLANILHA_LINK = "https://docs.google.com/spreadsheets/d/1FeYQop9z6rVAFQuIs1EOiuQcE84t0iBgamYtMWUIAU0/edit?usp=sharing"
PLANILHA = PASTA / "planilha_envio_aprovacao.xlsx"
FEITOS = PASTA / "envio_aprovacao_feitos.json"
RELATORIO = PASTA / "relatorio_envio_aprovacao.csv"
PARAR = PASTA / "envio_aprovacao_parar.flag"
PRINTS = PASTA / "envio_prints"

PASTA_RAIZ = "ALBUM APROVAÇAO"
PASTA_TAMANHO = {"20X20": "20x20", "24X34": "24x34"}
ESPERA_DRIVE_S = 30
ESPERA_UPLOAD_S = 900
TEMPLATE_VALOR = "aprovacao_album"


def _configurar_saida():
    for canal in (sys.stdout, sys.stderr):
        if hasattr(canal, "reconfigure"):
            try:
                canal.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass


_configurar_saida()


class PararPessoa(Exception):
    def __init__(self, resultado, detalhe):
        super().__init__(detalhe)
        self.resultado = resultado
        self.detalhe = detalhe


def normalizar(texto):
    return proc.normalizar(texto)


def carregar_json(arquivo):
    try:
        return json.loads(Path(arquivo).read_text(encoding="utf-8"))
    except Exception:
        return {}


def salvar_json(arquivo, dados):
    Path(arquivo).write_text(json.dumps(dados, indent=2, ensure_ascii=False), encoding="utf-8")


def salvar_relatorio(linhas):
    campos = ["NOME", "SGE", "CONTRATO", "CODIGO", "TAMANHOS", "JPG", "CAPA", "RESULTADO", "MENSAGEM", "AVISO"]
    try:
        with RELATORIO.open("w", newline="", encoding="utf-8-sig") as arquivo:
            escritor = csv.DictWriter(arquivo, fieldnames=campos, delimiter=";")
            escritor.writeheader()
            for linha in linhas:
                escritor.writerow({campo: linha.get(campo, "") for campo in campos})
    except PermissionError:
        print(f"    (não consegui gravar {RELATORIO.name} — ele está aberto no Excel?)")


def tirar_print(pagina, nome):
    if pagina is None:
        return ""
    PRINTS.mkdir(exist_ok=True)
    arquivo = PRINTS / f"{datetime.now():%Y%m%d_%H%M%S}_{re.sub(r'[^0-9A-Za-z_-]', '_', nome)}.png"
    try:
        pagina.screenshot(path=str(arquivo), full_page=True)
        return str(arquivo)
    except Exception:
        return ""


# ==========================================================
# Planilha
# ==========================================================

COLUNA_CODIGO = 9
NOMES_CABECALHO_CLIENTE = (
    "nome",
    "nome completo",
    "nome do cliente",
    "nome do aluno",
    "nome do formando",
    "cliente",
    "aluno",
    "formando",
)


def limpar_sge(valor):
    texto = str(valor or "").strip()
    if texto.upper().startswith("633."):
        texto = texto[4:].strip()
    texto = re.sub(r"\.\d+$", "", texto)
    return texto.strip()


def cabecalho_da_folha(folha):
    """Acha a linha que tem NOME/CLIENTE, SGE, TAMANHO e STATUS juntos, priorizando a coluna de nome preenchida."""

    linhas_amostra = list(folha.iter_rows(max_row=80, values_only=True))
    for numero, linha in enumerate(linhas_amostra[:40], start=1):
        textos = [normalizar(valor) for valor in linha]
        conjunto = set(textos)
        if not {"sge", "tamanho", "status"} <= conjunto:
            continue
        candidatos_nome = [
            (0 if texto == "nome" else 1, i + 1)
            for i, texto in enumerate(textos)
            if texto in NOMES_CABECALHO_CLIENTE
        ]
        if not candidatos_nome:
            continue

        melhor_nome = candidatos_nome[0][1]
        melhor_chave = None
        for prio, ind in candidatos_nome:
            preenchidas = sum(
                1
                for l_abaixo in linhas_amostra[numero:]
                if ind <= len(l_abaixo) and normalizar(l_abaixo[ind - 1])
            )
            chave = (preenchidas > 0, -prio, preenchidas, -ind)
            if melhor_chave is None or chave > melhor_chave:
                melhor_chave = chave
                melhor_nome = ind

        contrato = next((i + 1 for i, texto in enumerate(textos) if texto in ("cont", "contrato", "ctt")), 0)
        codigo = textos.index("maisbook") + 1 if "maisbook" in textos else COLUNA_CODIGO
        return numero, {
            "nome": melhor_nome,
            "sge": textos.index("sge") + 1,
            "tamanho": textos.index("tamanho") + 1,
            "status": textos.index("status") + 1,
            "contrato": contrato,
            "codigo": codigo,
        }
    return None


def so_contrato(valor):
    digitos = re.sub(r"\D", "", str(valor or "").split("/")[-1].split(".")[0])
    return digitos.lstrip("0") or ("0" if digitos else "")


def celula(linha, indice):
    if not indice or indice > len(linha):
        return ""
    return proc.texto_celula(linha[indice - 1])


def tamanhos_do_plano(plano):
    tamanho = proc.tamanho_do_plano(plano)
    if not tamanho:
        return []
    if tamanho == "24X34":
        return ["20X20", "24X34"]
    return [tamanho]


def ler_diagramados():
    caminho = proc.baixar_planilha_link(PLANILHA_LINK, PLANILHA)
    livro = proc.load_workbook(caminho, data_only=True)
    melhor = None
    aba_padrao = None
    for nome in livro.sheetnames:
        folha = livro[nome]
        achado = cabecalho_da_folha(folha)
        if not achado:
            continue
        linha_cab, colunas = achado
        if aba_padrao is None:
            aba_padrao = (0, nome, linha_cab, colunas)
        qtd = 0
        for numero, linha in enumerate(folha.iter_rows(values_only=True), start=1):
            if numero <= linha_cab:
                continue
            if normalizar(celula(linha, colunas["status"])) == "diagramado":
                qtd += 1
        if melhor is None or qtd > melhor[0]:
            melhor = (qtd, nome, linha_cab, colunas)
    if not melhor:
        melhor = aba_padrao
    if not melhor:
        livro.close()
        raise ValueError("Não achei uma aba com o cabeçalho NOME/CLIENTE, SGE, TAMANHO e STATUS")

    _, nome_aba, linha_cab, colunas = melhor
    folha = livro[nome_aba]
    pessoas = {}
    for numero, linha in enumerate(folha.iter_rows(values_only=True), start=1):
        if numero <= linha_cab:
            continue
        if normalizar(celula(linha, colunas["status"])) != "diagramado":
            continue
        nome = celula(linha, colunas["nome"]).strip()
        sge = limpar_sge(celula(linha, colunas["sge"]))
        plano = celula(linha, colunas["tamanho"]).strip()
        if not nome or not re.fullmatch(r"\d+", sge):
            continue
        pessoa = pessoas.setdefault(sge, {
            "nome": nome,
            "sge": sge,
            "contrato": so_contrato(celula(linha, colunas["contrato"])),
            "codigo": re.sub(r"\s+", "", celula(linha, colunas["codigo"])).upper(),
            "linha": numero,
            "aba": nome_aba,
            "col_status": colunas["status"],
            "planos": [],
            "tamanhos": [],
            "avisos": [],
            "duplicados": [],
        })
        if plano and plano not in pessoa["planos"]:
            pessoa["planos"].append(plano)
        for tamanho in tamanhos_do_plano(plano):
            if tamanho not in pessoa["tamanhos"]:
                pessoa["tamanhos"].append(tamanho)
    livro.close()
    lista = list(pessoas.values())
    for pessoa in lista:
        if not pessoa["codigo"]:
            pessoa["avisos"].append("sem código MaisBook na coluna I — pesquisado pelo SGE")
        if not pessoa["contrato"]:
            pessoa["avisos"].append("sem contrato na planilha — conferido pelo SGE")
    print(f"Aba: {nome_aba} (coluna Nome={colunas['nome']}, SGE={colunas['sge']}, Status={colunas['status']}) | STATUS DIAGRAMADO: {len(lista)}")
    return lista



# ==========================================================
# Servidor — só procura, não copia
# ==========================================================

IGNORAR_NO_NOME = {"da", "de", "do", "dos", "das", "e", "di", "du"}


def capa_e_da_pessoa(nome, arquivo):
    """As duas primeiras palavras do nome têm que aparecer inteiras no arquivo, não só 'da' ou 'silva'."""

    palavras = [p for p in normalizar(nome).split() if len(p) > 2 and p not in IGNORAR_NO_NOME]
    if not palavras:
        return False
    texto = set(normalizar(str(arquivo)).split())
    return all(palavra in texto for palavra in palavras[:2])


def procurar_capa_da_pessoa(albuns, nome, tamanho):
    pasta_capas = albuns / proc.PASTA_CAPAS
    if not pasta_capas.exists():
        return None
    melhor = None
    for arquivo in proc.listar_arquivos_capas(pasta_capas):
        if not capa_e_da_pessoa(nome, arquivo):
            continue
        nomes = [arquivo.stem]
        for pai in arquivo.parents:
            if pai == pasta_capas:
                break
            nomes.append(pai.name)
        mencoes = set()
        for item in nomes:
            mencoes |= proc.tamanhos_mencionados(item)
        if tamanho in mencoes and mencoes == {tamanho}:
            prioridade = 2
        elif tamanho in mencoes:
            prioridade = 1
        else:
            prioridade = 0
        candidato = (prioridade, proc.pontuacao_nome(nome, " ".join(nomes)), -len(arquivo.parts))
        if melhor is None or candidato > melhor[0]:
            melhor = (candidato, arquivo)
    return melhor[1] if melhor else None


def escolher_album_da_pessoa(indice, pessoa, tamanho, quantidade):
    """Como proc.escolher_album, mas só aceita a pasta com o nome inteiro da pessoa e prefere a do contrato."""

    nome = pessoa["nome"]
    contrato = pessoa.get("contrato") or ""
    melhor = None
    for remessa, pasta, mencoes in indice:
        if tamanho not in mencoes:
            continue
        try:
            relativo = pasta.relative_to(remessa)
        except ValueError:
            relativo = pasta
        if not capa_e_da_pessoa(nome, relativo):
            continue
        do_contrato = bool(contrato) and contrato in normalizar(str(relativo)).split()
        pasta_laminas, laminas, extras = proc.localizar_laminas(pasta, quantidade)
        qtd = len(laminas)
        chave = (
            do_contrato,
            proc.pontuacao_no_caminho(nome, pasta, remessa),
            len(mencoes) == 1,
            qtd == quantidade and len(extras) <= 2,
            qtd > 0,
            -abs(qtd - quantidade),
            -len(pasta.parts),
        )
        if melhor is None or chave > melhor[0]:
            melhor = (chave, remessa, pasta, pasta_laminas, laminas, extras)
    if not melhor:
        return None
    _, remessa, pasta, pasta_laminas, laminas, extras = melhor
    return {
        "remessa": remessa,
        "pasta_album": pasta,
        "pasta_laminas": pasta_laminas,
        "laminas": laminas,
        "extras": extras,
    }


def capa_do_album(albuns, nome, tamanho, escolha):
    pasta = escolha["pasta_album"] if escolha else None
    capa = procurar_capa_da_pessoa(albuns, nome, tamanho)
    if capa:
        print(f"   capa: {capa}")
        return capa
    if pasta:
        capa = proc.procurar_capa_antiga(pasta, tamanho)
    if not capa and escolha:
        extras = [
            arquivo for arquivo in (escolha.get("extras") or [])
            if proc.eh_lamina(arquivo) and proc.eh_nome_capa(arquivo.stem)
        ]
        capa = extras[0] if extras else None
    if capa:
        print(f"   capa: {capa}")
    return capa


def localizar_albuns(pessoa, albuns, indice):
    """Devolve os arquivos de cada tamanho, ou um resultado de falha."""

    if not pessoa["tamanhos"]:
        return None, "plano não identificado", "faltou arquivo"

    achados = []
    problemas = []
    for tamanho in pessoa["tamanhos"]:
        quantidade = proc.QUANTIDADE_LAMINAS[tamanho]
        escolha = escolher_album_da_pessoa(indice, pessoa, tamanho, quantidade)
        laminas = list(escolha["laminas"]) if escolha else []
        if escolha:
            print(f"   álbum {tamanho}: {escolha['pasta_album']}")
        capa = capa_do_album(albuns, pessoa["nome"], tamanho, escolha)
        if not escolha:
            problemas.append(f"{tamanho}: álbum não encontrado")
        elif len(laminas) < quantidade:
            problemas.append(f"{tamanho}: faltam lâminas ({len(laminas)}/{quantidade})")
        if not capa:
            problemas.append(f"{tamanho}: capa não encontrada")
        arquivos = list(laminas)
        if capa and capa not in arquivos:
            arquivos.append(capa)
        achados.append({
            "tamanho": tamanho,
            "laminas": len(laminas),
            "esperadas": quantidade,
            "capa": capa is not None,
            "arquivos": arquivos,
            "remessa": escolha["remessa"].name if escolha else "",
        })

    if problemas:
        return achados, "; ".join(problemas), "faltou arquivo"
    return achados, "", ""


def preparar_indice():
    albuns = proc.encontrar_albuns_colacao()
    if albuns is None:
        raise FileNotFoundError(f"Servidor não encontrado: {proc.SERVIDOR}")
    remessas = proc.encontrar_remessas(albuns)
    if not remessas:
        raise FileNotFoundError(f"Nenhuma remessa em {albuns}")
    return albuns, proc.indexar_pastas_album(remessas)


def texto_jpg(achados):
    return "; ".join(f"{item['tamanho']} {item['laminas']}/{item['esperadas']}" for item in achados)


def texto_capa(achados):
    if not achados:
        return "NAO"
    return "SIM" if all(item["capa"] for item in achados) else "NAO"


# ==========================================================
# MaisBook e Drive
# ==========================================================

def abrir_navegador(playwright):
    """O mesmo perfil do robô do MaisBook. O Chrome instalado passa no login do Google."""

    mb.PERFIL.mkdir(exist_ok=True)
    opcoes = dict(
        headless=False,
        no_viewport=True,
        ignore_default_args=["--enable-automation"],
        args=["--start-maximized", "--disable-blink-features=AutomationControlled"],
    )
    try:
        return playwright.chromium.launch_persistent_context(str(mb.PERFIL), channel="chrome", **opcoes)
    except Exception:
        return playwright.chromium.launch_persistent_context(
            str(mb.PERFIL), headless=False, no_viewport=True, args=["--start-maximized"])


def ir_clientes(pagina, sge):
    cfg = mb.carregar_config()
    url = f"{cfg['url_maisbook']}&acao=listar&search={quote(sge)}"
    base.ir(pagina, url, int(cfg.get("timeout_ms") or 20000))
    if pagina.locator("#user_login").count():
        if not mb.login_maisbook(pagina, int(cfg.get("timeout_ms") or 20000)):
            raise PararPessoa("erro no e-mail", "o MaisBook pediu login e não consegui entrar")
        base.ir(pagina, url, int(cfg.get("timeout_ms") or 20000))
    return url


def termo_busca(pessoa):
    return pessoa.get("codigo") or pessoa["sge"]


JS_LISTAR_CARTOES = r"""
() => {
  const norm = (s) => (s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ').trim().toUpperCase();
  const cabecalho = document.querySelector('.clientes-grid-header');
  const titulos = cabecalho ? [...cabecalho.children].map(c => norm(c.innerText)) : [];
  return [...document.querySelectorAll('.clientes-grid-row[data-cliente-id]')].map(linha => {
    const principal = linha.querySelector('.clientes-grid-main') || linha;
    const celulas = [...principal.children].map(c => (c.innerText || '').trim());
    const campo = (titulo, posicao) => {
      const i = titulos.indexOf(titulo);
      const valor = (i >= 0 && i < celulas.length ? celulas[i] : celulas[posicao]) || '';
      return valor.split('\n')[0].trim();
    };
    return {
      id: linha.getAttribute('data-cliente-id'),
      sge: campo('SGE', 0),
      ctt: campo('CTT', 1),
      nome: campo('NOME', 2),
      codigo: campo('CODIGO', 3),
    };
  });
}
"""


def descrever_cartao(cartao):
    return f"{cartao['nome']} (SGE {cartao['sge']}, CTT {cartao['ctt']}, código {cartao['codigo']})"


def cartao_e_da_pessoa(pessoa, cartao):
    if not capa_e_da_pessoa(pessoa["nome"], cartao["nome"]):
        return False
    if pessoa.get("contrato"):
        return so_contrato(cartao["ctt"]) == pessoa["contrato"]
    return cartao["sge"].strip() == pessoa["sge"]


def achar_cliente(pagina, pessoa):
    limite = time.time() + 20
    cartoes = []
    vazio_desde = None
    while time.time() < limite:
        cartoes = [c for c in pagina.evaluate(JS_LISTAR_CARTOES) if c.get("id")]
        if cartoes:
            break
        texto = base.texto_da_tela(pagina)
        if "mostrando 0" in texto or "nenhum cliente" in texto:
            vazio_desde = vazio_desde or time.time()
            if time.time() - vazio_desde >= 5:
                break
        else:
            vazio_desde = None
        time.sleep(0.5)
    if not cartoes:
        raise PararPessoa("não encontrado", f"MaisBook: nenhum cliente na pesquisa por {termo_busca(pessoa)}")

    certos = [c for c in cartoes if cartao_e_da_pessoa(pessoa, c)]
    outros = [c for c in cartoes if c not in certos]
    if len(cartoes) > 1:
        for cartao in cartoes:
            descricao = descrever_cartao(cartao)
            if descricao not in pessoa.setdefault("duplicados", []):
                pessoa["duplicados"].append(descricao)
    if not certos:
        achados = "; ".join(descrever_cartao(c) for c in outros[:3])
        raise PararPessoa(
            "não encontrado",
            f"MaisBook: nenhum cartão com o nome e o contrato {pessoa.get('contrato') or pessoa['sge']} — achei: {achados}",
        )
    if len(certos) > 1:
        raise PararPessoa(
            "duplicado",
            "duplicado — não enviado: " + "; ".join(descrever_cartao(c) for c in certos),
        )

    cliente_id = certos[0]["id"]
    linha = pagina.locator(f'.clientes-grid-row[data-cliente-id="{cliente_id}"]')
    linha.first.wait_for(state="visible", timeout=15000)
    return cliente_id, linha.first


def no_dialogo(dialog):
    try:
        if dialog.type == "beforeunload":
            dialog.accept()
        elif dialog.type == "alert":
            print(f"    aviso do navegador: {dialog.message}")
            dialog.accept()
        else:
            dialog.dismiss()
    except Exception:
        pass


def botao_novo(pagina):
    for seletor in ('button[guidedhelpid="new_menu_button"]', 'button[guidedhelpid="td_new_menu_button"]'):
        botoes = pagina.locator(seletor)
        for i in range(min(botoes.count(), 4)):
            botao = botoes.nth(i)
            if botao.is_visible():
                return botao
    return None


def esperar_drive_carregar(pagina):
    limite = time.time() + ESPERA_DRIVE_S
    while time.time() < limite:
        if botao_novo(pagina):
            return
        time.sleep(0.5)
    imagem = tirar_print(pagina, "drive_nao_carregou")
    detalhe = f"{pagina.url[:120]} | {imagem}"
    if "accounts.google" in pagina.url.lower():
        raise PararPessoa("erro no Drive", f"o Drive abriu na tela de login do Google — nada foi enviado | {detalhe}")
    raise PararPessoa("erro no Drive", f"o Drive não carregou a pasta do aluno | {detalhe}")


def pasta_atual(pagina):
    botoes = pagina.locator('[guidedhelpid="folder_path_button"]')
    textos = []
    for i in range(min(botoes.count(), 8)):
        try:
            texto = botoes.nth(i).inner_text(timeout=1000).strip()
        except Exception:
            texto = ""
        if texto:
            textos.append(texto)
    return textos[-1] if textos else ""


def esta_na_pasta(pagina, nome):
    alvo = normalizar(nome)
    if not alvo:
        return False
    titulo = normalizar(pagina.title())
    if titulo == alvo or titulo.startswith(alvo + " "):
        return True
    return normalizar(pasta_atual(pagina)) == alvo


def conferir_drive_da_pessoa(pagina, pessoa):
    limite = time.time() + 25
    while time.time() < limite:
        titulo = pagina.title()
        trilha = " ".join(
            pagina.locator('[guidedhelpid="folder_path_button"]').all_inner_texts()
        )
        if nomes_batem(pessoa["nome"], f"{titulo} {trilha}"):
            return
        time.sleep(1)
    raise PararPessoa("erro no Drive", "a pasta do Drive não é a desta pessoa — nada foi enviado")


JS_MARCAR_PASTA = r"""
(nome) => {
  const norm = (s) => (s || '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
  const alvo = norm(nome);
  const bate = (rotulo) => {
    const n = norm(String(rotulo).split('\n')[0]);
    if (!n || n === alvo) return n === alvo && !!n;
    if (!n.startsWith(alvo + ' ')) return false;
    return /pasta|folder/.test(n);
  };
  const els = [];
  const visitar = (raiz) => {
    if (!raiz || !raiz.querySelectorAll) return;
    for (const el of raiz.querySelectorAll('[data-robo-pasta]')) el.removeAttribute('data-robo-pasta');
    for (const el of raiz.querySelectorAll('[data-id], [role="row"], [role="option"], [role="gridcell"], [aria-label], [data-tooltip]')) {
      els.push(el);
    }
    for (const el of raiz.querySelectorAll('*')) if (el.shadowRoot) visitar(el.shadowRoot);
  };
  visitar(document);
  for (const el of els) {
    const rotulo = el.getAttribute('data-tooltip') || el.getAttribute('aria-label') || el.innerText || '';
    if (!bate(rotulo)) continue;
    let alvoEl = el;
    let pai = el;
    for (let i = 0; i < 8 && pai; i++) {
      if (pai.getAttribute && (pai.getAttribute('data-id') || pai.getAttribute('role') === 'row')) {
        alvoEl = pai;
        break;
      }
      pai = pai.parentElement;
    }
    alvoEl.setAttribute('data-robo-pasta', '1');
    try { alvoEl.scrollIntoView({block: 'center'}); } catch (e) {}
    return true;
  }
  return false;
}
"""

JS_ROLAR_LISTA = r"""
() => {
  let rolou = false;
  const visitar = (raiz) => {
    if (!raiz || !raiz.querySelectorAll) return;
    const todos = [raiz, ...raiz.querySelectorAll('*')];
    for (const el of todos) {
      if (el.shadowRoot) visitar(el.shadowRoot);
      try {
        if (el.scrollHeight > el.clientHeight + 40 && el.clientHeight > 120) {
          const antes = el.scrollTop;
          el.scrollTop = Math.min(el.scrollTop + Math.floor(el.clientHeight * 0.8), el.scrollHeight);
          if (el.scrollTop > antes) rolou = true;
        }
      } catch (e) {}
    }
  };
  visitar(document);
  return rolou;
}
"""

JS_FALTANDO = r"""
(nomes) => {
  const labels = [];
  const visitar = (raiz) => {
    if (!raiz || !raiz.querySelectorAll) return;
    for (const el of raiz.querySelectorAll('[aria-label], [data-tooltip]')) {
      labels.push(el.getAttribute('data-tooltip') || el.getAttribute('aria-label') || '');
    }
    for (const el of raiz.querySelectorAll('*')) if (el.shadowRoot) visitar(el.shadowRoot);
  };
  visitar(document);
  const texto = labels.join('\n').toLowerCase();
  return nomes.filter(nome => !texto.includes(String(nome).toLowerCase()));
}
"""

JS_STATUS_UPLOAD = r"""
() => {
  const norm = (s) => (s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
  const juntar = (no) => {
    if (!no) return '';
    let texto = '';
    try { texto = no.innerText || ''; } catch (e) {}
    if (!no.querySelectorAll) return texto;
    for (const el of no.querySelectorAll('*')) {
      if (el.shadowRoot) texto += ' ' + juntar(el.shadowRoot);
    }
    return texto;
  };
  const texto = norm(juntar(document.body));
  const numero = texto.match(/(\d+)\s+uploads?\s+conclu/);
  return {
    fazendo: texto.includes('fazendo upload') || texto.includes('uploading'),
    concluido: texto.includes('upload conclu') || texto.includes('uploads conclu') || texto.includes('upload complete'),
    falha: texto.includes('nao foi enviado') || texto.includes('falha no upload') || texto.includes('upload failed'),
    quantidade: numero ? Number(numero[1]) : 0,
  };
}
"""


def clicar_item_menu(pagina, texto):
    itens = pagina.get_by_role("menuitem").filter(has_text=texto)
    limite = time.time() + 8
    while time.time() < limite:
        for i in range(min(itens.count(), 8)):
            item = itens.nth(i)
            if item.is_visible():
                item.click()
                return
        time.sleep(0.3)
    raise PararPessoa("erro no Drive", f"não achei '{texto}' no menu Novo")


def criar_pasta(pagina, nome):
    pagina.keyboard.press("Escape")
    botao = botao_novo(pagina)
    if botao is None:
        raise PararPessoa("erro no Drive", "não achei o botão Novo do Drive")
    botao.click()
    clicar_item_menu(pagina, "Nova pasta")
    dialog = pagina.get_by_role("dialog").last
    dialog.wait_for(state="visible", timeout=10000)
    campo = dialog.locator("input").first
    campo.fill(nome)
    clicou = False
    for rotulo in ("Criar", "Create"):
        botao_criar = dialog.get_by_role("button", name=rotulo, exact=True)
        if botao_criar.count() and botao_criar.first.is_visible():
            botao_criar.first.click()
            clicou = True
            break
    if not clicou:
        campo.press("Enter")
    try:
        dialog.wait_for(state="hidden", timeout=8000)
    except Exception:
        try:
            texto = normalizar(dialog.inner_text(timeout=1000))
        except Exception:
            texto = ""
        if "ja existe" in texto or "already exists" in texto:
            pagina.keyboard.press("Escape")
            return False
        pagina.keyboard.press("Escape")
    return True


def esperar_entrar(pagina, nome, segundos):
    limite = time.time() + segundos
    while time.time() < limite:
        if esta_na_pasta(pagina, nome):
            pagina.wait_for_timeout(1200)
            return True
        time.sleep(0.4)
    return False


def abrir_pasta(pagina, nome):
    if esta_na_pasta(pagina, nome):
        return
    limite = time.time() + 25
    while time.time() < limite:
        if not pagina.evaluate(JS_MARCAR_PASTA, nome):
            time.sleep(0.6)
            continue
        alvo = pagina.locator('[data-robo-pasta="1"]').last
        id_pasta = alvo.evaluate(
            """(el) => el.getAttribute('data-id')
                || (el.closest('[data-id]') && el.closest('[data-id]').getAttribute('data-id'))
                || (el.querySelector('[data-id]') && el.querySelector('[data-id]').getAttribute('data-id'))
                || ''"""
        )
        if id_pasta:
            pagina.goto(f"https://drive.google.com/drive/folders/{id_pasta}", wait_until="domcontentloaded", timeout=60000)
            if esperar_entrar(pagina, nome, 15):
                return
        try:
            alvo.click(timeout=5000)
            pagina.keyboard.press("Enter")
        except Exception:
            try:
                alvo.dblclick(timeout=5000)
            except Exception:
                pass
        if esperar_entrar(pagina, nome, 8):
            return
        time.sleep(0.4)
    raise PararPessoa("erro no Drive", f"não consegui abrir a pasta {nome}")


def garantir_pasta(pagina, nome):
    if esta_na_pasta(pagina, nome):
        print(f"    já dentro de {nome}")
        return
    limite = time.time() + 18
    lista_parada = 0
    while time.time() < limite:
        if pagina.evaluate(JS_MARCAR_PASTA, nome):
            print(f"    pasta {nome} já existe — abrindo")
            abrir_pasta(pagina, nome)
            return
        if pagina.evaluate(JS_ROLAR_LISTA):
            lista_parada = 0
        else:
            lista_parada += 1
            if lista_parada >= 3:
                break
        time.sleep(0.5)
    print(f"    criando pasta {nome}")
    if not criar_pasta(pagina, nome):
        print(f"    pasta {nome} já existe — abrindo")
        abrir_pasta(pagina, nome)
        return
    url_antes = pagina.url
    pagina.wait_for_timeout(800)
    pagina.keyboard.press("Enter")
    if esperar_entrar(pagina, nome, 8):
        return
    if pagina.url != url_antes and "drive.google.com" in pagina.url:
        try:
            pagina.go_back(timeout=15000)
        except Exception:
            pass
        pagina.wait_for_timeout(800)
    abrir_pasta(pagina, nome)


def tratar_conflito(pagina):
    dialogos = pagina.get_by_role("dialog")
    for i in range(min(dialogos.count(), 3)):
        dialog = dialogos.nth(i)
        try:
            if not dialog.is_visible():
                continue
            texto = normalizar(dialog.inner_text(timeout=1000))
        except Exception:
            continue
        if "ja existe" not in texto and "already exists" not in texto:
            continue
        caixa = dialog.get_by_role("checkbox")
        if caixa.count() and caixa.first.is_visible():
            try:
                if not caixa.first.is_checked():
                    caixa.first.check()
            except Exception:
                pass
        for rotulo in ("Manter os dois arquivos", "Manter ambos", "Keep both", "Substituir", "Replace"):
            botao = dialog.get_by_role("button", name=rotulo)
            if botao.count() and botao.first.is_visible():
                botao.first.click()
                return


def arquivos_faltando(pagina, arquivos):
    nomes = [arquivo.name for arquivo in arquivos]
    try:
        faltando = pagina.evaluate(JS_FALTANDO, nomes)
    except Exception:
        return list(arquivos)
    por_nome = {arquivo.name: arquivo for arquivo in arquivos}
    return [por_nome[nome] for nome in faltando if nome in por_nome]


def enviar_arquivos(pagina, arquivos):
    if not arquivos:
        return
    pagina.keyboard.press("Escape")
    botao = botao_novo(pagina)
    if botao is None:
        raise PararPessoa("erro no Drive", "não achei o botão Novo para enviar os arquivos")
    botao.click()
    with pagina.expect_file_chooser(timeout=20000) as escolha:
        clicar_item_menu(pagina, "Upload de arquivo")
    escolha.value.set_files([str(arquivo) for arquivo in arquivos])


def esperar_carregamento(pagina, arquivos, enviados):
    nomes = [arquivo.name for arquivo in arquivos]
    limite = time.time() + ESPERA_UPLOAD_S
    viu_envio = False
    while time.time() < limite:
        tratar_conflito(pagina)
        try:
            status = pagina.evaluate(JS_STATUS_UPLOAD)
            faltando = pagina.evaluate(JS_FALTANDO, nomes)
        except Exception:
            time.sleep(2)
            continue
        if status.get("fazendo"):
            viu_envio = True
        if status.get("falha") and faltando:
            raise PararPessoa("erro no Drive", "o Drive não enviou todos os arquivos")
        if not faltando and not status.get("fazendo"):
            pagina.wait_for_timeout(3000)
            return
        quantidade = status.get("quantidade") or 0
        if viu_envio and status.get("concluido") and not status.get("fazendo") and quantidade >= len(enviados):
            pagina.wait_for_timeout(3000)
            return
        time.sleep(2)
    raise PararPessoa("erro no Drive", "as fotos não terminaram de carregar no Drive")


def colocar_arquivos(pagina, arquivos):
    faltam = arquivos_faltando(pagina, arquivos)
    if faltam:
        print(f"    enviando {len(faltam)} arquivo(s)")
        enviar_arquivos(pagina, faltam)
        esperar_carregamento(pagina, arquivos, faltam)
    else:
        print("    os arquivos já estavam na pasta")
        pagina.wait_for_timeout(2000)


def voltar_para_pasta(pagina, nome):
    if esta_na_pasta(pagina, nome):
        return
    botoes = pagina.locator('[guidedhelpid="folder_path_button"]').filter(has_text=nome)
    if botoes.count():
        botoes.first.click()
        if esperar_entrar(pagina, nome, 12):
            return
    if "drive.google.com" in pagina.url:
        pagina.go_back(timeout=15000)
        if esperar_entrar(pagina, nome, 12):
            return
    raise PararPessoa("erro no Drive", f"não voltei para a pasta {nome}")


def abrir_drive(contexto, linha):
    link = linha.locator('a[title="Abrir pasta no Drive"]')
    if not link.count():
        raise PararPessoa("erro no Drive", "não achei o botão do Drive neste cliente")
    endereco = (link.first.get_attribute("href") or "").strip()
    if "drive.google.com" in endereco:
        drive = contexto.new_page()
        drive.goto(endereco, wait_until="domcontentloaded", timeout=60000)
        return drive, True
    link.first.scroll_into_view_if_needed()
    antes = list(contexto.pages)
    link.first.click()
    limite = time.time() + 8
    while time.time() < limite:
        novas = [pagina for pagina in contexto.pages if pagina not in antes]
        if novas:
            novas[-1].wait_for_load_state("domcontentloaded")
            return novas[-1], True
        time.sleep(0.3)
    return contexto.pages[-1], False


def link_da_pasta(url):
    achado = re.search(r"/folders/([a-zA-Z0-9_-]+)", url or "")
    if not achado:
        return ""
    return "https://drive.google.com/drive/folders/" + achado.group(1)


def subir_para_o_drive(contexto, pagina_mb, pessoa, achados):
    _, linha = achar_cliente(pagina_mb, pessoa)
    drive, aba_nova = abrir_drive(contexto, linha)
    drive.on("dialog", no_dialogo)
    try:
        esperar_drive_carregar(drive)
        conferir_drive_da_pessoa(drive, pessoa)
        garantir_pasta(drive, PASTA_RAIZ)
        if len(achados) >= 2:
            for item in achados:
                nome_pasta = PASTA_TAMANHO.get(item["tamanho"], item["tamanho"].lower())
                print(f"    pasta {nome_pasta}")
                garantir_pasta(drive, nome_pasta)
                colocar_arquivos(drive, item["arquivos"])
                voltar_para_pasta(drive, PASTA_RAIZ)
        else:
            colocar_arquivos(drive, achados[0]["arquivos"])
        if not esta_na_pasta(drive, PASTA_RAIZ):
            voltar_para_pasta(drive, PASTA_RAIZ)
        link = link_da_pasta(drive.url)
        if not link:
            raise PararPessoa("erro no Drive", "não consegui o link da pasta ALBUM APROVAÇAO")
        print(f"    link: {link}")
        return link
    finally:
        if aba_nova:
            try:
                drive.close()
            except Exception:
                pass
        else:
            try:
                pagina_mb.go_back(timeout=15000)
            except Exception:
                pass
        try:
            pagina_mb.bring_to_front()
        except Exception:
            pass


def abrir_edicao(pagina, pessoa):
    cliente_id, _ = achar_cliente(pagina, pessoa)
    pagina.locator(f'button[title="Editar"][data-cliente-id="{cliente_id}"]').click()
    modal = pagina.locator("#modal-cliente")
    modal.wait_for(state="visible", timeout=15000)
    campo_sge = pagina.locator("#cliente_codigo_interno")
    limite = time.time() + 15
    while time.time() < limite and campo_sge.input_value().strip() != pessoa["sge"]:
        time.sleep(0.3)
    if campo_sge.input_value().strip() != pessoa["sge"]:
        modal.get_by_role("button", name="Cancelar", exact=True).first.click()
        raise PararPessoa("erro no e-mail", "o Editar abriu outro cliente — link não colado")
    return modal


def campo_album_pronto(pagina):
    campo = pagina.locator("#cliente_album_link")
    campo.wait_for(state="visible", timeout=10000)
    anterior, desde = None, time.time()
    limite = time.time() + 12
    while time.time() < limite:
        atual = campo.input_value()
        if atual != anterior:
            anterior, desde = atual, time.time()
        elif time.time() - desde >= 1.5:
            return campo
        time.sleep(0.3)
    return campo


def edicao_aberta(pagina):
    modal = pagina.locator("#modal-cliente.active")
    try:
        return modal.count() > 0 and modal.first.is_visible()
    except Exception:
        return False


def fechar_edicao(pagina):
    if not edicao_aberta(pagina):
        return
    pagina.locator("#modal-cliente").get_by_role("button", name="Cancelar", exact=True).first.click()
    pagina.locator("#modal-cliente.active").wait_for(state="hidden", timeout=10000)


def esperar_salvar(pagina):
    inicio = time.time()
    while time.time() - inicio < 25:
        if not edicao_aberta(pagina):
            return
        botao = pagina.locator("#modal-cliente button[type='submit']")
        try:
            texto = normalizar(botao.first.inner_text(timeout=1000)) if botao.count() else ""
        except Exception:
            texto = ""
        if "salvando" not in texto and time.time() - inicio > 2:
            fechar_edicao(pagina)
            return
        time.sleep(0.4)
    raise PararPessoa("erro no e-mail", "o MaisBook não terminou de salvar o Álbum pronto")


def colar_link_album(pagina, pessoa, link):
    modal = abrir_edicao(pagina, pessoa)
    campo = campo_album_pronto(pagina)
    campo.fill(link)
    if link_da_pasta(campo.input_value()) != link:
        fechar_edicao(pagina)
        raise PararPessoa("erro no e-mail", "não consegui colar o link em Álbum pronto")
    modal.get_by_role("button", name="Salvar", exact=True).click()
    esperar_salvar(pagina)

    abrir_edicao(pagina, pessoa)
    gravado = pagina.locator("#cliente_album_link").input_value()
    fechar_edicao(pagina)
    if link_da_pasta(gravado) != link:
        raise PararPessoa("erro no e-mail", "salvei, mas o link não ficou em Álbum pronto")
    print("    link salvo em Álbum pronto")
    return link


def enviar_email(pagina, pessoa):
    cliente_id, _ = achar_cliente(pagina, pessoa)
    pagina.locator(f'button[title="Enviar e-mail"][data-cliente-id="{cliente_id}"]').click()
    modal = pagina.locator("#modal-notificacao-cliente")
    modal.wait_for(state="visible", timeout=15000)
    info = modal.locator("#notificacao_cliente_info").inner_text()
    if pessoa["nome"] and not nomes_batem(pessoa["nome"], info):
        modal.locator("button.modal-close").first.click()
        raise PararPessoa("erro no e-mail", "o pop-up abriu outro cliente — e-mail não enviado")

    pagina.locator("#notificacao_template").evaluate(
        """(el, valor) => {
          el.value = valor;
          el.dispatchEvent(new Event('change', { bubbles: true }));
          if (window.jQuery) window.jQuery(el).val(valor).trigger('change');
        }""",
        TEMPLATE_VALOR,
    )
    valor = pagina.locator("#notificacao_template").input_value()
    assunto = pagina.locator("#notificacao_assunto").input_value().strip()
    if valor != TEMPLATE_VALOR or not assunto:
        modal.locator("button.modal-close").first.click()
        raise PararPessoa("erro no e-mail", "não consegui selecionar o template Aprovação de álbum")

    pagina.locator("#btn-enviar-notificacao").click()
    resultado = pagina.locator("#notificacao_result")
    resultado.wait_for(state="visible", timeout=45000)
    info_resultado = resultado.evaluate(
        """el => ({ texto: el.innerText || '', fundo: getComputedStyle(el).backgroundColor || '' })"""
    )
    texto = normalizar(info_resultado.get("texto"))
    verde = "209, 250, 229" in info_resultado.get("fundo", "")
    if "erro" in texto and "sucesso" not in texto:
        raise PararPessoa("erro no e-mail", info_resultado.get("texto") or "o MaisBook recusou o envio")
    if "sucesso" in texto or "enviado" in texto or verde:
        return info_resultado.get("texto") or "E-mail enviado com sucesso."
    raise PararPessoa("erro no e-mail", "o MaisBook não confirmou o envio")


# ==========================================================
# Atualização da Coluna K (Aguardando Aprovação) + Baixa SGE
# ==========================================================

ORDEM_COLUNA_K = [
    "Sem álbum",
    "AGUARDANDO FOTO",
    "EM DIAGRAMAÇÃO",
    "DIAGRAMADO",
    "Aguardando Aprovação",
    "Com Modificações",
    "Aprovado",
    "IMPRESSÃO",
    "Finalizado",
]


def indice_etapa_k(valor):
    alvo = normalizar(valor)
    for idx, etapa in enumerate(ORDEM_COLUNA_K):
        if normalizar(etapa) == alvo:
            return idx
    return -1


def atualizar_status_planilha(contexto, pessoa, novo_status="Aguardando Aprovação"):
    """
    Atualiza a Coluna K (STATUS) da pessoa para 'Aguardando Aprovação' na planilha do Google
    (e na cópia local), respeitando a regra de que a Coluna K só anda para baixo.
    """
    from openpyxl.utils import get_column_letter

    aba = pessoa.get("aba") or "GERAL ENVIOS"
    linha_num = int(pessoa.get("linha") or 0)
    col_status = int(pessoa.get("col_status") or 11)
    if linha_num < 1:
        return "linha inválida"

    # 1. Atualiza na cópia local .xlsx se existir
    try:
        if PLANILHA.is_file():
            wb = proc.load_workbook(PLANILHA)
            if aba in wb.sheetnames:
                ws = wb[aba]
                atual_local = ws.cell(row=linha_num, column=col_status).value
                idx_atual = indice_etapa_k(atual_local)
                idx_novo = indice_etapa_k(novo_status)
                if idx_atual < idx_novo and idx_atual not in (7, 8):
                    ws.cell(row=linha_num, column=col_status, value=novo_status)
                    wb.save(PLANILHA)
            wb.close()
    except Exception:
        pass

    # 2. Atualiza no Google Sheets usando o navegador já autenticado na conta Google
    letra_col = get_column_letter(col_status)
    ref_celula = f"{letra_col}{linha_num}"
    plan_pag = None
    criou_aba = False
    try:
        for pag in contexto.pages:
            if "/spreadsheets/d/" in (pag.url or ""):
                plan_pag = pag
                break
        if plan_pag is None:
            plan_pag = contexto.new_page()
            criou_aba = True
            plan_pag.goto(PLANILHA_LINK, wait_until="domcontentloaded", timeout=60000)

        caixa = plan_pag.locator("#t-name-box")
        caixa.first.wait_for(state="visible", timeout=25000)
        caixa.first.click()
        caixa.first.fill("'" + aba.replace("'", "''") + "'!" + ref_celula)
        caixa.first.press("Enter")
        plan_pag.wait_for_timeout(500)

        try:
            atual_barra = (
                plan_pag.locator("#t-formula-bar-input")
                .first.inner_text(timeout=3000)
                .replace("\ufeff", "")
                .replace("\u200b", "")
                .strip()
            )
        except Exception:
            atual_barra = ""

        idx_atual = indice_etapa_k(atual_barra)
        idx_novo = indice_etapa_k(novo_status)
        if idx_atual in (7, 8) or (idx_atual >= idx_novo and idx_atual != -1):
            print(f"    planilha: Coluna K ({ref_celula}) já está em '{atual_barra}' (não retrocede)")
            return f"Coluna K mantida em {atual_barra}"

        plan_pag.keyboard.type(novo_status, delay=10)
        plan_pag.keyboard.press("Delete")
        plan_pag.keyboard.press("Enter")
        plan_pag.wait_for_timeout(700)
        print(f"    planilha: Coluna K ({ref_celula}) atualizada para '{novo_status}'")
        return f"Coluna K ({ref_celula}) -> {novo_status}"
    except Exception as erro:
        aviso = f"não consegui editar a célula {ref_celula} direto no Google Sheets ({base.erro_curto(erro)})"
        print(f"    aviso planilha: {aviso}")
        return aviso
    finally:
        if criou_aba and plan_pag is not None:
            try:
                plan_pag.close()
            except Exception:
                pass


def dar_baixa_sge_enviado(playwright, pessoa, estado_sge):
    """
    Aciona o robô do SGE (baixa_sge/BAIXA.PY) para registrar 'Enviado para Aprovação'
    para o formando assim que o álbum é enviado no MaisBook.
    """
    try:
        if estado_sge.get("robo") is None:
            sge_cfg = base.carregar_config()
            sge_cfg["etapa"] = "Enviado para Aprovação"
            sge_contexto, sge_pagina = base.abrir_navegador(playwright)
            robo_sge = base.Robo(sge_contexto, sge_pagina, sge_cfg, True, False)
            sge_pagina.on("dialog", robo_sge._dialogo)
            contas_sge = base.ContasSGE()
            base.ir(sge_pagina, sge_cfg["url_clientes"], robo_sge.timeout)
            base.garantir_login(robo_sge, contas_sge, PARAR)
            estado_sge["robo"] = robo_sge
            estado_sge["contas"] = contas_sge
            estado_sge["contexto"] = sge_contexto
            estado_sge["cfg"] = sge_cfg

        robo_sge = estado_sge["robo"]
        contas_sge = estado_sge["contas"]
        sge_cfg = estado_sge["cfg"]
        sge_cfg["etapa"] = "Enviado para Aprovação"

        aluno_sge = {
            "sge": pessoa["sge"],
            "nome_planilha": pessoa["nome"],
            "status_coluna_k": "Aguardando Aprovação",
            "etapa_sge": "Enviado para Aprovação",
        }
        status_sge, nome_sge, det_sge, _ = base.processar_com_contas(robo_sge, contas_sge, aluno_sge)
        if status_sge in ("CONCLUIDO", "JA ESTAVA"):
            feitos_sge = base.carregar_feitos()
            feitos_sge[pessoa["sge"]] = {
                "nome": nome_sge or pessoa["nome"],
                "etapa": "Enviado para Aprovação",
                "status_coluna_k": "Aguardando Aprovação",
                "data": f"{datetime.now():%Y-%m-%d %H:%M:%S}",
                "status": status_sge,
            }
            base.salvar_feitos(feitos_sge)
        resumo_sge = f"SGE: {status_sge} (Enviado para Aprovação)" + (f" — {det_sge}" if det_sge else "")
        print(f"    {resumo_sge}")
        return resumo_sge
    except Exception as erro:
        aviso = f"SGE erro: {base.erro_curto(erro)}"
        print(f"    {aviso}")
        return aviso


# ==========================================================
# Execução e Modo Vigia Contínuo
# ==========================================================

def linha_relatorio(pessoa, achados, resultado, mensagem):
    achados = achados or []
    avisos = list(pessoa.get("avisos") or [])
    if pessoa.get("duplicados"):
        avisos.append("pesquisa trouxe mais de um cartão: " + "; ".join(pessoa["duplicados"]))
    return {
        "NOME": pessoa["nome"],
        "SGE": pessoa["sge"],
        "CONTRATO": pessoa.get("contrato", ""),
        "CODIGO": pessoa.get("codigo", ""),
        "TAMANHOS": ", ".join(pessoa["tamanhos"]),
        "JPG": texto_jpg(achados),
        "CAPA": texto_capa(achados),
        "RESULTADO": resultado,
        "MENSAGEM": mensagem,
        "AVISO": " | ".join(avisos),
    }


def pediu_parar():
    return PARAR.exists()


def analisar_argumentos():
    parser = argparse.ArgumentParser(description="Envia o álbum diagramado para aprovação no MaisBook")
    parser.add_argument("--sge", help="processa só este código")
    parser.add_argument("--limite", type=int)
    parser.add_argument("--simular", action="store_true", help="procura e abre o cliente, sem pasta, sem arquivo e sem e-mail")
    parser.add_argument("--continuar", action="store_true", help="retoma quem ainda não foi enviado")
    parser.add_argument("--listar", action="store_true", help="só lista quem está com STATUS DIAGRAMADO")
    parser.add_argument("--vigia", action="store_true", help="monitora continuamente a planilha e envia automaticamente quem entrar como DIAGRAMADO")
    parser.add_argument("--intervalo", type=int, default=60, help="segundos entre cada verificação no modo --vigia (padrão: 60)")
    parser.add_argument("--sem-planilha", action="store_true", help="não atualiza a Coluna K para Aguardando Aprovação após enviar")
    parser.add_argument("--sem-sge", action="store_true", help="não dá baixa automática no SGE (Enviado para Aprovação) após enviar")
    return parser.parse_args()


def executar_rodada(args, em_vigia=False):
    feitos = carregar_json(FEITOS)
    pessoas = ler_diagramados()
    if args.sge:
        codigo = limpar_sge(args.sge)
        pessoas = [pessoa for pessoa in pessoas if pessoa["sge"] == codigo]
        if not pessoas:
            print(f"O código {args.sge} não está com STATUS DIAGRAMADO.")
            return 0
    if args.listar:
        for pessoa in pessoas:
            marca = "JÁ ENVIADO" if pessoa["sge"] in feitos else "PENDENTE"
            print(f"linha {pessoa['linha']}: SGE {pessoa['sge']} — {pessoa['nome']} | "
                  f"código MB {pessoa.get('codigo') or '-'} | "
                  f"{', '.join(pessoa['tamanhos']) or '(plano não identificado)'} | {marca}")
        print(f"\nTotal com STATUS DIAGRAMADO: {len(pessoas)}")
        return len(pessoas)

    ja_enviadas = [pessoa for pessoa in pessoas if pessoa["sge"] in feitos]
    pessoas = [pessoa for pessoa in pessoas if pessoa["sge"] not in feitos]
    if ja_enviadas:
        print(f"Já enviadas, sem mandar de novo ({len(ja_enviadas)}): {', '.join(pessoa['sge'] for pessoa in ja_enviadas)}")
    if args.continuar:
        print(f"Continuando com {len(pessoas)} pessoa(s).")
    if args.limite:
        pessoas = pessoas[: args.limite]

    print("=" * 70)
    print(
        "ENVIO DE ÁLBUM PARA APROVAÇÃO —",
        "SIMULAÇÃO" if args.simular else ("MODO VIGIA CONTÍNUO" if em_vigia else "ENVIO REAL"),
    )
    print(f"Pessoas pendentes com STATUS DIAGRAMADO: {len(pessoas)}")
    print("=" * 70)
    if not pessoas:
        print("Nenhuma pessoa pendente para processar agora.")
        return 0

    if not em_vigia:
        PARAR.unlink(missing_ok=True)
    print("\n" + "=" * 70)
    print("FASE 1 — SERVIDOR (capa e álbum)")
    print("=" * 70)
    albuns, indice = preparar_indice()
    prontos = []
    fora = []
    linhas = []
    for pessoa in pessoas:
        if pediu_parar():
            print("\nPARADO A PEDIDO antes de procurar o restante.")
            break
        print(f"\n{pessoa['nome']} — SGE {pessoa['sge']} — {', '.join(pessoa['tamanhos']) or pessoa['planos']}")
        achados, mensagem, resultado = localizar_albuns(pessoa, albuns, indice)
        if resultado:
            for parte in mensagem.split("; "):
                if parte.strip():
                    print(f"NÃO ENCONTREI: {pessoa['nome']} (SGE {pessoa['sge']}) — {parte.strip()}")
            print(f"    {resultado} — {mensagem}")
            linhas.append(linha_relatorio(pessoa, achados, resultado, mensagem))
            salvar_relatorio(linhas)
            fora.append((pessoa, mensagem))
            continue
        print("    " + texto_jpg(achados) + " | capa " + texto_capa(achados))
        prontos.append((pessoa, achados))

    print("\n" + "-" * 70)
    print(f"VÃO PARA O MAISBOOK ({len(prontos)}):")
    for pessoa, _ in prontos:
        print(f"    {pessoa['nome']} — SGE {pessoa['sge']} — {', '.join(pessoa['tamanhos'])}")
    print(f"FICARAM DE FORA ({len(fora)}):")
    for pessoa, motivo in fora:
        print(f"    {pessoa['nome']} — SGE {pessoa['sge']} — {motivo}")
    print("-" * 70)

    if not prontos:
        print("\nNinguém com capa e álbum completos — o MaisBook não foi aberto.")
        print(f"RELATORIO: {RELATORIO}")
        return 0

    from playwright.sync_api import sync_playwright

    print("\n" + "=" * 70)
    print("FASE 2 — MAISBOOK + PLANILHA (Coluna K) + SGE (um aluno por vez)")
    print("=" * 70)
    enviados_rodada = 0
    with sync_playwright() as p:
        contexto = abrir_navegador(p)
        pagina = contexto.pages[0] if contexto.pages else contexto.new_page()
        pagina.on("dialog", no_dialogo)
        estado_sge = {"robo": None, "contas": None, "contexto": None, "cfg": None}
        try:
            for numero, (pessoa, achados) in enumerate(prontos, start=1):
                if pediu_parar():
                    print(f"\nPARADO A PEDIDO antes de {pessoa['nome']}.")
                    break
                print(
                    f"\n[{numero}/{len(prontos)}] MaisBook — {pessoa['nome']} ({pessoa['sge']}) "
                    f"| contrato {pessoa.get('contrato') or '-'} | pesquisa {termo_busca(pessoa)}"
                )
                try:
                    ir_clientes(pagina, termo_busca(pessoa))
                    if args.simular:
                        achar_cliente(pagina, pessoa)
                        print("    simulado — cliente encontrado, sem pasta e sem e-mail")
                        linhas.append(linha_relatorio(pessoa, achados, "simulado", "cliente encontrado no MaisBook"))
                    else:
                        link = subir_para_o_drive(contexto, pagina, pessoa, achados)
                        ir_clientes(pagina, termo_busca(pessoa))
                        colar_link_album(pagina, pessoa, link)
                        ir_clientes(pagina, termo_busca(pessoa))
                        confirmacao = enviar_email(pagina, pessoa)
                        print(f"    enviado — {confirmacao}")

                        extras_msg = [confirmacao]
                        if not getattr(args, "sem_planilha", False):
                            info_plan = atualizar_status_planilha(contexto, pessoa, "Aguardando Aprovação")
                            if info_plan:
                                extras_msg.append(info_plan)
                            try:
                                pagina.bring_to_front()
                            except Exception:
                                pass

                        if not getattr(args, "sem_sge", False):
                            info_sge = dar_baixa_sge_enviado(p, pessoa, estado_sge)
                            if info_sge:
                                extras_msg.append(info_sge)
                            try:
                                pagina.bring_to_front()
                            except Exception:
                                pass

                        feitos[pessoa["sge"]] = {
                            "nome": pessoa["nome"],
                            "tamanhos": pessoa["tamanhos"],
                            "status_coluna_k": "Aguardando Aprovação",
                            "etapa_sge": "Enviado para Aprovação",
                            "data": f"{datetime.now():%Y-%m-%d %H:%M:%S}",
                        }
                        salvar_json(FEITOS, feitos)
                        enviados_rodada += 1
                        linhas.append(linha_relatorio(pessoa, achados, "enviado", " | ".join(extras_msg)))
                except PararPessoa as parada:
                    imagem = tirar_print(pagina, f"erro_{pessoa['sge']}")
                    detalhe = parada.detalhe + (f" | {imagem}" if imagem else "")
                    if parada.resultado == "duplicado":
                        print(f"DUPLICADO: {pessoa['nome']} (SGE {pessoa['sge']}) — {parada.detalhe}")
                    elif parada.resultado == "não encontrado" or "não encontr" in parada.detalhe.lower():
                        print(f"NÃO ENCONTREI: {pessoa['nome']} (SGE {pessoa['sge']}) — {parada.detalhe}")
                    print(f"    {parada.resultado} — {detalhe}")
                    linhas.append(linha_relatorio(pessoa, achados, parada.resultado, detalhe))
                except Exception as erro:
                    if "has been closed" in str(erro):
                        print("    o navegador foi fechado — parando")
                        linhas.append(linha_relatorio(pessoa, achados, "erro no Drive", "navegador fechado"))
                        break
                    imagem = tirar_print(pagina, f"erro_{pessoa['sge']}")
                    detalhe = f"{base.erro_curto(erro)}" + (f" | {imagem}" if imagem else "")
                    print(f"    erro — {detalhe}")
                    linhas.append(linha_relatorio(pessoa, achados, "erro no Drive", detalhe))
                salvar_relatorio(linhas)
        finally:
            if estado_sge.get("contexto") is not None:
                try:
                    estado_sge["contexto"].close()
                except Exception:
                    pass
            try:
                contexto.close()
            except Exception:
                pass

    duplicadas = [pessoa for pessoa, _ in prontos if pessoa.get("duplicados")]
    if duplicadas:
        print("\n" + "-" * 70)
        print(f"PESSOAS DUPLICADAS NO MAISBOOK ({len(duplicadas)}):")
        for pessoa in duplicadas:
            print(
                f"DUPLICADO: {pessoa['nome']} (SGE {pessoa['sge']}) — pesquisa {termo_busca(pessoa)} "
                f"trouxe {len(pessoa['duplicados'])} cartões"
            )
            for cartao in pessoa["duplicados"]:
                print(f"    {cartao}")
        print("-" * 70)

    print("\n" + "=" * 70)
    print(f"RELATORIO: {RELATORIO}")
    for resultado in ("enviado", "simulado", "faltou arquivo", "erro no Drive", "erro no e-mail", "não encontrado", "duplicado"):
        qtd = sum(1 for linha in linhas if linha["RESULTADO"] == resultado)
        if qtd:
            print(f"{resultado}: {qtd}")
    print("=" * 70)
    return enviados_rodada


def main():
    args = analisar_argumentos()
    if args.listar or not args.vigia:
        executar_rodada(args, em_vigia=False)
        return

    intervalo = max(10, int(args.intervalo or 60))
    PARAR.unlink(missing_ok=True)
    print("=" * 70)
    print(f"VIGIA CONTÍNUO MAISBOOK INICIADO (intervalo: {intervalo}s)")
    print("Monitorando STATUS = DIAGRAMADO -> Drive + MaisBook + Coluna K (Aguardando Aprovação) + SGE (Enviado para Aprovação)")
    print("=" * 70)
    ciclo = 0
    try:
        while not pediu_parar():
            ciclo += 1
            print(f"\n[CICLO #{ciclo} — {datetime.now():%d/%m/%Y %H:%M:%S}] Verificando planilha e servidor...")
            try:
                executar_rodada(args, em_vigia=True)
            except Exception as erro:
                print(f"[AVISO VIGIA] Erro no ciclo #{ciclo}: {erro}")
            if pediu_parar():
                break
            print(f"[VIGIA] Aguardando {intervalo}s para a próxima verificação (use Parar para encerrar)...")
            for _ in range(intervalo):
                if pediu_parar():
                    break
                time.sleep(1)
    except KeyboardInterrupt:
        print("\n[VIGIA] Interrompido pelo usuário.")
    finally:
        PARAR.unlink(missing_ok=True)
        print("\n[VIGIA] Monitoramento encerrado.")


if __name__ == "__main__":
    try:
        main()
    except (FileNotFoundError, ValueError) as erro:
        print(f"\nERRO: {erro}")
        sys.exit(1)

