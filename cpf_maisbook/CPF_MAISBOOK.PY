# -*- coding: utf-8 -*-
"""
ROBÔ MAISBOOK — RECADASTRAR O CPF DOS ALUNOS

Para cada aluno da planilha do Google Drive (aba GERAL ENVIOS, coluna SGE):
  1. MaisBook (wp-admin > Clientes): pesquisa pelo CÓDIGO DO MAISBOOK (coluna I da planilha;
     é o Cód. Acesso; se a célula estiver vazia, usa o SGE) e abre "Editar" do cartão com esse código.
     Se o cadastro já tem CPF válido, pula o aluno (JA TINHA CPF) sem abrir o SGE.
  2. Pega o CPF: da coluna CPF da planilha, se existir; senão, no SGE
     (Clientes > Código > Ações > Editar Cadastro > campo CPF/CNPJ — só lê, nunca salva no SGE)
  3. Apaga e cola de novo o CPF e espera NA MESMA JANELA o site trazer do SGE a
     instituição e o curso (até ~20s); depois clica em Salvar
  4. Reabre o cliente pelo mesmo código, confere se gravou e anota o que o site mudou
     (inclusive se ele trocou o Código SGE)

Nunca mexe em outro campo além do CPF e nunca clica em "Reset completo" / "Deletar".

Uso:
  python CPF_MAISBOOK.PY --listar          mostra o que foi lido da planilha
  python CPF_MAISBOOK.PY --simular         teste: acha tudo, mas NÃO cola o CPF
                                           (colar o CPF já faz o site gravar os dados do SGE)
  python CPF_MAISBOOK.PY                   recadastra de verdade
  python CPF_MAISBOOK.PY --sge 85208       um aluno só
  python CPF_MAISBOOK.PY --limite 3        só os 3 primeiros
  python CPF_MAISBOOK.PY --refazer         ignora cpf_feitos.json / cpf_pendentes.json
  python CPF_MAISBOOK.PY --recolar         cola o CPF mesmo em quem já tem CPF no MaisBook
  python CPF_MAISBOOK.PY --continuar       continua da linha onde parou (cpf_progresso.json)
  python CPF_MAISBOOK.PY --linha 120       começa na linha 120 da planilha
  python CPF_MAISBOOK.PY --conferir        só confere quem espera o SGE responder (não salva nada)
  python CPF_MAISBOOK.PY --conferir --minutos 120   continua conferindo por 2 horas
  python CPF_MAISBOOK.PY --coluna-nome "NOME COMPLETO" --coluna-cpf "CPF" --salvar
                                           grava no config quais colunas usar na planilha
  python CPF_MAISBOOK.PY --coluna-nome "CLIENTE" --listar
                                           testa a leitura com outra coluna de nome (sem gravar)

  python CPF_MAISBOOK.PY --coluna-codigo I --salvar
                                           letra da coluna com o código do MaisBook

Colunas da planilha (padrão / maisbook_config.json):
  coluna_codigo_maisbook → LETRA da coluna do código do MaisBook, usado na busca (padrão: I)
  cabecalho_sge  → código SGE (padrão: SGE)
  cabecalho_nome → nome do aluno, usado para conferir o cadastro (padrão: CLIENTE)
  cabecalho_cpf  → CPF (padrão: CPF; se vazio na planilha, busca no SGE)

A resposta do SGE (instituição e curso no MaisBook) não tem hora para chegar: quem ainda não
respondeu fica em cpf_pendentes.json e é conferido no fim da rodada e com --conferir.
  python CPF_MAISBOOK.PY --mapear          lê as telas e lista botões/campos em maisbook_mapa.txt
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
PASTA_BAIXA = PASTA.parent / "baixa_sge"
base = SourceFileLoader("baixa_base", str(PASTA_BAIXA / "BAIXA.PY")).load_module()
normalizar, AlunoPulado = base.normalizar, base.AlunoPulado

CONFIG = PASTA / "maisbook_config.json"
CREDENCIAIS = PASTA / "maisbook_credenciais.json"
PERFIL = PASTA / "maisbook_perfil"
PLANILHA = PASTA / "cpf_planilha.xlsx"
RELATORIO = PASTA / "cpf_maisbook.csv"
FEITOS = PASTA / "cpf_feitos.json"
PENDENTES = PASTA / "cpf_pendentes.json"
PROGRESSO = PASTA / "cpf_progresso.json"
PARAR = PASTA / "cpf_parar.flag"
MAPA = PASTA / "maisbook_mapa.txt"
base.PRINTS = PASTA / "maisbook_prints"

CONFIG_PADRAO = {
    "planilha_link": "https://docs.google.com/spreadsheets/d/1FeYQop9z6rVAFQuIs1EOiuQcE84t0iBgamYtMWUIAU0/edit?usp=sharing",
    "aba": "GERAL ENVIOS",
    "cabecalho_sge": "SGE",
    "cabecalho_nome": "CLIENTE",
    "cabecalho_cpf": "CPF",
    "coluna_codigo_maisbook": "I",
    # abre o MaisBook antes de buscar o CPF: quem já tem CPF válido lá é pulado (--recolar desliga)
    "pular_quem_tem_cpf": True,
    "url_maisbook": "https://maisbook.me/wp-admin/admin.php?page=maisbook-clientes",
    "menu_sge_cadastro": "Editar Cadastro",
    # se o site trouxer um curso diferente do que a pesquisa do SGE mostra, corrige pelo SGE
    "corrigir_curso": True,
    # campos do "Editar Cliente" que mostram se o site atualizou os dados do aluno
    "campos_atualizacao": {
        "cliente_nome": "Nome", "cliente_email": "E-mail", "cliente_telefone": "Telefone",
        "cliente_instituicao": "Instituição", "cliente_curso": "Curso",
        "cliente_numero_contrato": "Contrato", "cliente_plano": "Pacote",
    },
    # o aluno só conta como atualizado quando o site trouxer estes campos do SGE
    "campos_resposta_sge": ["cliente_instituicao", "cliente_curso"],
    # depois de colar o CPF, o MaisBook traz instituição/curso do SGE na própria janela (~10-20s)
    "espera_na_tela_s": 20,
    # depois de salvar, tempo para conferir se gravou
    "espera_atualizacao_s": 10,
    "intervalo_conferencia_min": 2,
    # no fim da rodada, quanto tempo continuar conferindo os que esperam o SGE (depois: --conferir)
    "conferir_ate_min": 60,
    "timeout_ms": 20000,
    "pausa_entre_alunos_ms": 800,
}

CONTADORES = {
    "ATUALIZADO": 0, "SIMULADO": 0, "AGUARDANDO SGE": 0, "JA TINHA CPF": 0,
    "NAO ENCONTRADO": 0, "CPF INVALIDO": 0, "ERRO": 0,
}


# ==========================================================
# Config e planilha
# ==========================================================

def carregar_config():
    try:
        dados = json.loads(CONFIG.read_text(encoding="utf-8"))
    except Exception:
        dados = {}
    for chave, valor in CONFIG_PADRAO.items():
        dados.setdefault(chave, json.loads(json.dumps(valor)))
    return dados


def salvar_config(dados):
    CONFIG.write_text(json.dumps(dados, indent=2, ensure_ascii=False), encoding="utf-8")


def baixar_planilha(cfg):
    return base.baixar_planilha_link(cfg["planilha_link"], PLANILHA)


def limpar_cpf(valor):
    digitos = re.sub(r"\D", "", str(valor or ""))
    return digitos.zfill(11) if 0 < len(digitos) < 11 else digitos


def cpf_valido(cpf):
    if not re.fullmatch(r"\d{11}", cpf or "") or cpf == cpf[0] * 11:
        return False
    for tamanho in (9, 10):
        soma = sum(int(cpf[i]) * (tamanho + 1 - i) for i in range(tamanho))
        if soma * 10 % 11 % 10 != int(cpf[tamanho]):
            return False
    return True


def formatar_cpf(cpf):
    return f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}" if len(cpf or "") == 11 else (cpf or "")


def indice_da_letra(letra):
    """'A' -> 0, 'I' -> 8, 'AA' -> 26."""

    letra = (letra or "").strip().upper()
    if not re.fullmatch(r"[A-Z]{1,3}", letra):
        raise ValueError(f"Coluna do código do MaisBook inválida: '{letra}' (use a letra, ex.: I)")
    indice = 0
    for caractere in letra:
        indice = indice * 26 + ord(caractere) - ord("A") + 1
    return indice - 1


def ler_planilha(cfg):
    livro = base.abrir_planilha(baixar_planilha(cfg))
    folha = livro[cfg["aba"]] if cfg.get("aba") in livro.sheetnames else livro.active
    rotulos = {chave: (cfg.get("cabecalho_" + chave) or "").strip() for chave in ("sge", "nome", "cpf")}
    procurados = {chave: normalizar(nome) for chave, nome in rotulos.items()}
    coluna_codigo = indice_da_letra(cfg.get("coluna_codigo_maisbook") or "I")

    colunas, alunos, vistos = None, [], set()
    for numero, linha in enumerate(folha.iter_rows(values_only=True), start=1):
        textos = [normalizar(base.valor_celula(v)) for v in linha]
        if colunas is None:
            if procurados["sge"] and procurados["sge"] in textos:
                colunas = {chave: textos.index(nome) if nome and nome in textos else None
                           for chave, nome in procurados.items()}
            continue

        def pegar(chave):
            indice = colunas[chave]
            return base.valor_celula(linha[indice]) if indice is not None and indice < len(linha) else ""

        sge = base.codigo_base(pegar("sge"), {"prefixo_codigo": "633."})
        if not re.fullmatch(r"\d+", sge) or sge in vistos:
            continue
        vistos.add(sge)
        codigo_bruto = base.valor_celula(linha[coluna_codigo]) if coluna_codigo < len(linha) else ""
        codigo_mb = codigo_bruto if re.fullmatch(r"[A-Za-z0-9._-]*\d[A-Za-z0-9._-]*", codigo_bruto) else ""
        alunos.append({"linha": numero, "sge": sge, "codigo_mb": codigo_mb, "codigo_bruto": codigo_bruto,
                       "nome": pegar("nome"), "cpf": limpar_cpf(pegar("cpf"))})

    livro.close()
    if colunas is None:
        raise ValueError(
            f"Não achei a coluna '{rotulos['sge']}' na aba '{folha.title}'. "
            f"Use --coluna-sge \"NOME DA COLUNA\" (e --salvar para gravar no config)."
        )
    faltando = [f"'{rotulos[k]}' ({k})" for k in ("nome", "cpf") if rotulos[k] and colunas.get(k) is None]
    if faltando:
        raise ValueError(
            f"Na aba '{folha.title}' não achei a(s) coluna(s): {', '.join(faltando)}. "
            f"Confira o cabeçalho ou ajuste com --coluna-nome / --coluna-cpf."
        )
    return alunos


def carregar_json(arquivo):
    try:
        return json.loads(arquivo.read_text(encoding="utf-8"))
    except Exception:
        return {}


def nomes_batem(nome_planilha, nome_site):
    """As duas primeiras palavras do nome da planilha precisam estar no nome do site."""

    palavras = [p for p in normalizar(nome_planilha).split() if len(p) > 2][:2]
    alvo = normalizar(nome_site)
    return all(p in alvo for p in palavras)


# ==========================================================
# SGE: ler o CPF (sem salvar nada)
# ==========================================================

JS_COLUNA = r"""
(tr, titulo) => {
  const norm = s => (s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ').trim().toUpperCase();
  const tabela = tr.closest('table');
  const cabecalho = [...tabela.querySelectorAll('tr')].find(l => l.querySelector('th'));
  if (!cabecalho) return '';
  const titulos = [...cabecalho.children].map(c => norm(c.innerText));
  const celulas = [...tr.children];
  const i = titulos.indexOf(norm(titulo));
  return i >= 0 && i < celulas.length && titulos.length === celulas.length ? celulas[i].innerText.trim() : '';
}
"""


def coluna_da_linha(linha, titulo):
    try:
        return linha.evaluate(JS_COLUNA, titulo)
    except Exception:
        return ""

def cpf_com_contas(robo, contas, aluno):
    """Pega o CPF no SGE; se a sessão cair no meio, conta a queda, entra de novo (talvez com outra conta) e repete."""

    for tentativa in range(6):
        base.ir(robo.pagina, robo.cfg["url_clientes"], robo.timeout)
        base.garantir_login(robo, contas, PARAR)
        try:
            return cpf_no_sge(robo, aluno)
        except Exception:
            if base.na_tela_de_login(robo.pagina) and tentativa < 5:
                continue
            raise
    raise AlunoPulado("ERRO", "SGE: a sessão caiu várias vezes seguidas")


def cpf_no_sge(robo, aluno):
    pagina, cfg = robo.pagina, robo.cfg
    if base.na_tela_de_login(pagina):
        raise AlunoPulado("ERRO", "SGE: voltou para o login")

    campo = base.localizar(pagina, "campo", "Código", robo.timeout)
    if campo is None:
        raise AlunoPulado("ERRO", "SGE: não achei o campo 'Código' em Clientes")
    base.preencher(campo, aluno["sge"])
    botao = base.localizar(pagina, "botao", "Pesquisar", robo.timeout)
    if botao is None:
        raise AlunoPulado("ERRO", "SGE: não achei o botão Pesquisar")
    botao.click()
    base.esperar_carregar(pagina, robo.timeout)

    codigo = cfg["prefixo_codigo"] + aluno["sge"]
    linha = pagina.locator("tr", has=pagina.get_by_role("link", name=codigo, exact=True)).first
    try:
        linha.wait_for(timeout=robo.timeout)
    except Exception:
        raise AlunoPulado("NAO ENCONTRADO", f"SGE: cliente {codigo} não apareceu na pesquisa")
    if aluno["nome"] and not nomes_batem(aluno["nome"], linha.inner_text()):
        raise AlunoPulado("ERRO", f"SGE: o nome em {codigo} não bate com a planilha — nada alterado")
    aluno["curso_sge"] = coluna_da_linha(linha, "CURSO")

    linha.locator(".dropdown-toggle").first.click()
    opcao_texto = cfg["menu_sge_cadastro"]
    opcao = linha.locator("a", has_text=opcao_texto).first
    try:
        opcao.wait_for(state="visible", timeout=5000)
    except Exception:
        vistas = [t for _, _, t in base.clicaveis(pagina, 2000)]
        raise AlunoPulado("ERRO", f"SGE: não achei '{opcao_texto}' no menu Ações. Opções: " + " | ".join(vistas[-12:]))
    opcao.click()
    base.esperar_carregar(pagina, robo.timeout)

    campo_cpf = pagina.locator("input[id*='CpfCnpj']:visible").first
    try:
        campo_cpf.wait_for(timeout=robo.timeout)
    except Exception:
        raise AlunoPulado("ERRO", "SGE: o Editar Cadastro abriu sem o campo CPF/CNPJ")
    limite = time.time() + 5
    valor = ""
    while time.time() < limite and not valor:
        valor = limpar_cpf(campo_cpf.input_value())
        if not valor:
            time.sleep(0.5)
    return valor


# ==========================================================
# MaisBook
# ==========================================================

JS_ID_CLIENTE = r"""
(sge) => {
  const folhas = [...document.querySelectorAll('body *')]
    .filter(e => e.children.length === 0 && (e.innerText || '').trim() === sge);
  for (const e of folhas) {
    let a = e;
    while (a && a !== document.body) {
      const editar = a.querySelectorAll('[title="Editar"][data-cliente-id]');
      if (editar.length === 1) return editar[0].getAttribute('data-cliente-id');
      if (editar.length > 1) break;
      a = a.parentElement;
    }
  }
  return null;
}
"""

JS_ID_POR_NOME = r"""
(nome) => {
  const norm = s => (s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ').trim().toUpperCase();
  const palavras = norm(nome).split(' ').filter(p => p.length > 2).slice(0, 2);
  if (!palavras.length) return null;
  const editars = [...document.querySelectorAll('[title="Editar"][data-cliente-id]')];
  const achados = [];
  for (const ed of editars) {
    let card = ed.parentElement;
    for (let i = 0; i < 10 && card && card !== document.body; i++, card = card.parentElement) {
      const texto = norm(card.innerText);
      if (texto.length > 2500) break;
      if (palavras.every(p => texto.includes(p))) {
        achados.push(ed.getAttribute('data-cliente-id'));
        break;
      }
    }
  }
  const unicos = [...new Set(achados)];
  return unicos.length === 1 ? unicos[0] : null;
}
"""


JS_CODIGO_NA_JANELA = r"""
(codigo) => {
  const cpf = document.querySelector('#cliente_cpf');
  if (!cpf) return false;
  const janela = cpf.closest('.modal, [role="dialog"], form') || document.body;
  const alvo = (codigo || '').trim().toUpperCase();
  if (!alvo) return false;
  const campos = [...janela.querySelectorAll('input, textarea')];
  if (campos.some(e => (e.value || '').trim().toUpperCase() === alvo)) return true;
  return [...janela.querySelectorAll('*')]
    .some(e => e.children.length === 0 && (e.innerText || '').trim().toUpperCase() === alvo);
}
"""


def codigo_maisbook(aluno):
    """Código do MaisBook da coluna I; se a célula estiver vazia, o SGE."""

    return (aluno.get("codigo_mb") or "").strip() or aluno.get("sge", "")


def login_maisbook(pagina, timeout_ms):
    try:
        cred = json.loads(CREDENCIAIS.read_text(encoding="utf-8"))
    except Exception:
        return False
    pagina.fill("#user_login", cred.get("usuario", ""))
    pagina.fill("#user_pass", cred.get("senha", ""))
    pagina.click("#wp-submit")
    base.esperar_carregar(pagina, timeout_ms)
    if pagina.locator("#user_login").count():
        return False
    print("    (login no MaisBook feito automaticamente)")
    return True


def abrir_edicao(mb, aluno, cliente_id=None, busca=None):
    """
    Pesquisa pelo código do MaisBook, clica em Editar no cartão com esse código exato
    e devolve o modal. Para reabrir o mesmo cliente, passe cliente_id (e busca, se a
    pesquisa precisar ser por outro texto, como o nome).
    """

    pagina, cfg = mb.pagina, mb.cfg
    codigo = codigo_maisbook(aluno)
    termo = (busca or codigo or "").strip()
    url = f"{cfg['url_maisbook']}&acao=listar&search={quote(termo)}"
    base.ir(pagina, url, mb.timeout)
    if pagina.locator("#user_login").count():
        if not login_maisbook(pagina, mb.timeout):
            raise AlunoPulado("ERRO", "O MaisBook pediu login e não consegui entrar (confira maisbook_credenciais.json)")
        base.ir(pagina, url, mb.timeout)

    limite = time.time() + mb.timeout / 1000
    achado = None
    while time.time() < limite:
        if cliente_id:
            achado = cliente_id if pagina.locator(f'[title="Editar"][data-cliente-id="{cliente_id}"]').count() else None
        else:
            achado = pagina.evaluate(JS_ID_CLIENTE, codigo)
        if achado or "MOSTRANDO 0 " in base.texto_da_tela(pagina) or "NENHUM CLIENTE" in base.texto_da_tela(pagina):
            break
        time.sleep(0.5)
    if not achado:
        raise AlunoPulado("NAO ENCONTRADO", f"MaisBook: nenhum cliente com o código '{termo}'")
    mb.cliente_id = achado

    pagina.click(f'[title="Editar"][data-cliente-id="{achado}"]')
    pagina.locator("#cliente_cpf").wait_for(state="visible", timeout=mb.timeout)
    limite = time.time() + 10
    # o código da coluna I pode ser o Cód. Acesso ou o Código SGE: vale qualquer campo da janela
    while not cliente_id and not pagina.evaluate(JS_CODIGO_NA_JANELA, codigo):
        if time.time() >= limite:
            raise AlunoPulado("ERRO", f"MaisBook: o Editar abriu outro cliente (sem o código {codigo} na janela) — nada alterado")
        time.sleep(0.3)

    # a janela é preenchida em duas etapas (CPF e curso chegam ~2s depois)
    base.esperar_carregar(pagina, mb.timeout)
    anterior, estavel_desde = None, time.time()
    limite = time.time() + 15
    while time.time() < limite:
        atual = ler_campos(pagina, cfg)
        if atual != anterior:
            anterior, estavel_desde = atual, time.time()
        elif time.time() - estavel_desde >= 2.5:
            break
        time.sleep(0.5)
    return pagina.locator("#cliente_cpf").locator(
        "xpath=ancestor::*[.//button[normalize-space()='Salvar']][1]")


def ler_campos(pagina, cfg):
    valores = {}
    for campo in ["cliente_cpf", "cliente_codigo_interno", *cfg["campos_atualizacao"]]:
        try:
            valores[campo] = (pagina.locator(f"#{campo}").input_value(timeout=2000) or "").strip()
        except Exception:
            valores[campo] = ""
    return valores


def fechar_edicao(modal, pagina):
    for nome in ("Cancelar", "×"):
        try:
            botao = modal.get_by_role("button", name=nome, exact=True)
            if botao.count() and botao.first.is_visible():
                botao.first.click()
                pagina.wait_for_timeout(500)
                return
        except Exception:
            pass
    pagina.keyboard.press("Escape")


def digitar_cpf(campo, cpf):
    campo.click()
    campo.press("Control+A")
    campo.press("Backspace")
    campo.type(cpf, delay=30)
    if limpar_cpf(campo.input_value()) != cpf:
        campo.fill(formatar_cpf(cpf))
    if limpar_cpf(campo.input_value()) != cpf:
        raise AlunoPulado("ERRO", f"MaisBook: não consegui digitar o CPF (ficou '{campo.input_value()}')")


def abrir_cliente(mb, aluno):
    """Abre o Editar do aluno no MaisBook, confere o nome e devolve (modal, campos)."""

    pagina = mb.pagina
    try:
        modal = abrir_edicao(mb, aluno)
    except AlunoPulado as pulo:
        # numa colagem anterior o site pode ter trocado o código: tenta achar só pelo nome
        if pulo.resultado != "NAO ENCONTRADO" or not mb.real:
            raise
        modal = abrir_por_nome(mb, aluno)
        if modal is None:
            raise
    campos = ler_campos(pagina, mb.cfg)
    if aluno["nome"] and not nomes_batem(aluno["nome"], campos.get("cliente_nome", "")):
        fechar_edicao(modal, pagina)
        raise AlunoPulado("ERRO", f"MaisBook: nome '{campos.get('cliente_nome')}' não bate com a planilha — nada alterado")
    return modal, campos


def recadastrar(mb, aluno, cpf, origem, modal, antes):
    pagina, cfg = mb.pagina, mb.cfg
    rotulos = cfg["campos_atualizacao"]

    anterior = formatar_cpf(limpar_cpf(antes["cliente_cpf"])) or "vazio"
    cliente_id = mb.cliente_id

    # colar o CPF já faz o site buscar e GRAVAR os dados do SGE (mesmo cancelando),
    # então a simulação para aqui
    if not mb.real:
        fechar_edicao(modal, pagina)
        dados = f"; já tem {resposta_sge(antes, cfg)}" if sge_respondeu(antes, cfg) else ""
        curso_sge = aluno.get("curso_sge", "") if cfg.get("corrigir_curso", True) else ""
        if curso_sge and antes.get("cliente_curso") and normalizar(antes["cliente_curso"]) != normalizar(curso_sge):
            dados += f"; curso no site '{antes['cliente_curso']}' ≠ SGE '{curso_sge}' (seria corrigido)"
        elif curso_sge:
            dados += f"; curso no SGE: {curso_sge}"
        return "SIMULADO", f"colaria {formatar_cpf(cpf)} ({origem}); CPF no MaisBook: {anterior}{dados}"

    campo_cpf = pagina.locator("#cliente_cpf")
    digitar_cpf(campo_cpf, cpf)
    campo_cpf.press("Tab")

    # com o CPF colado, o site busca instituição/curso no SGE ali mesmo na janela
    inicio = time.time()
    na_tela = antes
    while time.time() - inicio < int(cfg["espera_na_tela_s"]):
        pagina.wait_for_timeout(1000)
        na_tela = ler_campos(pagina, cfg)
        if limpar_cpf(na_tela["cliente_cpf"]) != cpf:
            digitar_cpf(campo_cpf, cpf)
            campo_cpf.press("Tab")
            continue
        # se já estava tudo preenchido, dá só uns segundos para o site responder
        if sge_respondeu(na_tela, cfg) and (not sge_respondeu(antes, cfg) or time.time() - inicio >= 4):
            break
    trouxe = [rotulos.get(c, c) for c in rotulos if na_tela.get(c) != antes.get(c)]
    na_janela = (f"o site trouxe {', '.join(trouxe)}" if trouxe else "nada mudou na janela") + \
                (f" ({resposta_sge(na_tela, cfg)})" if sge_respondeu(na_tela, cfg) else "")

    # o Código SGE do aluno é sempre o da planilha
    sge_site = na_tela.get("cliente_codigo_interno", "")
    if sge_site != aluno["sge"]:
        base.preencher(pagina.locator("#cliente_codigo_interno"), aluno["sge"])
        na_janela += f"; o site tinha trocado o SGE para {sge_site or 'vazio'} — voltei para {aluno['sge']}"

    # o curso certo é o que aparece na pesquisa do SGE para este código
    curso_sge = aluno.get("curso_sge", "").strip() if cfg.get("corrigir_curso", True) else ""
    curso_site = na_tela.get("cliente_curso", "")
    if curso_sge and normalizar(curso_site) != normalizar(curso_sge):
        base.preencher(pagina.locator("#cliente_curso"), curso_sge)
        na_janela += f"; curso corrigido: '{curso_site or 'vazio'}' → '{curso_sge}' (do SGE)"

    modal.get_by_role("button", name="Salvar", exact=True).first.click()
    base.esperar_carregar(pagina, mb.timeout)
    pagina.wait_for_timeout(1500)

    # confere se o CPF gravou
    limite = time.time() + int(cfg["espera_atualizacao_s"])
    while True:
        modal = reabrir(mb, aluno, cliente_id)
        depois = ler_campos(pagina, cfg)
        fechar_edicao(modal, pagina)
        gravado = limpar_cpf(depois["cliente_cpf"]) == cpf
        if (gravado and sge_respondeu(depois, cfg)) or time.time() >= limite:
            break
        pagina.wait_for_timeout(3000)

    if not gravado:
        raise AlunoPulado("ERRO", f"Salvei, mas o CPF não ficou gravado no MaisBook (está '{depois['cliente_cpf']}')")
    texto = f"CPF {formatar_cpf(cpf)} ({origem}; antes: {anterior}); {na_janela}"
    novo_sge = depois.get("cliente_codigo_interno", "")
    if novo_sge != aluno["sge"]:
        texto += f"; ATENÇÃO: depois de salvar o SGE ficou {novo_sge or 'vazio'} (planilha: {aluno['sge']})"
    if curso_sge and normalizar(depois.get("cliente_curso")) != normalizar(curso_sge):
        texto += f"; ATENÇÃO: depois de salvar o curso ficou '{depois.get('cliente_curso')}' (SGE: '{curso_sge}')"
    if sge_respondeu(depois, cfg):
        return "ATUALIZADO", f"{texto}; {resposta_sge(depois, cfg)}"
    faltando = [rotulos.get(c, c) for c in cfg["campos_resposta_sge"] if not depois.get(c)]
    return "AGUARDANDO SGE", f"{texto}; esperando o site trazer {', '.join(faltando)}"


def reabrir(mb, aluno, cliente_id):
    """Reabre o mesmo cliente pelo código; se o site trocou o código, procura pelo nome."""

    try:
        return abrir_edicao(mb, aluno, cliente_id=cliente_id)
    except AlunoPulado as pulo:
        if pulo.resultado != "NAO ENCONTRADO" or not (aluno.get("nome") or "").strip():
            raise
        return abrir_edicao(mb, aluno, cliente_id=cliente_id, busca=aluno["nome"].strip())


def abrir_por_nome(mb, aluno):
    """Abre o cliente achado pelo nome, se houver exatamente um cartão batendo."""

    nome = (aluno.get("nome") or "").strip()
    if not nome:
        return None
    pagina = mb.pagina
    base.ir(pagina, f"{mb.cfg['url_maisbook']}&acao=listar&search={quote(nome)}", mb.timeout)
    pagina.wait_for_timeout(1500)
    achado = pagina.evaluate(JS_ID_POR_NOME, nome)
    if not achado:
        return None
    return abrir_edicao(mb, aluno, cliente_id=achado, busca=nome)


def sge_respondeu(campos, cfg):
    return all(campos.get(c) for c in cfg["campos_resposta_sge"])


def resposta_sge(campos, cfg):
    rotulos = cfg["campos_atualizacao"]
    return ", ".join(f"{rotulos.get(c, c)}: {campos.get(c)}" for c in cfg["campos_resposta_sge"])


def salvar_json(arquivo, dados):
    arquivo.write_text(json.dumps(dados, indent=2, ensure_ascii=False), encoding="utf-8")


def salvar_relatorio(linhas_csv):
    try:
        with RELATORIO.open("w", newline="", encoding="utf-8-sig") as arquivo:
            escritor = csv.writer(arquivo, delimiter=";")
            escritor.writerow(["LINHA", "SGE", "NOME", "CPF", "ORIGEM CPF", "RESULTADO", "DETALHE"])
            escritor.writerows(linhas_csv)
    except PermissionError:
        print(f"    (não consegui gravar {RELATORIO.name} — ele está aberto no Excel?)")


def pediu_parar():
    if PARAR.exists():
        PARAR.unlink(missing_ok=True)
        return True
    return False


def conferir_pendentes(mb, feitos, pendentes, linhas_csv, ate_min):
    """Reabre os alunos que estão esperando a resposta do SGE até ela chegar (ou o tempo acabar)."""

    cfg = mb.cfg
    intervalo = max(float(cfg.get("intervalo_conferencia_min") or 3), 0.5)
    inicio, rodada = time.time(), 0
    while pendentes:
        rodada += 1
        print(f"\n--- Conferindo {len(pendentes)} aluno(s) esperando o SGE responder (rodada {rodada}) ---")
        for sge, dados in list(pendentes.items()):
            if pediu_parar():
                print("    Conferência parada a pedido. Depois é só usar --conferir.")
                return
            try:
                cliente_id = dados.get("cliente_id")
                aluno_pend = {"sge": sge, "codigo_mb": dados.get("codigo_mb", ""), "nome": dados.get("nome", "")}
                modal = reabrir(mb, aluno_pend, cliente_id) if cliente_id else abrir_edicao(mb, aluno_pend)
                campos = ler_campos(mb.pagina, cfg)
                fechar_edicao(modal, mb.pagina)
            except Exception as erro:
                if "has been closed" in str(erro):
                    raise
                print(f"    SGE {sge}: não consegui abrir agora ({getattr(erro, 'detalhe', '') or base.erro_curto(erro)})")
                continue
            if not sge_respondeu(campos, cfg):
                continue
            try:
                espera = (datetime.now() - datetime.strptime(dados["salvo_em"], "%Y-%m-%d %H:%M:%S")).total_seconds() / 60
                quando = f" — {espera:.0f} min depois de salvar"
            except Exception:
                quando = ""
            detalhe = resposta_sge(campos, cfg) + quando
            print(f"    SGE {sge} — {dados.get('nome', '')} — ATUALIZADO — {detalhe}")
            CONTADORES["ATUALIZADO"] += 1
            linhas_csv.append([dados.get("linha", ""), sge, dados.get("nome", ""), formatar_cpf(dados.get("cpf", "")),
                               "", "ATUALIZADO", detalhe])
            feitos[sge] = {"nome": dados.get("nome", ""), "cpf": dados.get("cpf", ""),
                           "data": f"{datetime.now():%Y-%m-%d %H:%M:%S}"}
            pendentes.pop(sge)
            salvar_json(FEITOS, feitos)
            salvar_json(PENDENTES, pendentes)

        if not pendentes:
            print("    Todos responderam.")
            break
        if (time.time() - inicio) / 60 + intervalo > ate_min:
            print(f"    Ainda esperando {len(pendentes)}. Rode mais tarde: python CPF_MAISBOOK.PY --conferir")
            break
        print(f"    Ainda esperando {len(pendentes)} — confiro de novo em {intervalo:g} min "
              "(Parar a qualquer momento; depois é só rodar --conferir)")
        fim = time.time() + intervalo * 60
        while time.time() < fim:
            if pediu_parar():
                print("    Conferência parada a pedido. Depois é só usar --conferir.")
                return
            time.sleep(1)


# ==========================================================
# Execução
# ==========================================================

def abrir_maisbook(playwright):
    PERFIL.mkdir(exist_ok=True)
    contexto = playwright.chromium.launch_persistent_context(
        str(PERFIL), headless=False, no_viewport=True, args=["--start-maximized"])
    return contexto, (contexto.pages[0] if contexto.pages else contexto.new_page())


def rodar(cfg, args):
    from playwright.sync_api import sync_playwright

    feitos, pendentes = carregar_json(FEITOS), carregar_json(PENDENTES)
    real = not args.simular
    seguinte = {}
    if args.conferir:
        alunos, ja_feitos = [], []
    else:
        alunos = ler_planilha(cfg)
        if args.sge:
            procurado = args.sge.split(".")[-1]
            alunos = [a for a in alunos if procurado in (a["sge"], a["codigo_mb"])]
            if not alunos:
                print(f"O código {args.sge} não está na planilha.")
                return
        ja_feitos = [a for a in alunos if a["sge"] in feitos and not args.refazer]
        alunos = [a for a in alunos
                  if args.refazer or (a["sge"] not in feitos and a["sge"] not in pendentes)]
        inicio = args.linha
        if args.continuar:
            inicio = carregar_json(PROGRESSO).get("proxima_linha")
            if not inicio:
                print("Não há linha salva para continuar — começando do início.")
        if inicio:
            alunos = [a for a in alunos if a["linha"] >= int(inicio)]
            print(f"Começando na linha {inicio} da planilha.")
        seguinte = {a["linha"]: b["linha"] for a, b in zip(alunos, alunos[1:])}
        if args.limite:
            alunos = alunos[: args.limite]
    guardar_progresso = real and not args.conferir and not args.sge
    PARAR.unlink(missing_ok=True)

    print("=" * 70)
    if args.conferir:
        print(f"ROBÔ MAISBOOK — CONFERIR QUEM ESPERA O SGE — {len(pendentes)} aluno(s) na fila")
    else:
        print("ROBÔ MAISBOOK — RECADASTRAR CPF —", "MODO REAL (vai salvar)" if real else "SIMULAÇÃO (não cola o CPF)")
        print(f"Planilha: aba {cfg['aba']} | colunas: código MaisBook={cfg['coluna_codigo_maisbook']} "
              f"SGE='{cfg['cabecalho_sge']}' "
              f"nome='{cfg['cabecalho_nome']}' CPF='{cfg['cabecalho_cpf']}'")
        print(f"Alunos para processar: {len(alunos)} | já atualizados: {len(ja_feitos)}"
              f" | esperando o SGE: {len(pendentes)}")
    print("=" * 70)

    linhas_csv = []
    for a in ja_feitos:
        resultado = feitos[a["sge"]].get("resultado", "ATUALIZADO")
        CONTADORES[resultado] = CONTADORES.get(resultado, 0) + 1
        linhas_csv.append([a["linha"], a["sge"], a["nome"], formatar_cpf(feitos[a["sge"]].get("cpf", "")), "",
                           resultado, "registrado em cpf_feitos.json"])
    if args.minutos is not None:
        ate_min = args.minutos
    else:
        ate_min = 0 if args.conferir else float(cfg.get("conferir_ate_min") or 0)
    conferir = real and (args.conferir or ate_min > 0)

    if alunos or (conferir and pendentes):
        with sync_playwright() as p:
            mb_contexto, mb_pagina = abrir_maisbook(p)
            mb = base.Robo(mb_contexto, mb_pagina, cfg, real, False)
            mb_pagina.on("dialog", mb._dialogo)

            sge, contas = None, None

            def abrir_sge():
                nonlocal sge, contas
                if sge is None:
                    sge_cfg = base.carregar_config()
                    sge_cfg["menu_sge_cadastro"] = cfg["menu_sge_cadastro"]
                    sge_contexto, sge_pagina = base.abrir_navegador(p)
                    sge = base.Robo(sge_contexto, sge_pagina, sge_cfg, False, False)
                    sge_pagina.on("dialog", sge._dialogo)
                    contas = base.ContasSGE()
                    print(f"    (abrindo o SGE — contas: {', '.join(c['usuario'] for c in contas.contas)})")
                return sge

            pular_com_cpf = bool(cfg.get("pular_quem_tem_cpf", True)) and not args.recolar
            parado = False
            for posicao, aluno in enumerate(alunos, start=1):
                if pediu_parar():
                    print(f"\nPARADO A PEDIDO na linha {aluno['linha']}. Para continuar: --continuar")
                    parado = True
                    break
                print(f"\n[{posicao}/{len(alunos)}] código MaisBook {codigo_maisbook(aluno)} | SGE {aluno['sge']}"
                      f" — {aluno['nome']} (linha {aluno['linha']})")
                cpf, origem = aluno["cpf"], "planilha"

                def registrar(resultado, detalhe="", imagem=""):
                    CONTADORES[resultado] += 1
                    print(f"    {resultado}" + (f" — {detalhe}" if detalhe else "") + (f"\n      print: {imagem}" if imagem else ""))
                    linhas_csv.append([aluno["linha"], aluno["sge"], aluno["nome"], formatar_cpf(cpf), origem,
                                       resultado, (detalhe + (" | " + imagem if imagem else "")).strip(" |")])
                    agora = f"{datetime.now():%Y-%m-%d %H:%M:%S}"
                    salvar_relatorio(linhas_csv)
                    if resultado in ("ATUALIZADO", "JA TINHA CPF"):
                        feitos[aluno["sge"]] = {"nome": aluno["nome"], "cpf": cpf, "data": agora,
                                                "resultado": resultado}
                        pendentes.pop(aluno["sge"], None)
                        salvar_json(FEITOS, feitos)
                        salvar_json(PENDENTES, pendentes)
                    elif resultado == "AGUARDANDO SGE":
                        pendentes[aluno["sge"]] = {"nome": aluno["nome"], "codigo_mb": aluno.get("codigo_mb", ""),
                                                   "cpf": cpf, "linha": aluno["linha"],
                                                   "cliente_id": mb.cliente_id, "salvo_em": agora}
                        salvar_json(PENDENTES, pendentes)

                pagina_erro = mb.pagina
                try:
                    # 1º o MaisBook: se o cadastro já tem CPF válido, não precisa ir ao SGE
                    modal, antes = abrir_cliente(mb, aluno)
                    cpf_site = limpar_cpf(antes.get("cliente_cpf"))
                    if pular_com_cpf and cpf_valido(cpf_site):
                        fechar_edicao(modal, mb.pagina)
                        detalhe = f"o MaisBook já tem o CPF {formatar_cpf(cpf_site)} — pulei"
                        if cpf and cpf != cpf_site:
                            detalhe += f"; ATENÇÃO: a planilha tem outro CPF ({formatar_cpf(cpf)})"
                        if not sge_respondeu(antes, cfg):
                            faltando = [cfg["campos_atualizacao"].get(c, c)
                                        for c in cfg["campos_resposta_sge"] if not antes.get(c)]
                            detalhe += f"; sem {', '.join(faltando)} no MaisBook"
                        cpf, origem = cpf_site, "MaisBook"
                        registrar("JA TINHA CPF", detalhe)
                    else:
                        if not cpf:
                            origem = "SGE"
                            pagina_erro = abrir_sge().pagina
                            cpf = cpf_com_contas(sge, contas, aluno)
                            pagina_erro = mb.pagina
                        if not cpf:
                            fechar_edicao(modal, mb.pagina)
                            registrar("CPF INVALIDO", f"sem CPF ({origem})")
                        elif not cpf_valido(cpf):
                            fechar_edicao(modal, mb.pagina)
                            registrar("CPF INVALIDO", f"'{formatar_cpf(cpf)}' não é um CPF válido ({origem})")
                        else:
                            registrar(*recadastrar(mb, aluno, cpf, origem, modal, antes))
                except AlunoPulado as pulo:
                    registrar(pulo.resultado, pulo.detalhe, base.tirar_print(pagina_erro, f"erro_{aluno['sge']}"))
                except Exception as erro:
                    if "has been closed" in str(erro):
                        print(f"\nO navegador foi fechado — parando na linha {aluno['linha']}. Para continuar: --continuar")
                        parado = True
                        break
                    registrar("ERRO", base.erro_curto(erro), base.tirar_print(pagina_erro, f"erro_{aluno['sge']}"))

                if guardar_progresso:
                    proxima = seguinte.get(aluno["linha"])
                    salvar_json(PROGRESSO, {
                        "ultima_linha": aluno["linha"], "sge": aluno["sge"], "nome": aluno["nome"],
                        "proxima_linha": proxima, "concluido": proxima is None,
                        "data": f"{datetime.now():%Y-%m-%d %H:%M:%S}"})
                time.sleep(int(cfg.get("pausa_entre_alunos_ms") or 0) / 1000)

            if sge:
                try:
                    sge.contexto.close()
                except Exception:
                    pass
            if conferir and pendentes and not parado:
                try:
                    conferir_pendentes(mb, feitos, pendentes, linhas_csv, ate_min)
                except KeyboardInterrupt:
                    print("\nConferência interrompida. Depois é só rodar: python CPF_MAISBOOK.PY --conferir")
                except Exception as erro:
                    print(f"\nConferência parou: {base.erro_curto(erro)}")
            try:
                mb_contexto.close()
            except Exception:
                pass

    salvar_relatorio(linhas_csv)
    print("\n" + "=" * 70)
    print(f"RELATORIO: {RELATORIO}")
    principal = f"ATUALIZADOS: {CONTADORES['ATUALIZADO']}" if real else f"SIMULADOS: {CONTADORES['SIMULADO']}"
    print(f"RESUMO CPF — {principal} | JÁ TINHAM CPF (pulados): {CONTADORES['JA TINHA CPF']} | "
          f"ESPERANDO O SGE: {len(pendentes)} | "
          f"NÃO ENCONTRADOS: {CONTADORES['NAO ENCONTRADO']} | CPF INVÁLIDO: {CONTADORES['CPF INVALIDO']} | "
          f"ERROS: {CONTADORES['ERRO']}")
    if pendentes and real:
        print("Para conferir os que ainda esperam o SGE: python CPF_MAISBOOK.PY --conferir")
    print("=" * 70)


def mapear(cfg):
    from playwright.sync_api import sync_playwright

    MAPA.write_text(f"MAPA DO MAISBOOK — {datetime.now():%d/%m/%Y %H:%M}\n", encoding="utf-8")
    with sync_playwright() as p:
        contexto, pagina = abrir_maisbook(p)
        base.ir(pagina, cfg["url_maisbook"], int(cfg["timeout_ms"]))
        if pagina.locator("#user_login").count():
            login_maisbook(pagina, int(cfg["timeout_ms"]))
        print("Navegador aberto. Vá até a tela que quer mapear.")
        tela = 0
        while input("\nEnter = ler esta tela | q + Enter = terminar: ").strip().lower() != "q":
            pagina = contexto.pages[-1]
            tela += 1
            linhas = [f"\n{'=' * 70}", f"TELA {tela} — {pagina.title()}", f"URL: {pagina.url}", "=" * 70]
            for frame in pagina.frames:
                try:
                    dados = frame.evaluate(base.JS_MAPA)
                except Exception:
                    continue
                for chave, lista in dados.items():
                    if lista:
                        linhas.append(f"\n{chave.upper()}:")
                        linhas.extend(f"  - {item}" for item in lista)
            imagem = PASTA / f"maisbook_mapa_tela{tela}.png"
            pagina.screenshot(path=str(imagem), full_page=True)
            with MAPA.open("a", encoding="utf-8") as arquivo:
                arquivo.write("\n".join(linhas) + f"\nPRINT: {imagem}\n")
            print(f"Tela {tela} salva em {MAPA.name}")
        contexto.close()


def main():
    parser = argparse.ArgumentParser(description="Robô MaisBook — recadastrar o CPF dos alunos")
    parser.add_argument("--listar", action="store_true", help="só mostra o que foi lido da planilha")
    parser.add_argument("--mapear", action="store_true")
    parser.add_argument("--simular", action="store_true", help="teste: acha tudo, mas não cola o CPF")
    parser.add_argument("--sge", help="processa só este código")
    parser.add_argument("--limite", type=int)
    parser.add_argument("--refazer", action="store_true", help="ignora cpf_feitos.json e cpf_pendentes.json")
    parser.add_argument("--recolar", action="store_true",
                        help="cola o CPF de novo mesmo em quem já tem CPF no MaisBook")
    parser.add_argument("--linha", type=int, help="começa nesta linha da planilha")
    parser.add_argument("--continuar", action="store_true", help="continua da linha onde parou (cpf_progresso.json)")
    parser.add_argument("--conferir", action="store_true",
                        help="só confere quem está esperando a resposta do SGE (não salva nada)")
    parser.add_argument("--minutos", type=float,
                        help="quanto tempo continuar conferindo os que esperam o SGE (padrão: conferir_ate_min)")
    parser.add_argument("--link", help="outro link de planilha do Google Drive")
    parser.add_argument("--aba", help="nome da aba da planilha")
    parser.add_argument("--coluna-sge", dest="coluna_sge",
                        help="nome do cabeçalho da coluna do código SGE (padrão: SGE)")
    parser.add_argument("--coluna-codigo", dest="coluna_codigo",
                        help="LETRA da coluna com o código do MaisBook, usado na busca (padrão: I)")
    parser.add_argument("--coluna-nome", dest="coluna_nome",
                        help="nome do cabeçalho da coluna do nome do aluno (padrão: CLIENTE)")
    parser.add_argument("--coluna-cpf", dest="coluna_cpf",
                        help="nome do cabeçalho da coluna do CPF (padrão: CPF)")
    parser.add_argument("--salvar", action="store_true",
                        help="grava link/aba/colunas no maisbook_config.json")
    args = parser.parse_args()

    cfg = carregar_config()
    if not CONFIG.exists():
        salvar_config(cfg)
    if args.link:
        cfg["planilha_link"] = args.link
    if args.aba:
        cfg["aba"] = args.aba
    if args.coluna_sge:
        cfg["cabecalho_sge"] = args.coluna_sge.strip()
    if args.coluna_codigo:
        indice_da_letra(args.coluna_codigo)
        cfg["coluna_codigo_maisbook"] = args.coluna_codigo.strip().upper()
    if args.coluna_nome:
        cfg["cabecalho_nome"] = args.coluna_nome.strip()
    if args.coluna_cpf:
        cfg["cabecalho_cpf"] = args.coluna_cpf.strip()
    if args.salvar:
        salvar_config(cfg)
        print(f"Config salvo em {CONFIG.name}:")
        print(f"  aba={cfg['aba']} | código MaisBook=coluna {cfg['coluna_codigo_maisbook']} | "
              f"SGE={cfg['cabecalho_sge']} | nome={cfg['cabecalho_nome']} | CPF={cfg['cabecalho_cpf']}")

    if args.mapear:
        mapear(cfg)
    elif args.listar:
        print(f"Colunas: código MaisBook=coluna {cfg['coluna_codigo_maisbook']} | SGE='{cfg['cabecalho_sge']}' | "
              f"nome='{cfg['cabecalho_nome']}' | CPF='{cfg['cabecalho_cpf']}' | aba='{cfg['aba']}'")
        alunos = ler_planilha(cfg)
        for aluno in alunos:
            cpf = formatar_cpf(aluno["cpf"]) if aluno["cpf"] else "(CPF vem do SGE)"
            if aluno["codigo_mb"]:
                codigo = aluno["codigo_mb"]
            elif aluno["codigo_bruto"]:
                codigo = f"('{aluno['codigo_bruto']}' não é código — usa o SGE {aluno['sge']})"
            else:
                codigo = f"(vazio — usa o SGE {aluno['sge']})"
            print(f"linha {aluno['linha']}: código MaisBook {codigo} | SGE {aluno['sge']} — {aluno['nome']} | {cpf}")
        print(f"\nTotal: {len(alunos)} alunos (aba {cfg['aba']})")
    else:
        rodar(cfg, args)


if __name__ == "__main__":
    try:
        main()
    except (FileNotFoundError, KeyError, ValueError) as erro:
        print(f"\nERRO: {erro}")
        sys.exit(1)
