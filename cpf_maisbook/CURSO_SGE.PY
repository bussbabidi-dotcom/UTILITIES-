# -*- coding: utf-8 -*-
"""
ROBÔ CURSO PELO SGE — PREENCHER O CURSO NA PLANILHA DO GOOGLE

Para cada linha da planilha do Google Drive (bloco "curso_sge" do maisbook_config.json):
  1. Lê o nome do aluno na coluna do nome (título do cabeçalho, ex: CLIENTE, ou letra, ex: C)
  2. SGE (Clientes): pesquisa pelo nome e lê a coluna CURSO do resultado (só lê, nunca salva no SGE)
  3. Google Planilhas (no Chrome, logado na sua conta): vai até a célula da coluna de destino
     pela caixa de nome e digita o curso (sobrescreve o que estiver lá)

Uso:
  python CURSO_SGE.PY --listar          mostra os nomes lidos da planilha
  python CURSO_SGE.PY --simular         pesquisa no SGE, mas NÃO cola na planilha
  python CURSO_SGE.PY                   preenche de verdade
  python CURSO_SGE.PY --limite 3        só os 3 primeiros
  python CURSO_SGE.PY --linha 120       começa na linha 120 da planilha
  python CURSO_SGE.PY --continuar       continua da linha onde parou (curso_progresso.json)
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

PASTA = Path(__file__).resolve().parent
cpf = SourceFileLoader("cpf_maisbook", str(PASTA / "CPF_MAISBOOK.PY")).load_module()
base, normalizar, AlunoPulado, nomes_batem = cpf.base, cpf.normalizar, cpf.AlunoPulado, cpf.nomes_batem

CONFIG = PASTA / "maisbook_config.json"
PERFIL_PLANILHA = PASTA / "planilha_perfil"
PLANILHA = PASTA / "curso_planilha.xlsx"
RELATORIO = PASTA / "curso_sge.csv"
PROGRESSO = PASTA / "curso_progresso.json"
PARAR = PASTA / "curso_parar.flag"
base.PRINTS = PASTA / "maisbook_prints"

CONFIG_PADRAO = {
    "planilha_link": "",
    "aba": "",
    "coluna_nome": "Nome",
    "coluna_destino": "CURSO",
    # rótulo do campo de nome na tela Clientes do SGE
    "rotulo_nome_sge": "Nome",
    "timeout_ms": 20000,
    "pausa_entre_alunos_ms": 500,
    # na primeira vez é preciso entrar na conta Google no navegador do robô
    "espera_login_google_min": 10,
}

TITULOS_NOME_SGE = ("NOME", "CLIENTE", "NOME DO CLIENTE", "RAZAO SOCIAL", "NOME/RAZAO SOCIAL")

CONTADORES = {"COLADO": 0, "SIMULADO": 0, "NAO ENCONTRADO": 0, "VARIOS NO SGE": 0, "ERRO": 0}


# ==========================================================
# Config e planilha
# ==========================================================

def carregar_config():
    try:
        dados = json.loads(CONFIG.read_text(encoding="utf-8"))
    except Exception:
        dados = {}
    bloco = dados.get("curso_sge") or {}
    for chave, valor in CONFIG_PADRAO.items():
        bloco.setdefault(chave, valor)
    return bloco


def salvar_config(bloco):
    try:
        dados = json.loads(CONFIG.read_text(encoding="utf-8"))
    except Exception:
        dados = {}
    dados["curso_sge"] = bloco
    CONFIG.write_text(json.dumps(dados, indent=2, ensure_ascii=False), encoding="utf-8")


def parece_letra(texto):
    return bool(re.fullmatch(r"[A-Za-z]{1,3}", (texto or "").strip()))


def ler_planilha(cfg):
    """Devolve (alunos, letra da coluna de destino, nome da aba)."""

    from openpyxl.utils import get_column_letter

    if not cfg.get("planilha_link"):
        raise ValueError("Falta o link da planilha (aba 'Curso pelo SGE').")
    livro = base.abrir_planilha(base.baixar_planilha_link(cfg["planilha_link"], PLANILHA))
    aba = (cfg.get("aba") or "").strip()
    if aba and aba not in livro.sheetnames:
        livro.close()
        raise ValueError(f"A aba '{aba}' não existe. Abas: {', '.join(livro.sheetnames)}")
    folha = livro[aba] if aba else livro.active
    linhas = [[normalizar(base.valor_celula(v)) for v in linha] for linha in folha.iter_rows(values_only=True)]
    originais = [[base.valor_celula(v) for v in linha] for linha in folha.iter_rows(values_only=True)]
    titulo_aba = folha.title
    livro.close()

    coluna_nome, coluna_destino = cfg["coluna_nome"].strip(), cfg["coluna_destino"].strip()
    alvo_nome = normalizar(coluna_nome)
    cabecalho = next((i for i, textos in enumerate(linhas) if alvo_nome in textos), None)
    if cabecalho is not None:
        i_nome = linhas[cabecalho].index(alvo_nome)
    elif parece_letra(coluna_nome):
        cabecalho, i_nome = 0, base.indice_coluna(coluna_nome) - 1
    else:
        raise ValueError(f"Não achei a coluna '{coluna_nome}' na aba '{titulo_aba}'")

    alvo_destino = normalizar(coluna_destino)
    if alvo_destino in linhas[cabecalho]:
        i_destino = linhas[cabecalho].index(alvo_destino)
    elif parece_letra(coluna_destino):
        i_destino = base.indice_coluna(coluna_destino) - 1
    else:
        raise ValueError(f"Não achei a coluna de destino '{coluna_destino}' no cabeçalho (linha {cabecalho + 1})")

    alunos = []
    for numero, linha in enumerate(originais[cabecalho + 1:], start=cabecalho + 2):
        nome = linha[i_nome].strip() if i_nome < len(linha) else ""
        if nome:
            atual = linha[i_destino] if i_destino < len(linha) else ""
            alunos.append({"linha": numero, "nome": nome, "atual": atual})
    return alunos, get_column_letter(i_destino + 1), titulo_aba


def salvar_json(arquivo, dados):
    arquivo.write_text(json.dumps(dados, indent=2, ensure_ascii=False), encoding="utf-8")


def carregar_json(arquivo):
    try:
        return json.loads(arquivo.read_text(encoding="utf-8"))
    except Exception:
        return {}


def salvar_relatorio(linhas_csv):
    try:
        with RELATORIO.open("w", newline="", encoding="utf-8-sig") as arquivo:
            escritor = csv.writer(arquivo, delimiter=";")
            escritor.writerow(["LINHA", "NOME", "CURSO NO SGE", "RESULTADO", "DETALHE"])
            escritor.writerows(linhas_csv)
    except PermissionError:
        print(f"    (não consegui gravar {RELATORIO.name} — ele está aberto no Excel?)")


def pediu_parar():
    if PARAR.exists():
        PARAR.unlink(missing_ok=True)
        return True
    return False


# ==========================================================
# SGE: pesquisar pelo nome e ler o curso
# ==========================================================

JS_RESULTADOS = r"""
() => {
  const norm = s => (s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ').trim().toUpperCase();
  const vis = e => { const r = e.getBoundingClientRect(); return r.width > 0 && r.height > 0; };
  const saida = [];
  for (const tabela of document.querySelectorAll('table')) {
    if (!vis(tabela)) continue;
    const cabecalho = [...tabela.querySelectorAll('tr')].find(l => l.querySelector('th'));
    if (!cabecalho) continue;
    const titulos = [...cabecalho.children].map(c => norm(c.innerText));
    if (!titulos.includes('CURSO')) continue;
    for (const tr of tabela.querySelectorAll('tr')) {
      if (tr === cabecalho || tr.querySelector('th')) continue;
      const celulas = [...tr.children];
      if (celulas.length !== titulos.length) continue;
      const linha = {};
      titulos.forEach((t, i) => { linha[t] = celulas[i].innerText.trim(); });
      linha['__TEXTO__'] = tr.innerText.trim();
      saida.push(linha);
    }
  }
  return saida;
}
"""


def resultados_na_tela(pagina):
    linhas = []
    for frame in pagina.frames:
        try:
            linhas.extend(frame.evaluate(JS_RESULTADOS))
        except Exception:
            continue
    return linhas


def nome_da_linha(linha):
    for titulo in TITULOS_NOME_SGE:
        if linha.get(titulo):
            return linha[titulo]
    return linha.get("__TEXTO__", "")


def escolher_curso(nome, linhas):
    """Devolve (curso, detalhe). Levanta AlunoPulado quando não dá para decidir."""

    alvo = normalizar(nome)
    exatos = [l for l in linhas if normalizar(nome_da_linha(l)) == alvo]
    candidatos = exatos or [l for l in linhas if nomes_batem(nome, nome_da_linha(l))]
    if not candidatos:
        vistos = " | ".join(nome_da_linha(l) for l in linhas[:5])
        raise AlunoPulado("NAO ENCONTRADO", "SGE: nenhum resultado com esse nome" + (f" (vi: {vistos})" if vistos else ""))
    cursos = {normalizar(l.get("CURSO")): l.get("CURSO", "").strip() for l in candidatos}
    if len(cursos) > 1:
        opcoes = " | ".join(f"{nome_da_linha(l)}: {l.get('CURSO', '')}" for l in candidatos[:6])
        raise AlunoPulado("VARIOS NO SGE", f"{len(candidatos)} alunos com cursos diferentes — {opcoes}")
    curso = next(iter(cursos.values()))
    if not curso:
        raise AlunoPulado("NAO ENCONTRADO", f"SGE: '{nome_da_linha(candidatos[0])}' está sem curso")
    detalhe = "" if exatos else f"nome no SGE: {nome_da_linha(candidatos[0])}"
    if len(candidatos) > 1:
        detalhe = (detalhe + "; " if detalhe else "") + f"{len(candidatos)} cadastros, mesmo curso"
    return curso, detalhe


def curso_no_sge(robo, aluno):
    pagina, cfg = robo.pagina, robo.cfg
    if base.na_tela_de_login(pagina):
        raise AlunoPulado("ERRO", "SGE: voltou para o login")

    rotulo = robo.rotulo_nome
    campo = base.localizar(pagina, "campo", rotulo, robo.timeout)
    if campo is None:
        raise AlunoPulado("ERRO", f"SGE: não achei o campo '{rotulo}' em Clientes (troque 'rotulo_nome_sge' na config)")
    base.preencher(campo, aluno["nome"])
    botao = base.localizar(pagina, "botao", "Pesquisar", robo.timeout)
    if botao is None:
        raise AlunoPulado("ERRO", "SGE: não achei o botão Pesquisar")
    botao.click()
    base.esperar_carregar(pagina, robo.timeout)

    limite = time.time() + 8
    linhas = resultados_na_tela(pagina)
    while not linhas and time.time() < limite:
        if base.na_tela_de_login(pagina):
            raise AlunoPulado("ERRO", "SGE: voltou para o login")
        time.sleep(0.5)
        linhas = resultados_na_tela(pagina)
    return escolher_curso(aluno["nome"], linhas)


def curso_com_contas(robo, contas, aluno):
    """Pesquisa o curso; se a sessão cair no meio, entra de novo (talvez com outra conta) e repete."""

    for tentativa in range(6):
        base.ir(robo.pagina, robo.cfg["url_clientes"], robo.timeout)
        base.garantir_login(robo, contas, PARAR)
        try:
            return curso_no_sge(robo, aluno)
        except Exception:
            if base.na_tela_de_login(robo.pagina) and tentativa < 5:
                continue
            raise
    raise AlunoPulado("ERRO", "SGE: a sessão caiu várias vezes seguidas")


# ==========================================================
# Google Planilhas (no navegador)
# ==========================================================

class SemPermissao(Exception):
    pass


RE_FAZER_LOGIN = re.compile(r"^\s*(Fazer login|Sign in)\s*$", re.I)
RE_SOMENTE_VER = re.compile(r"^\s*(Somente ver|Somente visualiza\w+|Apenas visualiza\w+|View only)\s*$", re.I)


def visivel(localizador):
    try:
        return any(localizador.nth(i).is_visible() for i in range(min(localizador.count(), 5)))
    except Exception:
        return False


def estado_google(pagina):
    """'carregando', 'sem_login', 'somente_ver' ou 'editar'."""

    if "/spreadsheets/" not in pagina.url or not visivel(pagina.locator("#t-name-box")):
        return "carregando"
    if visivel(pagina.get_by_text(RE_SOMENTE_VER)):
        return "sem_login" if visivel(pagina.get_by_text(RE_FAZER_LOGIN)) else "somente_ver"
    if visivel(pagina.get_by_text(RE_FAZER_LOGIN)):
        return "sem_login"
    return "editar"


def abrir_google(playwright, cfg):
    PERFIL_PLANILHA.mkdir(exist_ok=True)
    opcoes = dict(headless=False, no_viewport=True, ignore_default_args=["--enable-automation"],
                  args=["--start-maximized", "--disable-blink-features=AutomationControlled"])
    try:
        # o login do Google costuma recusar o Chromium do Playwright; o Chrome instalado passa
        contexto = playwright.chromium.launch_persistent_context(str(PERFIL_PLANILHA), channel="chrome", **opcoes)
    except Exception:
        contexto = playwright.chromium.launch_persistent_context(str(PERFIL_PLANILHA), **opcoes)
    pagina = contexto.pages[0] if contexto.pages else contexto.new_page()
    pagina.goto(cfg["planilha_link"], timeout=60000)

    minutos = float(cfg.get("espera_login_google_min") or 10)
    limite = time.time() + minutos * 60
    avisou, estado = False, "carregando"
    while time.time() < limite:
        if PARAR.exists():
            raise SemPermissao("pediram para parar antes de entrar na conta Google")
        pagina = contexto.pages[-1]
        pagina.wait_for_timeout(1500)
        if "/spreadsheets/" not in pagina.url and "accounts.google" not in pagina.url:
            try:
                pagina.goto(cfg["planilha_link"], timeout=60000)
            except Exception:
                pass
            continue
        estado = estado_google(pagina)
        if estado == "editar":
            break
        if estado == "somente_ver":
            raise SemPermissao("a conta Google que está logada só pode VER esta planilha. "
                               "Peça permissão de EDITOR para essa conta (ou entre com outra no navegador do robô).")
        if estado == "sem_login" and not avisou:
            print(f"ATENÇÃO: o navegador do robô não está logado no Google (a planilha abriu em 'Somente ver').\n"
                  f"    Clique em 'Fazer login' no navegador que abriu e entre com uma conta que pode EDITAR a planilha "
                  f"(espero até {minutos:g} min).")
            avisou = True
    else:
        raise SemPermissao("não deu para editar a planilha no Google "
                           + ("(ninguém entrou na conta Google)" if estado == "sem_login" else f"(estado: {estado})"))
    if avisou:
        print("    (Google: conta logada com permissão de edição)")
    pagina.wait_for_timeout(1500)
    return contexto, pagina


def ir_para_celula(pagina, aba, celula):
    caixa = pagina.locator("#t-name-box")
    caixa.click()
    caixa.fill("'" + aba.replace("'", "''") + "'!" + celula)
    caixa.press("Enter")
    pagina.wait_for_timeout(400)


def valor_na_barra(pagina):
    try:
        texto = pagina.locator("#t-formula-bar-input").first.inner_text(timeout=3000)
        return texto.replace("\ufeff", "").replace("\u200b", "").strip()
    except Exception:
        return None


def escrever_celula(pagina, aba, celula, texto):
    if texto[:1] in "=+-@":
        texto = "'" + texto
    ir_para_celula(pagina, aba, celula)
    pagina.keyboard.type(texto, delay=10)
    # o autocompletar do Planilhas deixa o resto da sugestão selecionado; Delete tira
    pagina.keyboard.press("Delete")
    pagina.keyboard.press("Enter")
    pagina.wait_for_timeout(500)

    ir_para_celula(pagina, aba, celula)
    gravado = valor_na_barra(pagina)
    if gravado is not None and normalizar(gravado.lstrip("'")) != normalizar(texto.lstrip("'")):
        estado = estado_google(pagina)
        if estado in ("sem_login", "somente_ver"):
            raise SemPermissao("a planilha está em 'Somente ver' — a conta Google do robô não pode editar")
        raise AlunoPulado("ERRO", f"Google: a célula {celula} ficou '{gravado or 'vazia'}' em vez de '{texto}'")


# ==========================================================
# Rodada
# ==========================================================

def rodar(cfg, args):
    from playwright.sync_api import sync_playwright

    real = not args.simular
    alunos, letra, aba = ler_planilha(cfg)
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
    PARAR.unlink(missing_ok=True)

    print("=" * 70)
    print("ROBÔ CURSO PELO SGE —", "MODO REAL (vai colar na planilha)" if real else "SIMULAÇÃO (não cola)")
    print(f"Planilha: aba {aba} | coluna de destino: {letra} | alunos para processar: {len(alunos)}")
    print("=" * 70)

    linhas_csv = []
    if alunos:
        with sync_playwright() as p:
            sge_cfg = base.carregar_config()
            sge_contexto, sge_pagina = base.abrir_navegador(p)
            sge = base.Robo(sge_contexto, sge_pagina, sge_cfg, False, False)
            sge.rotulo_nome = cfg.get("rotulo_nome_sge") or "Nome"
            sge_pagina.on("dialog", sge._dialogo)
            contas = base.ContasSGE()
            print(f"Contas do SGE: {', '.join(c['usuario'] for c in contas.contas)}")
            base.ir(sge_pagina, sge_cfg["url_clientes"], sge.timeout)
            base.garantir_login(sge, contas, PARAR)

            google = google_pagina = None
            if real:
                try:
                    google, google_pagina = abrir_google(p, cfg)
                except SemPermissao as erro:
                    print(f"\nERRO — Google: {erro}\nNada foi colado. Rode de novo depois de resolver.")
                    sge_contexto.close()
                    return

            for posicao, aluno in enumerate(alunos, start=1):
                if pediu_parar():
                    print(f"\nPARADO A PEDIDO na linha {aluno['linha']}. Para continuar: --continuar")
                    break
                celula = f"{letra}{aluno['linha']}"
                print(f"\n[{posicao}/{len(alunos)}] {aluno['nome']} (linha {aluno['linha']})")
                curso, pagina_erro = "", sge.pagina

                def registrar(resultado, detalhe="", imagem=""):
                    CONTADORES[resultado] += 1
                    print(f"    {resultado}" + (f" — {detalhe}" if detalhe else "") + (f"\n      print: {imagem}" if imagem else ""))
                    linhas_csv.append([aluno["linha"], aluno["nome"], curso, resultado,
                                       (detalhe + (" | " + imagem if imagem else "")).strip(" |")])
                    salvar_relatorio(linhas_csv)

                try:
                    curso, detalhe = curso_com_contas(sge, contas, aluno)
                    antes = f"antes: '{aluno['atual']}'" if aluno["atual"] else "estava vazia"
                    texto = f"{celula} = {curso} ({antes})" + (f"; {detalhe}" if detalhe else "")
                    if real:
                        pagina_erro = google_pagina
                        escrever_celula(google_pagina, aba, celula, curso)
                        registrar("COLADO", texto)
                    else:
                        registrar("SIMULADO", texto)
                except SemPermissao as erro:
                    registrar("ERRO", f"Google: {erro}", base.tirar_print(pagina_erro, f"curso_linha{aluno['linha']}"))
                    print(f"\nPARADO na linha {aluno['linha']}: a planilha não aceita edição. Para continuar: --continuar")
                    break
                except AlunoPulado as pulo:
                    registrar(pulo.resultado, pulo.detalhe, base.tirar_print(pagina_erro, f"curso_linha{aluno['linha']}"))
                except Exception as erro:
                    if "has been closed" in str(erro):
                        print(f"\nO navegador foi fechado — parando na linha {aluno['linha']}. Para continuar: --continuar")
                        break
                    registrar("ERRO", base.erro_curto(erro), base.tirar_print(pagina_erro, f"curso_linha{aluno['linha']}"))

                if real:
                    proxima = seguinte.get(aluno["linha"])
                    salvar_json(PROGRESSO, {
                        "ultima_linha": aluno["linha"], "nome": aluno["nome"],
                        "proxima_linha": proxima, "concluido": proxima is None,
                        "data": f"{datetime.now():%Y-%m-%d %H:%M:%S}"})
                time.sleep(int(cfg.get("pausa_entre_alunos_ms") or 0) / 1000)

            for contexto in (sge_contexto, google):
                try:
                    if contexto:
                        contexto.close()
                except Exception:
                    pass

    salvar_relatorio(linhas_csv)
    print("\n" + "=" * 70)
    print(f"RELATORIO: {RELATORIO}")
    principal = f"COLADOS: {CONTADORES['COLADO']}" if real else f"SIMULADOS: {CONTADORES['SIMULADO']}"
    print(f"RESUMO CURSO — {principal} | NÃO ENCONTRADOS: {CONTADORES['NAO ENCONTRADO']} | "
          f"VÁRIOS NO SGE: {CONTADORES['VARIOS NO SGE']} | ERROS: {CONTADORES['ERRO']}")
    print("=" * 70)


def main():
    parser = argparse.ArgumentParser(description="Robô Curso pelo SGE — preencher o curso na planilha do Google")
    parser.add_argument("--listar", action="store_true", help="só mostra os nomes lidos da planilha")
    parser.add_argument("--simular", action="store_true", help="pesquisa no SGE, mas não cola na planilha")
    parser.add_argument("--limite", type=int)
    parser.add_argument("--linha", type=int, help="começa nesta linha da planilha")
    parser.add_argument("--continuar", action="store_true", help="continua da linha onde parou (curso_progresso.json)")
    args = parser.parse_args()

    cfg = carregar_config()
    if args.listar:
        alunos, letra, aba = ler_planilha(cfg)
        for aluno in alunos:
            atual = f" | agora: {aluno['atual']}" if aluno["atual"] else ""
            print(f"linha {aluno['linha']}: {aluno['nome']}{atual}")
        print(f"\nTotal: {len(alunos)} nomes (aba {aba}) | o curso vai para a coluna {letra}")
    else:
        rodar(cfg, args)


if __name__ == "__main__":
    try:
        main()
    except (FileNotFoundError, KeyError, ValueError) as erro:
        print(f"\nERRO: {erro}")
        sys.exit(1)
