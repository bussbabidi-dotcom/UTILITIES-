# -*- coding: utf-8 -*-
"""
ROBÔ DE BAIXA NO SGE — ALTERAÇÃO DE STATUS DOS ITENS DOS PEDIDOS

Para cada linha da planilha (código SGE, nome, novo status e, opcionalmente, o item):
  1. Clientes: pesquisa pelo CÓDIGO e abre o cliente (633.<código>)
  2. Do cliente, vai até os pedidos dele (caminho aprendido com você)
  3. Abre cada pedido 633.<código>.N e lê a tabela "Itens deste Pedido"
  4. Nos itens certos (nome da planilha, ou ENCADERNAÇÃO / ÁLBUM / BOOK, menos DIGITAL)
     clica em Ações > Alterar Status, escolhe o Novo Status e clica em Alterar

Os botões e campos são encontrados pelo TEXTO da tela, nunca por posição de clique.

MODO APRENDER: quando o robô não acha um botão, o caminho até os pedidos ou o
status pedido na planilha, ele lista no terminal o que está vendo na tela, numerado.
Você digita o número e ele grava a escolha no sge_config.json (não pergunta de novo).

Uso:
  python BAIXA.PY --listar                 mostra o que foi lido da planilha
  python BAIXA.PY                          altera o status de verdade
  python BAIXA.PY --simular                teste: abre tudo, mas clica em Cancelar no fim
  python BAIXA.PY --sge 73467              um aluno só
  python BAIXA.PY --limite 2               só os 2 primeiros
  python BAIXA.PY --esquecer               apaga o que foi aprendido (caminho, menu, status)
  python BAIXA.PY --mapear                 lê as telas e lista botões/campos em sge_mapa.txt
  python BAIXA.PY --listar-itens           itens e status da planilha (com contagem)
  python BAIXA.PY --link URL               planilha do Google (ou --planilha arquivo.xlsx)
  python BAIXA.PY --status "Enviado"       mesmo status para todos (no lugar da coluna)
  python BAIXA.PY --data 01/10/2026 --obs "texto"   data e observação da alteração
  python BAIXA.PY --mesmo-status           reenvia: reaplica o status mesmo se já estiver nele
  python BAIXA.PY --continuar              continua da linha salva (baixa_progresso.json)

A interface BAIXA_UI.pyw monta esses comandos.
"""

import argparse
import csv
import json
import re
import shutil
import sys
import tempfile
import time
import unicodedata
import urllib.request
from collections import Counter
from datetime import datetime
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass


PASTA = Path(__file__).resolve().parent
CONFIG = PASTA / "sge_config.json"
PERFIL = PASTA / "sge_perfil"
MAPA = PASTA / "sge_mapa.txt"
RELATORIO = PASTA / "baixa_sge.csv"
FEITAS = PASTA / "baixas_feitas.json"
PRINTS = PASTA / "sge_prints"
CREDENCIAIS = PASTA / "sge_credenciais.json"
PLANILHA_BAIXADA = PASTA / "baixa_planilha.xlsx"
PROGRESSO = PASTA / "baixa_progresso.json"
PARAR = PASTA / "baixa_parar.flag"

CONFIG_PADRAO = {
    "url_login": "https://sistema.sge.com.br/SCA/Forms/Login.aspx",
    "url_clientes": "https://sistema.sge.com.br/SGE/Forms/Cliente.aspx",
    "prefixo_codigo": "633.",
    "planilha": str(PASTA / "BAIXA_TESTE.xlsx"),
    # se preenchido, a planilha é baixada deste link do Google (tem prioridade sobre "planilha")
    "planilha_link": "",
    "aba": "",
    "linha_inicial": 2,
    "col_sge": "A",
    "col_nome": "B",
    "col_status": "C",
    "col_item": "D",
    "itens_incluir": ["ENCADERNA", "ALBUM", "BOOK"],
    "itens_excluir": ["DIGITAL"],
    # o álbum vendido separado tem prioridade sobre o plano/pacote
    "itens_prioritarios": ["ENCADERNACAO ALBUM"],
    # id da planilha do Google -> itens prioritários só quando essa planilha é usada
    "prioritarios_por_planilha": {},
    "pular_pedidos_com_texto": ["CANCELADO"],
    "ignorar_itens_com_status": ["ENTREGUE", "CANCELADO"],
    # novo status -> status que já estão além dele (o item não volta)
    "status_adiante": {"ENVIADO PARA APROVACAO": ["EM IMPRESSAO"]},
    "observacao": "",
    # data da alteração (dd/mm/aaaa); vazia = hoje
    "data_status": "",
    "conferir_nome": True,
    "timeout_ms": 20000,
    "pausa_entre_alunos_ms": 800,
    # aprendido com o usuário
    "texto_tela_cliente": ["Grade Financeira", "Mais Ações"],
    "caminho_pedidos": ["Mais Ações", "Ver Pedidos"],
    "menu_status": "Alterar Status",
    "mapa_status": {},
    "status_sge_conhecidos": [],
    "palavras_genericas": ["PACOTE", "PLANO", "STUDIO", "DAY", "KIT", "COMBO"],
    "apelidos_itens": {},
}

APRENDIDOS = ("caminho_pedidos", "menu_status", "mapa_status", "apelidos_itens")

CONTADORES = {
    "ALTERADO": 0, "REENVIADO": 0, "SIMULADO": 0, "JA NO STATUS": 0,
    "NAO ENCONTRADO": 0, "SEM ITEM": 0, "ERRO": 0,
}


# ==========================================================
# Config, textos e planilha
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


def guardar_aprendido(cfg, chave, valor):
    """Grava só o que foi aprendido, sem levar junto as opções temporárias da linha de comando."""

    cfg[chave] = valor
    disco = carregar_config()
    disco[chave] = valor
    salvar_config(disco)
    print(f"    (aprendido e salvo no sge_config.json: {chave})")


def normalizar(texto):
    texto = unicodedata.normalize("NFKD", str(texto or ""))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", texto).strip().upper()


def valor_celula(valor):
    if valor is None:
        return ""
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    return str(valor).strip()


def codigo_base(sge, cfg):
    """'633.73467' ou '73467' -> '73467'."""

    sge = str(sge).strip()
    prefixo = cfg.get("prefixo_codigo") or ""
    if prefixo and sge.startswith(prefixo):
        sge = sge[len(prefixo):]
    return sge


def indice_coluna(coluna):
    from openpyxl.utils import column_index_from_string

    coluna = str(coluna or "").strip()
    if not coluna:
        return None
    if coluna.isdigit():
        return int(coluna)
    return column_index_from_string(coluna.upper())


def abrir_planilha(caminho):
    from openpyxl import load_workbook

    try:
        return load_workbook(caminho, data_only=True, read_only=True)
    except PermissionError:
        copia = Path(tempfile.gettempdir()) / ("baixa_sge_" + Path(caminho).name)
        shutil.copy2(caminho, copia)
        return load_workbook(copia, data_only=True, read_only=True)


def link_exportacao(link):
    achado = re.search(r"/spreadsheets/d/([\w-]+)", link or "")
    if not achado:
        raise ValueError(f"Link da planilha inválido: {link}")
    gid = re.search(r"[#&?]gid=(\d+)", link)
    return (f"https://docs.google.com/spreadsheets/d/{achado.group(1)}/export?format=xlsx"
            + (f"&gid={gid.group(1)}" if gid else ""))


def baixar_planilha_link(link, destino):
    """Baixa a planilha do Google (export xlsx); se falhar, usa a última cópia baixada."""

    destino = Path(destino)
    temporario = destino.with_suffix(".baixando")
    try:
        urllib.request.urlretrieve(link_exportacao(link), temporario)
        if temporario.read_bytes()[:2] != b"PK":
            raise ValueError("o Google não devolveu a planilha (o link está compartilhado como 'qualquer pessoa com o link'?)")
        temporario.replace(destino)
    except Exception as erro:
        temporario.unlink(missing_ok=True)
        if not destino.exists():
            raise ValueError(f"Não consegui baixar a planilha: {erro}")
        print(f"(não consegui baixar a planilha agora — usando a última cópia: {erro})")
    return destino


def caminho_planilha(cfg):
    if cfg.get("planilha_link"):
        return baixar_planilha_link(cfg["planilha_link"], PLANILHA_BAIXADA)
    caminho = Path(cfg["planilha"])
    if not caminho.exists():
        raise FileNotFoundError(f"Planilha não encontrada: {caminho}")
    return caminho


def ler_planilha(cfg):
    livro = abrir_planilha(caminho_planilha(cfg))
    if cfg.get("aba") and cfg["aba"] not in livro.sheetnames:
        raise ValueError(f"A aba '{cfg['aba']}' não existe. Abas: {', '.join(livro.sheetnames)}")
    folha = livro[cfg["aba"]] if cfg.get("aba") else livro.active
    colunas = {chave: indice_coluna(cfg.get("col_" + chave)) for chave in ("sge", "nome", "status", "item")}
    inicio = int(cfg.get("linha_inicial") or 2)

    alunos = []
    for numero, linha in enumerate(folha.iter_rows(min_row=inicio, values_only=True), start=inicio):
        def pegar(chave):
            indice = colunas[chave]
            if not indice or indice > len(linha):
                return ""
            return valor_celula(linha[indice - 1])

        sge = pegar("sge")
        if not sge:
            continue
        alunos.append({
            "linha": numero,
            "sge": codigo_base(sge, cfg),
            "nome": pegar("nome"),
            "status": pegar("status"),
            "item": pegar("item"),
        })

    livro.close()
    return alunos


def carregar_feitas():
    try:
        return json.loads(FEITAS.read_text(encoding="utf-8"))
    except Exception:
        return {}


def salvar_feitas(feitas):
    FEITAS.write_text(json.dumps(feitas, indent=2, ensure_ascii=False), encoding="utf-8")


PALAVRAS_IGNORADAS = {"PLANO", "DE", "DA", "DO", "E", "-"}


def palavras(texto, genericas=()):
    ignorar = PALAVRAS_IGNORADAS | {normalizar(g) for g in genericas}
    return {p for p in re.split(r"[\s/()\-]+", normalizar(texto))
            if p and p not in ignorar and not re.fullmatch(r"\d+X\d+", p)}


def item_serve(nome_item, aluno, cfg):
    """
    Compara o item da planilha com o do SGE pelas palavras que identificam o pacote,
    sem tamanho e sem palavras genéricas:
      'Diamond Book 24X34'      -> 'PLANO DIAMOND BOOK'
      'Studio Day Berlim 24x34' -> 'PACOTE BERLIM'
    """

    nome = normalizar(nome_item)
    excluido = any(normalizar(p) in nome for p in cfg.get("itens_excluir", []))
    if aluno["item"]:
        desejado = normalizar(aluno["item"])
        apelido = cfg.get("apelidos_itens", {}).get(desejado)
        if apelido:
            return normalizar(apelido) == nome
        if excluido and not any(normalizar(p) in desejado for p in cfg.get("itens_excluir", [])):
            return False
        if desejado in nome:
            return True
        genericas = cfg.get("palavras_genericas", [])
        do_sge, da_planilha = palavras(nome_item, genericas), palavras(aluno["item"], genericas)
        return bool(do_sge and da_planilha) and (da_planilha <= do_sge or do_sge <= da_planilha)
    if excluido:
        return False
    return any(normalizar(p) in nome for p in cfg.get("itens_incluir", []))


def id_planilha(link):
    achado = re.search(r"/spreadsheets/d/([\w-]+)", link or "")
    return achado.group(1) if achado else ""


def lista_prioritarios(cfg):
    """Prioritários gerais + os da planilha do link em uso (itens que não aparecem nas colunas dela)."""

    lista = list(cfg.get("itens_prioritarios", []))
    da_planilha = cfg.get("prioritarios_por_planilha", {}).get(id_planilha(cfg.get("planilha_link")), [])
    return lista + [p for p in da_planilha if p not in lista]


def eh_prioritario(nome_item, cfg):
    nome = normalizar(nome_item)
    if any(normalizar(p) in nome for p in cfg.get("itens_excluir", []) if p):
        return False
    genericas = cfg.get("palavras_genericas", [])
    do_sge = palavras(nome_item, genericas)
    for prioritario in lista_prioritarios(cfg):
        if not prioritario:
            continue
        chave = palavras(prioritario, genericas)
        if normalizar(prioritario) in nome or (chave and chave <= do_sge):
            return True
    return False


def ja_adiante(status_item, alvo_status, cfg):
    atual = normalizar(status_item)
    adiante = cfg.get("status_adiante", {}).get(normalizar(alvo_status), [])
    return any(normalizar(s) and normalizar(s) in atual for s in adiante)


def status_ignorado(status_item, cfg):
    """Itens Entregue/Cancelado nunca são mexidos."""

    atual = normalizar(status_item)
    return any(normalizar(s) in atual for s in cfg.get("ignorar_itens_com_status", []) if s)


def status_desejado(cfg, status_planilha):
    """Texto do status no SGE: o aprendido (mapa_status) ou o da planilha."""

    return cfg.get("mapa_status", {}).get(normalizar(status_planilha), status_planilha)


def mesmo_status(atual, desejado):
    atual, desejado = normalizar(atual), normalizar(desejado)
    return bool(atual) and bool(desejado) and (atual == desejado or desejado in atual)


# ==========================================================
# Perguntas ao usuário (modo aprender)
# ==========================================================

def ler_resposta(texto, opcoes):
    return input(texto)


def terminal_interativo():
    return sys.stdin is not None and sys.stdin.isatty()


def perguntar(robo, titulo, opcoes, rodape="Enter = pular"):
    """Mostra as opções numeradas e devolve o índice escolhido (ou None)."""

    if not robo.interativo or not opcoes:
        return None
    print(f"\n    >>> {titulo}")
    for numero, opcao in enumerate(opcoes, start=1):
        print(f"       {numero:2d}. {opcao}")
    while True:
        resposta = ler_resposta(f"    Digite o número ({rodape}): ", opcoes).strip()
        if not resposta:
            return None
        if resposta.isdigit() and 1 <= int(resposta) <= len(opcoes):
            return int(resposta) - 1
        print("    Número inválido.")


# ==========================================================
# Navegador e localização pelo texto
# ==========================================================

class AlunoPulado(Exception):
    def __init__(self, resultado, detalhe=""):
        super().__init__(detalhe)
        self.resultado = resultado
        self.detalhe = detalhe


class Robo:
    def __init__(self, contexto, pagina, cfg, real, interativo):
        self.contexto = contexto
        self.pagina = pagina
        self.cfg = cfg
        self.real = real
        self.interativo = interativo
        self.timeout = int(cfg["timeout_ms"])
        self.contas = None
        # reaplicar o mesmo status (reenvio com nova data/observação)
        self.reenviar = False
        self.reenviados = set()

    def seguir_nova_aba(self):
        """Se um clique abriu outra aba, passa a trabalhar nela."""

        abertas = [p for p in self.contexto.pages if not p.is_closed()]
        if abertas and abertas[-1] is not self.pagina:
            self.pagina = abertas[-1]
            self.pagina.bring_to_front()
            self.pagina.on("dialog", self._dialogo)

    def _dialogo(self, dialogo):
        if self.real:
            dialogo.accept()
        else:
            dialogo.dismiss()


def abrir_navegador(playwright):
    """Abre o perfil do SGE; se ele já estiver aberto por outro robô, usa sge_perfil_2, _3..."""

    for numero in range(1, 5):
        perfil = PERFIL if numero == 1 else PERFIL.with_name(f"{PERFIL.name}_{numero}")
        perfil.mkdir(exist_ok=True)
        try:
            contexto = playwright.chromium.launch_persistent_context(
                str(perfil), headless=False, no_viewport=True, args=["--start-maximized"],
            )
        except Exception:
            if numero == 4:
                raise
            print(f"(o navegador {perfil.name} já está aberto por outro robô — usando outro perfil)")
            continue
        pagina = contexto.pages[0] if contexto.pages else contexto.new_page()
        return contexto, pagina


def css_texto(texto):
    return texto.replace("\\", "\\\\").replace('"', '\\"')


JS_MARCAR_CAMPO = r"""
([rotulo, marca]) => {
  const norm = s => (s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ').trim().toUpperCase();
  const vis = e => { const r = e.getBoundingClientRect(); return r.width > 0 && r.height > 0 && getComputedStyle(e).visibility !== 'hidden'; };
  const alvo = norm(rotulo);
  const rotulos = [...document.querySelectorAll('label, span, td, th, div, b, strong, p')]
    .filter(e => vis(e) && norm(e.innerText) === alvo);
  const campos = [...document.querySelectorAll(
    'input:not([type=hidden]):not([type=button]):not([type=submit]):not([type=checkbox]):not([type=radio]), select, textarea')]
    .filter(c => !c.disabled && !c.readOnly);
  for (const r of rotulos) {
    for (const c of campos) {
      if ((r.compareDocumentPosition(c) & Node.DOCUMENT_POSITION_FOLLOWING) && (c.tagName === 'SELECT' || vis(c))) {
        c.setAttribute('data-robo', marca);
        return true;
      }
    }
  }
  return false;
}
"""

JS_CLICAVEIS = r"""
() => {
  const vis = e => { const r = e.getBoundingClientRect(); const s = getComputedStyle(e);
                     return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none'; };
  const txt = s => (s || '').replace(/\s+/g, ' ').trim().slice(0, 80);
  document.querySelectorAll('[data-robo-opc]').forEach(e => e.removeAttribute('data-robo-opc'));
  const els = [...document.querySelectorAll('a, button, input[type=button], input[type=submit], [role=menuitem], [role=button]')]
    .filter(e => vis(e) && !e.closest('.main-sidebar, aside.sidebar'));
  const saida = [], vistos = new Set();
  els.forEach(e => {
    const t = txt(e.innerText || e.value || e.title || e.getAttribute('aria-label') || e.getAttribute('data-original-title'));
    if (!t || vistos.has(t)) return;
    vistos.add(t);
    e.setAttribute('data-robo-opc', String(saida.length));
    saida.push(t);
  });
  return saida.slice(0, 80);
}
"""

_marcas = [0]


def candidatos(frame, tipo, texto):
    if tipo == "botao":
        return [
            frame.get_by_role("button", name=texto, exact=True),
            frame.get_by_role("link", name=texto, exact=True),
            frame.get_by_role("menuitem", name=texto, exact=True),
            frame.locator(f'input[type="submit"][value="{css_texto(texto)}"], input[type="button"][value="{css_texto(texto)}"]'),
            frame.get_by_text(texto, exact=True),
        ]
    if tipo == "botao_parecido":
        return [
            frame.get_by_role("button", name=texto),
            frame.get_by_role("link", name=texto),
            frame.get_by_role("menuitem", name=texto),
            frame.get_by_text(texto),
        ]
    if tipo == "campo":
        return [
            frame.get_by_label(texto, exact=True),
            frame.get_by_placeholder(texto, exact=True),
        ]
    return [frame.get_by_text(texto, exact=True), frame.get_by_text(texto)]


def localizar(pagina, tipo, texto, timeout_ms):
    """Procura em todos os frames até achar um elemento visível ou estourar o tempo."""

    limite = time.time() + timeout_ms / 1000
    while True:
        for frame in pagina.frames:
            try:
                for loc in candidatos(frame, tipo, texto):
                    for i in range(min(loc.count(), 20)):
                        item = loc.nth(i)
                        if item.is_visible():
                            return item
                if tipo == "campo":
                    _marcas[0] += 1
                    marca = f"m{_marcas[0]}"
                    if frame.evaluate(JS_MARCAR_CAMPO, [texto, marca]):
                        return frame.locator(f'[data-robo="{marca}"]').first
            except Exception:
                continue
        if time.time() >= limite:
            return None
        time.sleep(0.3)


def clicaveis(pagina, timeout_ms=8000):
    """Lista (frame, índice, texto) de tudo que dá para clicar na tela."""

    limite = time.time() + timeout_ms / 1000
    while True:
        lista = []
        for frame in pagina.frames:
            try:
                for indice, texto in enumerate(frame.evaluate(JS_CLICAVEIS)):
                    lista.append((frame, indice, texto))
            except Exception:
                continue
        if lista or time.time() >= limite:
            return lista
        time.sleep(0.5)


def na_tela_de_login(pagina):
    try:
        return "LOGIN" in pagina.url.upper() or pagina.locator("input[type=password]:visible").count() > 0
    except Exception:
        return False


def fazer_login(pagina, timeout_ms):
    """Entra no SGE com o usuário/senha do sge_credenciais.json. Devolve True se entrou."""

    try:
        cred = json.loads(CREDENCIAIS.read_text(encoding="utf-8"))
    except Exception:
        return False
    if not cred.get("usuario") or not cred.get("senha"):
        return False

    senha = pagina.locator("input[type=password]:visible").first
    try:
        senha.wait_for(timeout=timeout_ms)
    except Exception:
        return False
    usuario = pagina.locator(
        "input[type=email]:visible, input[type=text]:visible, input:not([type]):visible").first
    usuario.fill(cred["usuario"])
    senha.fill(cred["senha"])

    entrar = localizar(pagina, "botao", "Entrar", 5000)
    if entrar is not None:
        entrar.click()
    else:
        senha.press("Enter")

    limite = time.time() + timeout_ms / 1000
    while time.time() < limite:
        pagina.wait_for_timeout(500)
        # "Já existe uma sessão aberta para este usuário... Deseja prosseguir?"
        if "SESSAO ABERTA" in normalizar(texto_da_tela(pagina)):
            sim = pagina.get_by_role("button", name="Sim", exact=True)
            if not sim.count():
                sim = pagina.get_by_text("Sim", exact=True)
            try:
                sim.first.click(timeout=3000)
            except Exception:
                pass
            continue
        if not na_tela_de_login(pagina):
            esperar_carregar(pagina, timeout_ms)
            print("    (login feito automaticamente)")
            return True
    return False


class ContasSGE:
    """
    Contas do SGE em sge_credenciais.json ({"contas": [{"usuario", "senha"}, ...]}).
    As contas ficam logadas o tempo todo na empresa: ao entrar sempre derruba a sessão aberta ("Sim").
    Se a sessão cair 2 vezes, alguém está usando a conta — troca para a próxima (dá voltas na lista).
    """

    QUEDAS_PARA_TROCAR = 2

    def __init__(self):
        try:
            dados = json.loads(CREDENCIAIS.read_text(encoding="utf-8"))
        except Exception:
            dados = {}
        contas = [c for c in dados.get("contas", []) if c.get("usuario") and c.get("senha")]
        if not contas and dados.get("usuario"):
            contas = [{"usuario": dados["usuario"], "senha": dados.get("senha", "")}]
        if not contas:
            raise ValueError("Nenhuma conta do SGE em sge_credenciais.json")
        self.contas = contas
        self.atual = 0
        self.quedas = 0
        self.logado = False

    @property
    def usuario(self):
        return self.contas[self.atual]["usuario"]

    def proxima(self):
        self.atual = (self.atual + 1) % len(self.contas)
        self.quedas = 0

    def caiu(self):
        """A sessão caiu (o SGE voltou para o login no meio do trabalho)."""

        self.quedas += 1
        print(f"    (SGE: a sessão da conta {self.usuario} caiu — {self.quedas}ª vez)")
        if self.quedas >= self.QUEDAS_PARA_TROCAR and len(self.contas) > 1:
            anterior = self.usuario
            self.proxima()
            print(f"    (SGE: caiu {self.QUEDAS_PARA_TROCAR} vezes — trocando {anterior} por {self.usuario})")


def login_conta(pagina, conta, timeout_ms):
    """Entra com uma conta (clicando 'Sim' se já houver sessão aberta). Devolve True se entrou."""

    senha = pagina.locator("input[type=password]:visible").first
    try:
        senha.wait_for(timeout=timeout_ms)
    except Exception:
        return False
    pagina.locator("input[type=email]:visible, input[type=text]:visible, input:not([type]):visible").first.fill(conta["usuario"])
    senha.fill(conta["senha"])
    entrar = localizar(pagina, "botao", "Entrar", 5000)
    if entrar is not None:
        entrar.click()
    else:
        senha.press("Enter")

    limite = time.time() + timeout_ms / 1000
    while time.time() < limite:
        pagina.wait_for_timeout(500)
        if "SESSAO ABERTA" in normalizar(texto_da_tela(pagina)):
            botao = pagina.get_by_role("button", name="Sim", exact=True)
            if not botao.count():
                botao = pagina.get_by_text("Sim", exact=True)
            try:
                botao.first.click(timeout=3000)
            except Exception:
                pass
            continue
        if not na_tela_de_login(pagina):
            esperar_carregar(pagina, timeout_ms)
            return True
    return False


def entrar_sge(robo, contas, parar=None):
    """Login com a conta atual; se ela não entrar, tenta as próximas (dá voltas na lista)."""

    pagina, url = robo.pagina, robo.cfg["url_clientes"]
    tentadas = 0
    while True:
        if not na_tela_de_login(pagina):
            ir(pagina, url, robo.timeout)
            if not na_tela_de_login(pagina):
                contas.logado = True
                return
        if login_conta(pagina, contas.contas[contas.atual], robo.timeout):
            print(f"    (SGE: entrei com a conta {contas.usuario})")
            contas.logado = True
            ir(pagina, url, robo.timeout)
            return
        print(f"    (SGE: a conta {contas.usuario} não entrou (confira a senha) — tentando a próxima)")
        contas.proxima()
        tentadas += 1
        if tentadas >= len(contas.contas):
            print("    (SGE: nenhuma conta conseguiu entrar — espero 60s e tento de novo)")
            for _ in range(60):
                if parar is not None and parar.exists():
                    raise AlunoPulado("ERRO", "SGE: nenhuma conta entrou e pediram para parar")
                time.sleep(1)
            tentadas = 0
        ir(pagina, url, robo.timeout)


def garantir_login(robo, contas, parar=None):
    """Se o SGE voltou para o login, conta a queda (troca de conta na 2ª) e entra de novo. Devolve True se caiu."""

    if not na_tela_de_login(robo.pagina):
        contas.logado = True
        return False
    if contas.logado:
        contas.caiu()
    entrar_sge(robo, contas, parar)
    return True


def ir(pagina, url, timeout_ms):
    """goto que aguenta o SGE redirecionando sozinho (ex.: depois do login)."""

    for tentativa in range(3):
        try:
            pagina.goto(url, timeout=max(timeout_ms, 30000))
            esperar_carregar(pagina, timeout_ms)
            return
        except Exception:
            if tentativa == 2:
                raise
            pagina.wait_for_timeout(2000)


def texto_da_tela(pagina):
    partes = []
    for frame in pagina.frames:
        try:
            partes.append(frame.locator("body").inner_text(timeout=3000))
        except Exception:
            pass
    return normalizar(" ".join(partes))


def esperar_carregar(pagina, timeout_ms):
    try:
        pagina.wait_for_load_state("networkidle", timeout=timeout_ms)
    except Exception:
        pass


def destacar(elemento):
    try:
        elemento.evaluate("e => { e.style.outline = '4px solid red'; e.scrollIntoView({block: 'center'}); }")
    except Exception:
        pass


def tirar_print(pagina, nome):
    PRINTS.mkdir(exist_ok=True)
    arquivo = PRINTS / f"{datetime.now():%Y%m%d_%H%M%S}_{re.sub(r'[^0-9A-Za-z_-]', '_', nome)}.png"
    try:
        pagina.screenshot(path=str(arquivo), full_page=True)
        return str(arquivo)
    except Exception:
        return ""


def preencher(elemento, valor):
    elemento.click()
    elemento.fill(valor)
    if (elemento.input_value() or "").strip() != valor:
        elemento.press("Control+A")
        elemento.type(valor, delay=40)


# ==========================================================
# Cliente -> pedidos
# ==========================================================

JS_PEDIDOS = r"""
(padrao) => {
  const re = new RegExp(padrao);
  const vistos = new Set(), saida = [];
  document.querySelectorAll('a, td, span').forEach(e => {
    const t = (e.innerText || '').trim();
    if (re.test(t) && !vistos.has(t)) {
      vistos.add(t);
      const linha = e.closest('tr');
      saida.push({numero: t, linha: linha ? linha.innerText.replace(/\s+/g, ' ') : ''});
    }
  });
  return saida;
}
"""


def pedidos_na_tela(robo, aluno):
    padrao = "^" + re.escape(robo.cfg["prefixo_codigo"] + aluno["sge"]) + r"\.\d+$"
    for frame in robo.pagina.frames:
        try:
            achados = frame.evaluate(JS_PEDIDOS, padrao)
        except Exception:
            continue
        if achados:
            return achados
    return []


def esperar_pedidos(robo, aluno, timeout_ms):
    limite = time.time() + timeout_ms / 1000
    while True:
        robo.seguir_nova_aba()
        achados = pedidos_na_tela(robo, aluno)
        if achados or time.time() >= limite:
            return achados
        time.sleep(0.5)


def abrir_cliente(robo, aluno):
    """Clientes -> Código -> Pesquisar -> clica no 633.<código>."""

    pagina = robo.pagina
    ir(pagina, robo.cfg["url_clientes"], robo.timeout)
    if robo.contas is not None:
        garantir_login(robo, robo.contas, PARAR)
    elif na_tela_de_login(pagina):
        if not fazer_login(pagina, robo.timeout):
            raise AlunoPulado("ERRO", "O SGE pediu login e não consegui entrar (confira sge_credenciais.json)")
        ir(pagina, robo.cfg["url_clientes"], robo.timeout)

    campo = localizar(pagina, "campo", "Código", robo.timeout)
    if campo is None:
        raise AlunoPulado("ERRO", "Não achei o campo 'Código' na tela de Clientes")
    preencher(campo, aluno["sge"])

    botao = localizar(pagina, "botao", "Pesquisar", robo.timeout)
    if botao is None:
        raise AlunoPulado("ERRO", "Não achei o botão Pesquisar em Clientes")
    botao.click()
    esperar_carregar(pagina, robo.timeout)

    codigo = robo.cfg["prefixo_codigo"] + aluno["sge"]
    link = localizar(pagina, "botao", codigo, robo.timeout)
    if link is None:
        raise AlunoPulado("NAO ENCONTRADO", f"Cliente {codigo} não apareceu na pesquisa")
    link.click()
    esperar_carregar(pagina, robo.timeout)

    marcas = [normalizar(t) for t in robo.cfg.get("texto_tela_cliente", []) if t]
    limite = time.time() + robo.timeout / 1000
    while marcas:
        robo.seguir_nova_aba()
        tela = texto_da_tela(robo.pagina)
        if any(m in tela for m in marcas):
            break
        if time.time() >= limite:
            raise AlunoPulado("ERRO", f"Cliquei em {codigo}, mas a Grade Financeira não abriu")
        time.sleep(0.5)
    pagina = robo.pagina
    esperar_carregar(pagina, robo.timeout)

    if robo.cfg.get("conferir_nome") and aluno["nome"]:
        limite = time.time() + robo.timeout / 1000
        while normalizar(aluno["nome"]) not in texto_da_tela(pagina):
            if time.time() >= limite:
                raise AlunoPulado("ERRO", "O nome da planilha não aparece na tela do cliente — nada alterado")
            time.sleep(0.5)


def ir_para_pedidos(robo, aluno):
    """Abre o cliente e segue o caminho (aprendido) até a lista de pedidos dele."""

    abrir_cliente(robo, aluno)
    achados = pedidos_na_tela(robo, aluno)
    if achados:
        return achados

    for texto in robo.cfg.get("caminho_pedidos", []):
        alvo = localizar(robo.pagina, "botao", texto, 3000) or localizar(robo.pagina, "botao_parecido", texto, robo.timeout)
        if alvo is None:
            print(f"    Não achei '{texto}' na tela. Vou perguntar o caminho.")
            break
        alvo.click()
        esperar_carregar(robo.pagina, robo.timeout)
        achados = esperar_pedidos(robo, aluno, 2500)
        if achados:
            return achados
    else:
        if robo.cfg.get("caminho_pedidos"):
            achados = esperar_pedidos(robo, aluno, robo.timeout)
            if achados:
                return achados
            tela = texto_da_tela(robo.pagina)
            if "SUA PESQUISA ENCONTROU" in tela or "NENHUM" in tela or not robo.interativo:
                return []

    return aprender_caminho(robo, aluno)


def aprender_caminho(robo, aluno):
    """Pergunta, clique a clique, como chegar aos pedidos do cliente; grava se der certo."""

    if not robo.interativo:
        opcoes = [t for _, _, t in clicaveis(robo.pagina)]
        raise AlunoPulado("ERRO", "Não sei chegar aos pedidos. Opções na tela: " + " | ".join(opcoes[:25]))

    abrir_cliente(robo, aluno)
    caminho = []
    for _ in range(6):
        opcoes = clicaveis(robo.pagina)
        escolha = perguntar(
            robo,
            f"Estou na tela do cliente {aluno['sge']}. Onde clico para chegar aos PEDIDOS dele?",
            [t for _, _, t in opcoes],
            "Enter = este cliente não tem pedidos",
        )
        if escolha is None:
            return []
        frame, indice, texto = opcoes[escolha]
        frame.locator(f'[data-robo-opc="{indice}"]').first.click()
        caminho.append(texto)
        esperar_carregar(robo.pagina, robo.timeout)
        achados = esperar_pedidos(robo, aluno, 4000)
        if achados:
            guardar_aprendido(robo.cfg, "caminho_pedidos", caminho)
            return achados
    raise AlunoPulado("ERRO", "Depois de 6 cliques ainda não vi os pedidos 633.<código>.N")


def abrir_pedido(robo, numero):
    link = localizar(robo.pagina, "botao", numero, robo.timeout)
    if link is None:
        raise AlunoPulado("ERRO", f"Não achei o pedido {numero} na lista")
    link.click()
    esperar_carregar(robo.pagina, robo.timeout)
    limite = time.time() + robo.timeout / 1000
    while "ITENS DESTE PEDIDO" not in texto_da_tela(robo.pagina):
        robo.seguir_nova_aba()
        if time.time() >= limite:
            raise AlunoPulado("ERRO", f"O pedido {numero} não abriu (sem 'Itens deste Pedido')")
        time.sleep(0.4)


# ==========================================================
# Itens do pedido e alteração de status
# ==========================================================

JS_ITENS = r"""
() => {
  const norm = s => (s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ').trim().toUpperCase();
  const cab = [...document.querySelectorAll('th, td')].find(e => norm(e.innerText) === 'STATUS ITEM');
  if (!cab) return null;
  const tabela = cab.closest('table');
  const titulos = [...cab.parentElement.children].map(c => norm(c.innerText));
  const cItem = titulos.findIndex(t => t === 'ITEM'), cStatus = titulos.findIndex(t => t === 'STATUS ITEM');
  const cCod = titulos.findIndex(t => t.startsWith('COD'));
  const saida = [];
  [...tabela.querySelectorAll('tr')].forEach((tr, i) => {
    if (tr === cab.parentElement) return;
    const cel = tr.children;
    if (cel.length < titulos.length || cel.length <= Math.max(cItem, cStatus)) return;
    tr.setAttribute('data-robo-item', String(i));
    saida.push({linha: String(i), cod: cCod >= 0 ? cel[cCod].innerText.trim() : '',
                item: cel[cItem].innerText.trim(), status: cel[cStatus].innerText.trim()});
  });
  return saida;
}
"""

JS_OPCOES = "s => [...s.options].map(o => o.text.trim())"

JS_ESCOLHER = r"""
(s, texto) => {
  const op = [...s.options].find(o => o.text.trim() === texto);
  if (!op) return false;
  s.value = op.value;
  if (window.jQuery) { window.jQuery(s).val(op.value).trigger('change'); }
  else { s.dispatchEvent(new Event('input', {bubbles: true})); s.dispatchEvent(new Event('change', {bubbles: true})); }
  return true;
}
"""


def ler_itens(pagina):
    for frame in pagina.frames:
        try:
            itens = frame.evaluate(JS_ITENS)
        except Exception:
            continue
        if itens is not None:
            return frame, itens
    return pagina.main_frame, None


JS_MARCAR_MODAL = r"""
() => {
  const norm = s => (s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ').trim().toUpperCase();
  const vis = e => { const r = e.getBoundingClientRect(); const s = getComputedStyle(e);
                     return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none'; };
  document.querySelectorAll('[data-robo-modal]').forEach(e => e.removeAttribute('data-robo-modal'));
  const titulos = [...document.querySelectorAll('h1, h2, h3, h4, h5, h6, .modal-title, span, div, p, b, strong')]
    .filter(e => vis(e) && norm(e.innerText).startsWith('ALTERACAO DE STATUS') && norm(e.innerText).length < 160);
  for (const t of titulos) {
    let caixa = t;
    for (let i = 0; i < 12 && caixa; i++) {
      if (caixa.querySelector('select') && [...caixa.querySelectorAll('button, a, input[type=submit], input[type=button]')]
            .some(b => norm(b.innerText || b.value) === 'ALTERAR')) {
        caixa.setAttribute('data-robo-modal', '1');
        return true;
      }
      caixa = caixa.parentElement;
    }
  }
  return false;
}
"""


def modal_aberto(pagina, timeout_ms):
    """Acha a janela 'Alteração de Status' pelo título, subindo até o bloco que tem a lista e o botão Alterar."""

    limite = time.time() + timeout_ms / 1000
    while time.time() < limite:
        for frame in pagina.frames:
            try:
                if frame.evaluate(JS_MARCAR_MODAL):
                    return frame.locator('[data-robo-modal="1"]').first
            except Exception:
                continue
        time.sleep(0.4)
    return None


def mesmo_item(nome_tabela, texto_janela):
    """'PLANO ESSENCIAL BOOK 1' (com número de nota) casa com o título da janela."""

    nome = {p for p in palavras(nome_tabela) if not p.isdigit()}
    return bool(nome) and nome <= palavras(texto_janela)


def fechar_modal(modal):
    for nome in ("Cancelar", "Fechar", "×"):
        for botao in (modal.get_by_role("button", name=nome, exact=True),
                      modal.get_by_role("link", name=nome, exact=True),
                      modal.get_by_text(nome, exact=True)):
            if botao.count() and botao.first.is_visible():
                botao.first.click()
                return
    try:
        modal.page.keyboard.press("Escape")
    except Exception:
        pass


def achar_opcao_menu(robo, frame, linha):
    """Opção do menu de Ações que abre a alteração de status do ITEM."""

    desejado = normalizar(robo.cfg.get("menu_status") or "Alterar Status")
    parecidas = []
    for escopo in (linha, frame):
        itens = escopo.locator(".dropdown-menu a, .dropdown-menu li, .dropdown-menu button, [role=menuitem], a, button")
        for i in range(min(itens.count(), 80)):
            candidato = itens.nth(i)
            try:
                if not candidato.is_visible():
                    continue
                texto = normalizar(candidato.inner_text())
            except Exception:
                continue
            if texto == desejado:
                return candidato
            if "STATUS" in texto and "PEDIDO" not in texto:
                parecidas.append(candidato)
        if len(parecidas) == 1:
            return parecidas[0]

    opcoes = [o for o in clicaveis(robo.pagina) if "PEDIDO" not in normalizar(o[2])]
    escolha = perguntar(robo, "Abri o menu de Ações do item. Qual opção ALTERA O STATUS do item?",
                        [t for _, _, t in opcoes])
    if escolha is None:
        raise AlunoPulado("ERRO", "Menu de Ações sem opção de status. Opções na tela: "
                          + " | ".join(t for _, _, t in opcoes[:25]))
    frame_opc, indice, texto = opcoes[escolha]
    guardar_aprendido(robo.cfg, "menu_status", texto)
    return frame_opc.locator(f'[data-robo-opc="{indice}"]').first


def escolher_status(robo, opcoes, status_planilha):
    """Opção do Novo Status: aprendida, igual, ou parecida; se não achar, pergunta e aprende."""

    desejado = status_desejado(robo.cfg, status_planilha)
    alvo = normalizar(desejado)
    exatas = [o for o in opcoes if normalizar(o) == alvo]
    if exatas:
        return exatas[0]
    parecidas = [o for o in opcoes if alvo and alvo in normalizar(o)]
    if len(parecidas) == 1:
        return parecidas[0]

    validas = [o for o in opcoes if o and not o.startswith("(")]
    escolha = perguntar(robo, f"Não achei o status '{status_planilha}' no SGE. Qual destes é o certo?", validas)
    if escolha is None:
        raise AlunoPulado("ERRO", f"Status '{status_planilha}' não encontrado. Status do SGE: " + " | ".join(validas))
    mapa = dict(robo.cfg.get("mapa_status", {}))
    mapa[normalizar(status_planilha)] = validas[escolha]
    guardar_aprendido(robo.cfg, "mapa_status", mapa)
    return validas[escolha]


JS_SELECIONADA = "s => s.selectedIndex >= 0 ? s.options[s.selectedIndex].text.trim() : ''"

JS_CARREGANDO = r"""
() => [...document.querySelectorAll('[class*=loading], [class*=spinner], [class*=progress], [class*=blockUI], [id*=UpdateProgress], [class*=loader]')]
  .some(e => { const r = e.getBoundingClientRect(); const s = getComputedStyle(e);
               return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none' && s.opacity !== '0'; })
"""


def esperar_sge_parar(pagina, timeout_ms):
    """Espera a 'rodinha' de carregamento do SGE sumir."""

    esperar_carregar(pagina, timeout_ms)
    limite = time.time() + min(timeout_ms, 10000) / 1000
    while time.time() < limite:
        ocupado = False
        for frame in pagina.frames:
            try:
                if frame.evaluate(JS_CARREGANDO):
                    ocupado = True
                    break
            except Exception:
                continue
        if not ocupado:
            return
        time.sleep(0.4)


def aprender_status_sge(cfg, opcoes):
    """Guarda os status que o SGE oferece (a UI mostra essa lista)."""

    conhecidos = list(cfg.get("status_sge_conhecidos", []))
    novos = [o for o in opcoes if o and not o.startswith("(") and o not in conhecidos]
    if novos:
        guardar_aprendido(cfg, "status_sge_conhecidos", conhecidos + novos)


JS_CAMPO = r"""
(modal, [nome, padrao]) => {
  const norm = s => (s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ').trim().toUpperCase();
  const vis = e => { const r = e.getBoundingClientRect(); const s = getComputedStyle(e);
                     return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none'; };
  const campos = [...modal.querySelectorAll('input, textarea')].filter(e =>
    vis(e) && !['hidden', 'button', 'submit', 'checkbox', 'radio', 'file'].includes((e.type || '').toLowerCase()));
  const rotulo = e => {
    let t = [e.id, e.name, e.placeholder, e.title].join(' ');
    if (e.id) { const l = modal.querySelector('label[for="' + e.id + '"]'); if (l) t += ' ' + l.innerText; }
    let irmao = e.previousElementSibling;
    if (irmao) t += ' ' + (irmao.innerText || '');
    const pai = e.parentElement;
    if (pai && pai.innerText.length < 80) t += ' ' + pai.innerText;
    if (pai && pai.previousElementSibling) t += ' ' + (pai.previousElementSibling.innerText || '').slice(0, 60);
    return norm(t);
  };
  const re = new RegExp(padrao);
  const achado = (nome === 'data' && campos.find(e => (e.type || '').toLowerCase() === 'date'))
              || campos.find(e => re.test(rotulo(e)))
              || (nome === 'obs' && campos.find(e => e.tagName === 'TEXTAREA'));
  if (!achado) return null;
  modal.querySelectorAll('[data-robo-campo="' + nome + '"]').forEach(e => e.removeAttribute('data-robo-campo'));
  achado.setAttribute('data-robo-campo', nome);
  return (achado.type || 'text').toLowerCase();
}
"""


def campo_do_modal(modal, nome, padrao):
    """Acha um campo da janela pelo rótulo (id, name, placeholder, label ou texto ao lado). Devolve (locator, tipo)."""

    tipo = modal.evaluate(JS_CAMPO, [nome, padrao])
    if tipo is None:
        return None, None
    return modal.locator(f'[data-robo-campo="{nome}"]').first, tipo


def data_da_rodada(cfg):
    texto = (cfg.get("data_status") or "").strip()
    if not texto:
        return datetime.now().strftime("%d/%m/%Y")
    try:
        return datetime.strptime(texto, "%d/%m/%Y").strftime("%d/%m/%Y")
    except ValueError:
        raise ValueError(f"Data inválida: '{texto}' (use dd/mm/aaaa)")


def campos_do_modal(modal):
    try:
        return modal.evaluate(r"""m => [...m.querySelectorAll('input, textarea, select')]
          .filter(e => e.type !== 'hidden')
          .map(e => e.tagName.toLowerCase() + ' id="' + (e.id || '') + '" name="' + (e.name || '') + '" tipo=' + (e.type || ''))""")
    except Exception:
        return []


def preencher_data_obs(robo, modal):
    """Etapa depois de escolher o status: data da alteração e observação. Devolve o texto para o relatório."""

    pagina, cfg = robo.pagina, robo.cfg
    data = data_da_rodada(cfg)
    esperar_sge_parar(pagina, robo.timeout)
    campo, tipo = campo_do_modal(modal, "data", r"DATA|DT_|DTA")
    if campo is None:
        imagem = tirar_print(pagina, "janela_status_sem_data")
        raise AlunoPulado("ERRO", "Não achei o campo de DATA na janela 'Alteração de Status' — nada alterado. "
                                  "Campos da janela: " + " | ".join(campos_do_modal(modal)) + f" | print: {imagem}")
    if tipo == "date":
        campo.fill(datetime.strptime(data, "%d/%m/%Y").strftime("%Y-%m-%d"))
    else:
        preencher(campo, data)
    campo.press("Tab")
    pagina.wait_for_timeout(400)
    esperar_sge_parar(pagina, robo.timeout)
    lido = re.sub(r"\D", "", campo.input_value() or "")
    if lido not in (re.sub(r"\D", "", data), datetime.strptime(data, "%d/%m/%Y").strftime("%Y%m%d")):
        raise AlunoPulado("ERRO", f"Coloquei a data {data}, mas o campo ficou '{campo.input_value()}' — nada alterado")

    texto = f" | data {data}"
    observacao = (cfg.get("observacao") or "").strip()
    if observacao:
        obs, _ = campo_do_modal(modal, "obs", r"OBSERV|OBS\b")
        if obs is None:
            raise AlunoPulado("ERRO", "Não achei o campo de Observação na janela — nada alterado. "
                                      "Campos da janela: " + " | ".join(campos_do_modal(modal)))
        obs.fill(observacao)
        if (obs.input_value() or "").strip() != observacao:
            raise AlunoPulado("ERRO", "Escrevi a observação, mas o campo não ficou com o texto — nada alterado")
        texto += f" | obs: {observacao[:60]}"
    return texto


def alterar_item(robo, frame, item, aluno):
    pagina = robo.pagina
    linha = frame.locator(f'[data-robo-item="{item["linha"]}"]')

    acoes = linha.locator("td").last.locator("a, button")
    if acoes.count() == 0:
        raise AlunoPulado("ERRO", "Não achei o botão de Ações do item")

    modal = None
    for tentativa in range(2):
        esperar_sge_parar(pagina, robo.timeout)
        acoes.first.click()
        pagina.wait_for_timeout(800)
        achar_opcao_menu(robo, frame, linha).click()
        modal = modal_aberto(pagina, robo.timeout)
        if modal is not None:
            break
        pagina.keyboard.press("Escape")
    if modal is None:
        raise AlunoPulado("ERRO", "A janela 'Alteração de Status' não abriu (tentei 2 vezes)")
    esperar_sge_parar(pagina, robo.timeout)
    if not mesmo_item(item["item"], modal.inner_text()):
        fechar_modal(modal)
        raise AlunoPulado("ERRO", f"A janela aberta não é do item {item['item']} — nada alterado")

    opcoes = modal.locator("select").first.evaluate(JS_OPCOES)
    aprender_status_sge(robo.cfg, opcoes)
    try:
        escolhida = escolher_status(robo, opcoes, aluno["status"])
    except AlunoPulado:
        fechar_modal(modal)
        raise
    reenvio = mesmo_status(item["status"], escolhida)
    if reenvio and not robo.reenviar:
        fechar_modal(modal)
        return "JA NO STATUS", item["status"]

    # o SGE recarrega a janela quando o status é escolhido: espera e confere (até 3 vezes)
    for tentativa in range(3):
        modal.locator("select").first.evaluate(JS_ESCOLHER, escolhida)
        pagina.wait_for_timeout(1000)
        esperar_sge_parar(pagina, robo.timeout)
        modal = modal_aberto(pagina, robo.timeout) or modal
        if modal.locator("select").first.evaluate(JS_SELECIONADA) == escolhida:
            break
    else:
        fechar_modal(modal)
        raise AlunoPulado("ERRO", f"Escolhi '{escolhida}', mas o SGE desfez a escolha — nada alterado")

    try:
        extras = preencher_data_obs(robo, modal)
    except AlunoPulado:
        fechar_modal(modal)
        raise

    botao = modal.get_by_role("button", name="Alterar", exact=True)
    if botao.count() == 0:
        botao = modal.get_by_text("Alterar", exact=True)
    if botao.count() == 0:
        fechar_modal(modal)
        raise AlunoPulado("ERRO", "Não achei o botão Alterar na janela")

    if not robo.real:
        destacar(botao.first)
        pagina.wait_for_timeout(700)
        fechar_modal(modal)
        acao = "reenviaria no mesmo status" if reenvio else "escolheria"
        return "SIMULADO", f"{acao} '{escolhida}'{extras}"

    esperar_sge_parar(pagina, robo.timeout)
    botao.first.click()
    pagina.wait_for_timeout(1000)
    esperar_sge_parar(pagina, robo.timeout)

    limite = time.time() + robo.timeout / 1000
    while time.time() < limite:
        # no reenvio o status não muda: vale quando a janela fechou e o item continua no status
        if reenvio and modal_aberto(pagina, 500) is not None:
            time.sleep(0.5)
            continue
        _, itens = ler_itens(pagina)
        for atual in itens or []:
            if atual["item"] == item["item"] and atual["cod"] == item["cod"] and mesmo_status(atual["status"], escolhida):
                if reenvio:
                    robo.reenviados.add((item["cod"], item["item"]))
                    return "REENVIADO", f"{atual['status']}{extras}"
                return "ALTERADO", f"{item['status']} -> {atual['status']}{extras}"
        time.sleep(0.5)
    if reenvio:
        raise AlunoPulado("ERRO", "Cliquei em Alterar (reenvio), mas a janela não fechou — conferir manualmente")
    raise AlunoPulado("ERRO", f"Cliquei em Alterar, mas o item não mostrou '{escolhida}' — conferir manualmente")


# ==========================================================
# Processo por aluno
# ==========================================================

def erro_curto(erro):
    return f"{type(erro).__name__}: {str(erro).splitlines()[0][:200]}"


def aprender_item(robo, aluno, nomes):
    """Pergunta qual item do SGE corresponde ao item da planilha e grava em apelidos_itens."""

    unicos = list(dict.fromkeys(nomes))
    escolha = perguntar(robo, f"Não achei '{aluno['item']}' nos pedidos. Qual destes itens é ele?",
                        unicos, "Enter = nenhum")
    if escolha is None:
        return False
    apelidos = dict(robo.cfg.get("apelidos_itens", {}))
    apelidos[normalizar(aluno["item"])] = unicos[escolha]
    guardar_aprendido(robo.cfg, "apelidos_itens", apelidos)
    return True


def processar_aluno(robo, aluno, registrar, repetir=True):
    if not aluno["status"]:
        registrar("", "", "ERRO", "Linha sem o novo status na planilha")
        return

    pedidos = ir_para_pedidos(robo, aluno)
    if not pedidos:
        registrar("", "", "NAO ENCONTRADO", f"Nenhum pedido {robo.cfg['prefixo_codigo']}{aluno['sge']}.N")
        return

    pular = [normalizar(t) for t in robo.cfg.get("pular_pedidos_com_texto", [])]
    ativos = [p for p in pedidos if not any(t and t in normalizar(p["linha"]) for t in pular)]
    for p in pedidos:
        if p not in ativos:
            registrar(p["numero"], "", "SEM ITEM", "pedido ignorado (cancelado)")

    def ir_ao_pedido(numero, da_lista):
        if not da_lista:
            robo.pagina.go_back()
            esperar_carregar(robo.pagina, robo.timeout)
            if not esperar_pedidos(robo, aluno, 4000):
                ir_para_pedidos(robo, aluno)
        abrir_pedido(robo, numero)

    # 1) lê os pedidos: álbum separado tem prioridade; senão, o plano/pacote da planilha
    alvo_status = status_desejado(robo.cfg, aluno["status"])
    def concluido(i):
        if ja_adiante(i["status"], alvo_status, robo.cfg):
            return True
        if robo.reenviar:
            return (i["cod"], i["item"]) in robo.reenviados
        return mesmo_status(i["status"], alvo_status)

    escolhido = plano = aberto = None
    vistos, nomes = [], []
    for posicao, pedido in enumerate(ativos):
        numero = pedido["numero"]
        try:
            ir_ao_pedido(numero, da_lista=posicao == 0)
            aberto = numero
            _, itens = ler_itens(robo.pagina)
            if itens is None:
                raise AlunoPulado("ERRO", "Não achei a tabela de itens (coluna 'Status Item')")

            vistos.extend(f"{i['item']} ({i['status']})" for i in itens)
            livres = [i for i in itens
                      if not status_ignorado(i["status"], robo.cfg) or concluido(i)]
            nomes.extend(i["item"] for i in livres)
            albuns = [i for i in livres if eh_prioritario(i["item"], robo.cfg)]
            if albuns:
                escolhido = (numero, albuns)
                break
            planos = [i for i in livres if item_serve(i["item"], aluno, robo.cfg)]
            if planos and plano is None:
                plano = (numero, planos)
        except AlunoPulado as pulo:
            aberto = None
            registrar(numero, "", pulo.resultado, pulo.detalhe, tirar_print(robo.pagina, f"erro_{numero}"))
        except Exception as erro:
            aberto = None
            registrar(numero, "", "ERRO", erro_curto(erro), tirar_print(robo.pagina, f"erro_{numero}"))

    escolhido = escolhido or plano
    if not escolhido:
        if not ativos:
            return
        if repetir and aluno["item"] and nomes and aprender_item(robo, aluno, nomes):
            return processar_aluno(robo, aluno, registrar, repetir=False)
        procurado = f"'{aluno['item']}'" if aluno["item"] else "encadernação/álbum/book"
        registrar("", "", "SEM ITEM", f"Nenhum item {procurado}. Itens nos pedidos: " + " | ".join(vistos))
        return

    # 2) altera só os itens escolhidos, num pedido só
    numero, alvos = escolhido

    def registrar_concluido(item):
        if (item["cod"], item["item"]) in robo.reenviados:
            return
        detalhe = item["status"] if mesmo_status(item["status"], alvo_status) else f"{item['status']} (etapa adiante, ignorado)"
        registrar(numero, item["item"], "JA NO STATUS", detalhe)

    if all(concluido(i) for i in alvos):
        for item in alvos:
            registrar_concluido(item)
        return
    if aberto != numero:
        try:
            ir_ao_pedido(numero, da_lista=False)
        except Exception as erro:
            registrar(numero, "", "ERRO", erro_curto(erro), tirar_print(robo.pagina, f"erro_{numero}"))
            return

    for item in alvos:
        if concluido(item):
            registrar_concluido(item)
            continue
        try:
            frame, atuais = ler_itens(robo.pagina)
            alvo = next((a for a in atuais or [] if a["item"] == item["item"] and a["cod"] == item["cod"]), item)
            resultado, detalhe = alterar_item(robo, frame, alvo, aluno)
            registrar(numero, item["item"], resultado, detalhe)
        except AlunoPulado as pulo:
            registrar(numero, item["item"], pulo.resultado, pulo.detalhe, tirar_print(robo.pagina, f"erro_{numero}"))
        except Exception as erro:
            registrar(numero, item["item"], "ERRO", erro_curto(erro), tirar_print(robo.pagina, f"erro_{numero}"))
        robo.pagina.wait_for_timeout(500)


def carregar_progresso():
    try:
        return json.loads(PROGRESSO.read_text(encoding="utf-8"))
    except Exception:
        return {}


def guardar_progresso(aluno, proxima):
    PROGRESSO.write_text(json.dumps({
        "ultima_linha": aluno["linha"], "sge": aluno["sge"], "nome": aluno["nome"],
        "proxima_linha": proxima, "concluido": proxima is None,
        "data": f"{datetime.now():%Y-%m-%d %H:%M:%S}",
    }, indent=2, ensure_ascii=False), encoding="utf-8")


def salvar_relatorio(linhas_csv):
    with RELATORIO.open("w", newline="", encoding="utf-8-sig") as arquivo:
        escritor = csv.writer(arquivo, delimiter=";")
        escritor.writerow(["LINHA", "SGE", "NOME", "PEDIDO", "ITEM", "NOVO STATUS", "RESULTADO", "DETALHE"])
        escritor.writerows(linhas_csv)


def processar_com_contas(robo, aluno, linhas_csv):
    """
    Processa o aluno; se o SGE voltar para o login no meio (sessão derrubada), conta a queda
    (troca de conta na 2ª), entra de novo e repete o aluno. O que já foi alterado continua valendo.
    """

    robo.reenviados = set()
    valeram, rodada = [], []

    def mostrar(lista):
        resultados = []
        for pedido, item, resultado, detalhe, imagem in lista:
            CONTADORES[resultado] += 1
            resultados.append(resultado)
            texto = f"    {pedido or '-'} | {item or '-'} | {resultado}" + (f" — {detalhe}" if detalhe else "")
            print(texto + (f"\n      print: {imagem}" if imagem else ""))
            linhas_csv.append([aluno["linha"], aluno["sge"], aluno["nome"], pedido, item, aluno["status"],
                               resultado, (detalhe + (" | " + imagem if imagem else "")).strip(" |")])
        return resultados

    for tentativa in range(5):
        rodada = []

        def registrar(pedido, item, resultado, detalhe="", imagem=""):
            rodada.append((pedido, item, resultado, detalhe, imagem))

        try:
            processar_aluno(robo, aluno, registrar)
        except AlunoPulado as pulo:
            registrar("", "", pulo.resultado, pulo.detalhe, tirar_print(robo.pagina, f"erro_{aluno['sge']}"))
        except Exception as erro:
            if "has been closed" in str(erro):
                mostrar(valeram + rodada)
                raise
            registrar("", "", "ERRO", erro_curto(erro), tirar_print(robo.pagina, f"erro_{aluno['sge']}"))

        falhou = any(r[2] in ("ERRO", "NAO ENCONTRADO") for r in rodada)
        if falhou and robo.contas is not None and na_tela_de_login(robo.pagina) and tentativa < 4:
            valeram += [r for r in rodada if r[2] in ("ALTERADO", "REENVIADO")]
            print("    (o SGE voltou para o login no meio do aluno — entrando de novo e repetindo)")
            garantir_login(robo, robo.contas, PARAR)
            continue
        break

    ja_valeram = {(r[0], r[1]) for r in valeram}
    return mostrar(valeram + [r for r in rodada if not ((r[0], r[1]) in ja_valeram and r[2] == "JA NO STATUS")])


def rodar(cfg, args):
    from playwright.sync_api import sync_playwright

    alunos = ler_planilha(cfg)
    if args.status:
        for aluno in alunos:
            aluno["status"] = args.status
    data_da_rodada(cfg)

    seguinte = {a["linha"]: (alunos[i + 1]["linha"] if i + 1 < len(alunos) else None)
                for i, a in enumerate(alunos)}
    comeco = args.linha
    if args.continuar and not comeco:
        comeco = carregar_progresso().get("proxima_linha")
        if comeco:
            print(f"Continuando da linha {comeco}.")
    if comeco:
        alunos = [a for a in alunos if a["linha"] >= int(comeco)]

    if args.sge:
        base = codigo_base(args.sge, cfg)
        alunos = [a for a in alunos if a["sge"] == base]
        if not alunos:
            print(f"O código {args.sge} não está na planilha.")
            return

    feitas = carregar_feitas()
    reenviar = bool(args.mesmo_status)
    marca_reenvio = f"|REENVIO|{data_da_rodada(cfg)}|{normalizar(cfg.get('observacao'))}" if reenviar else ""
    chave = lambda a: f"{a['sge']}|{normalizar(a['status'])}|{normalizar(a['item'])}{marca_reenvio}"
    ja_feitos = [a for a in alunos if chave(a) in feitas and not args.refazer]
    alunos = [a for a in alunos if chave(a) not in feitas or args.refazer]
    if args.limite:
        alunos = alunos[: args.limite]

    real = not args.simular
    interativo = not args.sem_pausa and terminal_interativo()
    guardar = real and not args.sge
    print("=" * 70)
    print("ROBÔ SGE — ALTERAR STATUS —", "MODO REAL (vai alterar)" if real else "SIMULAÇÃO (clica em Cancelar no fim)")
    print(f"Planilha: {cfg.get('planilha_link') or cfg['planilha']}")
    print(f"Alunos para processar: {len(alunos)} | já feitos antes: {len(ja_feitos)}")
    print(f"Status: {args.status or '(o da planilha)'} | data: {data_da_rodada(cfg)} | "
          f"observação: {cfg.get('observacao') or '(nenhuma)'}"
          + (" | REENVIO NO MESMO STATUS" if reenviar else ""))
    print("Modo aprender:", "LIGADO (vou perguntar quando não souber)" if interativo else "desligado")
    print("=" * 70)
    PARAR.unlink(missing_ok=True)

    linhas_csv = []
    for aluno in ja_feitos:
        CONTADORES["JA NO STATUS"] += 1
        linhas_csv.append([aluno["linha"], aluno["sge"], aluno["nome"], "", "", aluno["status"],
                           "JA NO STATUS", "registrado em baixas_feitas.json"])

    if alunos:
        with sync_playwright() as p:
            contexto, pagina = abrir_navegador(p)
            robo = Robo(contexto, pagina, cfg, real, interativo)
            robo.reenviar = reenviar
            pagina.on("dialog", robo._dialogo)
            robo.contas = ContasSGE()
            print(f"Contas do SGE: {', '.join(c['usuario'] for c in robo.contas.contas)}")
            ir(pagina, cfg["url_clientes"], robo.timeout)
            entrar_sge(robo, robo.contas, PARAR)

            for posicao, aluno in enumerate(alunos, start=1):
                if PARAR.exists():
                    print("\nPediram para parar — parando aqui. O relatório tem o que já foi feito.")
                    break
                print(f"\n[{posicao}/{len(alunos)}] SGE {aluno['sge']} — {aluno['nome']} -> {aluno['status']}"
                      + (f" (item: {aluno['item']})" if aluno["item"] else "") + f" (linha {aluno['linha']})")
                fechado = False
                try:
                    resultados = processar_com_contas(robo, aluno, linhas_csv)
                except Exception as erro:
                    if "has been closed" not in str(erro):
                        raise
                    print("\nO navegador foi fechado — parando aqui. O relatório tem o que já foi feito.")
                    fechado = True
                    resultados = []
                if fechado:
                    break

                certos = [r for r in resultados if r in ("ALTERADO", "REENVIADO", "JA NO STATUS")]
                if real and certos and all(r in ("ALTERADO", "REENVIADO", "JA NO STATUS", "SEM ITEM") for r in resultados):
                    feitas[chave(aluno)] = {"nome": aluno["nome"], "status": aluno["status"],
                                            "data": f"{datetime.now():%Y-%m-%d %H:%M:%S}"}
                    salvar_feitas(feitas)
                if guardar:
                    guardar_progresso(aluno, seguinte.get(aluno["linha"]))
                salvar_relatorio(linhas_csv)
                time.sleep(int(cfg.get("pausa_entre_alunos_ms") or 0) / 1000)

            try:
                contexto.close()
            except Exception:
                pass

    salvar_relatorio(linhas_csv)

    print("\n" + "=" * 70)
    print(f"RELATORIO: {RELATORIO}")
    principal = (f"ALTERADOS: {CONTADORES['ALTERADO']} | REENVIADOS: {CONTADORES['REENVIADO']}" if real
                 else f"SIMULADOS: {CONTADORES['SIMULADO']}")
    print(f"RESUMO BAIXA — {principal} | JA NO STATUS: {CONTADORES['JA NO STATUS']} | "
          f"NAO ENCONTRADOS: {CONTADORES['NAO ENCONTRADO']} | SEM ITEM: {CONTADORES['SEM ITEM']} | "
          f"ERROS: {CONTADORES['ERRO']}")
    print("=" * 70)


# ==========================================================
# Modo mapear
# ==========================================================

JS_MAPA = r"""
() => {
  const vis = e => { const r = e.getBoundingClientRect(); const s = getComputedStyle(e);
                     return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none'; };
  const txt = s => (s || '').replace(/\s+/g, ' ').trim().slice(0, 120);
  const out = {botoes: [], links: [], campos: [], selecoes: [], textos: []};
  document.querySelectorAll('button, input[type=button], input[type=submit], [role=button]').forEach(e => {
    if (vis(e)) out.botoes.push(txt(e.innerText || e.value || e.getAttribute('aria-label') || e.title) || '(sem texto) id=' + e.id);
  });
  document.querySelectorAll('a[href]').forEach(e => { if (vis(e)) { const t = txt(e.innerText || e.title); if (t) out.links.push(t); } });
  document.querySelectorAll('input:not([type=hidden]):not([type=button]):not([type=submit]), textarea').forEach(e => {
    if (vis(e)) out.campos.push('placeholder="' + (e.placeholder || '') + '" name="' + (e.name || '') + '" id="' + (e.id || '') + '" tipo=' + (e.type || ''));
  });
  document.querySelectorAll('select').forEach(e => {
    out.selecoes.push('name="' + (e.name || '') + '" opcoes: ' + [...e.options].map(o => txt(o.text)).slice(0, 30).join(' | '));
  });
  document.querySelectorAll('h1, h2, h3, h4, th, legend, label, .modal-title').forEach(e => {
    if (vis(e)) { const t = txt(e.innerText); if (t) out.textos.push(t); }
  });
  for (const k in out) out[k] = [...new Set(out[k])];
  return out;
}
"""


def mapear(cfg):
    from playwright.sync_api import sync_playwright

    MAPA.write_text(f"MAPA DO SGE — {datetime.now():%d/%m/%Y %H:%M}\n", encoding="utf-8")
    with sync_playwright() as p:
        contexto, pagina = abrir_navegador(p)
        pagina.goto(cfg["url_clientes"])
        print("Navegador aberto. Faça login e vá até a tela que quer mapear.")
        tela = 0
        while True:
            if input("\nEnter = ler esta tela | q + Enter = terminar: ").strip().lower() == "q":
                break
            pagina = contexto.pages[-1]
            tela += 1
            linhas = [f"\n{'=' * 70}", f"TELA {tela} — {pagina.title()}", f"URL: {pagina.url}", "=" * 70]
            for frame in pagina.frames:
                try:
                    dados = frame.evaluate(JS_MAPA)
                except Exception:
                    continue
                for chave, lista in dados.items():
                    if lista:
                        linhas.append(f"\n{chave.upper()}:")
                        linhas.extend(f"  - {item}" for item in lista)
            imagem = PASTA / f"sge_mapa_tela{tela}.png"
            pagina.screenshot(path=str(imagem), full_page=True)
            with MAPA.open("a", encoding="utf-8") as arquivo:
                arquivo.write("\n".join(linhas) + f"\nPRINT: {imagem}\n")
            print(f"Tela {tela} salva em {MAPA.name}")
        contexto.close()


# ==========================================================
# Entrada
# ==========================================================

CABECALHOS = {
    "sge": ["SGE", "COD SGE", "CODIGO SGE", "COD CLIENTE", "CODIGO CLIENTE", "CODIGO", "COD"],
    "nome": ["CLIENTE", "NOME", "ALUNO", "FORMANDO", "NOME DO CLIENTE", "NOME DO ALUNO"],
    "status": ["STATUS", "SITUACAO", "ETAPA", "NOVO STATUS", "STATUS ITEM"],
    "item": ["ITEM", "ITENS", "ITEMS", "PRODUTO", "PACOTE", "PLANO", "ALBUM", "BOOK"],
}


def casa_cabecalho(texto, chave):
    """0 = não casa, 2 = igual a um nome conhecido, 1 = contém um nome conhecido como palavra."""

    texto = re.sub(r"\s+", " ", re.sub(r"[.:]", " ", normalizar(texto))).strip()
    if not texto:
        return 0
    if texto in CABECALHOS[chave]:
        return 2
    palavras_texto = set(re.split(r"[\s/().:_\-]+", texto))
    return 1 if any(n in palavras_texto for n in CABECALHOS[chave] if " " not in n) else 0


def detectar_colunas(cfg):
    """Procura em cada aba a linha de cabeçalho (a dos filtros) e sugere as colunas SGE, Nome, Status e Item."""

    from openpyxl.utils import get_column_letter

    livro = abrir_planilha(caminho_planilha(cfg))
    ativa = livro.active.title if livro.active is not None else None
    abas = []
    for folha in livro.worksheets:
        melhor = None
        filtro = getattr(getattr(folha, "auto_filter", None), "ref", None)
        for numero, linha in enumerate(folha.iter_rows(min_row=1, max_row=40, values_only=True), start=1):
            textos = [valor_celula(v) for v in linha]
            sugestao, pontos = {}, 0
            for chave in CABECALHOS:
                nota, indice = 0, 0
                for i, t in enumerate(textos):
                    if i not in sugestao.values() and casa_cabecalho(t, chave) > nota:
                        nota, indice = casa_cabecalho(t, chave), i
                if nota and indice not in sugestao.values():
                    sugestao[chave] = indice
                    pontos += nota * (2 if chave == "sge" else 1)
            if "sge" in sugestao and (melhor is None or pontos > melhor["pontos"]):
                cabecalhos = [[get_column_letter(i + 1), t] for i, t in enumerate(textos) if t]
                melhor = {"linha": numero, "pontos": pontos, "cabecalhos": cabecalhos,
                          "sugestao": {c: get_column_letter(i + 1) for c, i in sugestao.items()}}
        abas.append({"aba": folha.title, "filtro": filtro or "", **(melhor or {"linha": 0, "pontos": 0,
                                                                             "cabecalhos": [], "sugestao": {}})})
    livro.close()

    escolhida = max(abas, key=lambda a: (a["pontos"], a["aba"] == ativa), default=None)
    return {"abas": abas, "sugerida": escolhida["aba"] if escolhida and escolhida["pontos"] else ""}


def mostrar_deteccao(cfg):
    resultado = detectar_colunas(cfg)
    for aba in resultado["abas"]:
        if not aba["pontos"]:
            print(f"  aba '{aba['aba']}': não achei cabeçalho com SGE/código")
            continue
        colunas = ", ".join(f"{c.upper()}={l}" for c, l in aba["sugestao"].items())
        filtro = f" | filtro em {aba['filtro']}" if aba["filtro"] else ""
        print(f"  aba '{aba['aba']}': cabeçalho na linha {aba['linha']} -> {colunas}{filtro}")
    print(f"Aba sugerida: {resultado['sugerida'] or '(nenhuma)'}")
    print("DETECTAR_JSON:" + json.dumps(resultado, ensure_ascii=False))


def listar_itens(cfg):
    """Itens e status da planilha com a contagem; a última linha é o JSON lido pela UI."""

    alunos = ler_planilha(cfg)
    itens = Counter(a["item"].strip() for a in alunos if a["item"].strip())
    status = Counter(a["status"].strip() for a in alunos if a["status"].strip())
    print(f"Planilha: {len(alunos)} linhas | {len(itens)} itens diferentes | {len(status)} status diferentes")
    da_planilha = cfg.get("prioritarios_por_planilha", {}).get(id_planilha(cfg.get("planilha_link")), [])
    for nome in da_planilha:
        print(f"  prioritário desta planilha: {nome}")
    for nome, quantos in itens.most_common():
        marca = " (prioritário)" if eh_prioritario(nome, cfg) else ""
        print(f"  item: {nome} — {quantos}{marca}")
    for nome, quantos in status.most_common():
        print(f"  status: {nome} — {quantos} -> no SGE: {status_desejado(cfg, nome)}")
    print("ITENS_JSON:" + json.dumps({"itens": itens.most_common(), "status": status.most_common(),
                                      "linhas": len(alunos), "da_planilha": da_planilha}, ensure_ascii=False))


def main():
    parser = argparse.ArgumentParser(description="Robô SGE — alterar status dos itens dos pedidos")
    parser.add_argument("--mapear", action="store_true")
    parser.add_argument("--listar", action="store_true", help="só mostra o que foi lido da planilha")
    parser.add_argument("--esquecer", action="store_true", help="apaga caminho/menu/status aprendidos")
    parser.add_argument("--listar-itens", action="store_true",
                        help="mostra os itens e status da planilha com a contagem (JSON para a UI)")
    parser.add_argument("--detectar", action="store_true",
                        help="acha a linha de cabeçalho (filtros), a aba e as colunas SGE/Nome/Status/Item")
    parser.add_argument("--planilha", help="arquivo .xlsx (desliga o link)")
    parser.add_argument("--link", help="link da planilha no Google (tem prioridade sobre o arquivo)")
    parser.add_argument("--aba")
    parser.add_argument("--linha-inicial", type=int)
    parser.add_argument("--col-sge")
    parser.add_argument("--col-nome")
    parser.add_argument("--col-status", help="coluna com o NOVO status")
    parser.add_argument("--col-item", help="coluna opcional com o nome do item/pacote")
    parser.add_argument("--salvar", action="store_true", help="grava planilha/aba/colunas no config")
    modo = parser.add_mutually_exclusive_group()
    modo.add_argument("--simular", action="store_true", help="teste: abre tudo e clica em Cancelar no fim")
    modo.add_argument("--real", action="store_true", help="padrão: altera de verdade")
    parser.add_argument("--limite", type=int)
    parser.add_argument("--sge", help="processa só este código")
    parser.add_argument("--refazer", action="store_true", help="ignora baixas_feitas.json")
    parser.add_argument("--sem-pausa", action="store_true", help="não espera Enter e não pergunta nada")
    parser.add_argument("--status", help="usa este status para todos (no lugar da coluna da planilha)")
    parser.add_argument("--data", help="data da alteração dd/mm/aaaa (padrão: hoje)")
    parser.add_argument("--obs", help="observação da alteração")
    parser.add_argument("--mesmo-status", action="store_true",
                        help="reaplica o status mesmo se o item já estiver nele (reenvio com nova data/observação)")
    parser.add_argument("--continuar", action="store_true", help="continua da linha salva em baixa_progresso.json")
    parser.add_argument("--linha", type=int, help="começa nesta linha da planilha")
    args = parser.parse_args()

    cfg = carregar_config()
    if not CONFIG.exists():
        salvar_config(cfg)

    if args.esquecer:
        for chave in APRENDIDOS:
            cfg[chave] = json.loads(json.dumps(CONFIG_PADRAO[chave]))
        salvar_config(cfg)
        print("Aprendizado apagado (caminho até os pedidos, opção do menu e mapa de status).")
        return

    for chave in ("planilha", "aba", "linha_inicial", "col_sge", "col_nome", "col_status", "col_item"):
        valor = getattr(args, chave)
        if valor is not None:
            cfg[chave] = valor
    if args.planilha is not None:
        cfg["planilha_link"] = ""
    if args.link is not None:
        cfg["planilha_link"] = args.link
    if args.data is not None:
        cfg["data_status"] = args.data
    if args.obs is not None:
        cfg["observacao"] = args.obs
    if args.salvar:
        salvar_config(cfg)

    if args.mapear:
        mapear(cfg)
    elif args.detectar:
        mostrar_deteccao(cfg)
    elif args.listar_itens:
        listar_itens(cfg)
    elif args.listar:
        for aluno in ler_planilha(cfg):
            item = f" | item: {aluno['item']}" if aluno["item"] else " | itens: encadernação/álbum/book"
            print(f"linha {aluno['linha']}: SGE {aluno['sge']} — {aluno['nome']} -> {aluno['status'] or '(SEM STATUS)'}{item}")
    else:
        rodar(cfg, args)


if __name__ == "__main__":
    try:
        main()
    except (FileNotFoundError, KeyError, ValueError) as erro:
        print(f"\nERRO: {erro}")
        sys.exit(1)
