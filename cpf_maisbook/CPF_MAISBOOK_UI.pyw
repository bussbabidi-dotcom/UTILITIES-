# -*- coding: utf-8 -*-
"""
Janela do robô de CPF do MaisBook (CPF_MAISBOOK.PY).

- Troca o link da planilha do Google Drive e a aba
- Troca os logins do MaisBook e do SGE
- Inicia, continua de onde parou e para o robô
  (Parar = termina o aluno atual e para; a linha fica salva em cpf_progresso.json)
- Aba "Curso pelo SGE": roda o CURSO_SGE.PY (pesquisa o nome no SGE e cola o curso na planilha)
"""

import json
import os
import queue
import re
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, scrolledtext, ttk

PASTA_PADRAO = Path(r"C:\Users\REIS\Documents\PAPA 2\vazio\cpf_maisbook")
PASTA = Path(__file__).resolve().parent
if not (PASTA / "CPF_MAISBOOK.PY").exists():
    PASTA = PASTA_PADRAO
ROBO = PASTA / "CPF_MAISBOOK.PY"
ROBO_CURSO = PASTA / "CURSO_SGE.PY"
PROGRESSO_CURSO = PASTA / "curso_progresso.json"
PARAR_CURSO = PASTA / "curso_parar.flag"
RELATORIO_CURSO = PASTA / "curso_sge.csv"
CONFIG = PASTA / "maisbook_config.json"
CRED_MAISBOOK = PASTA / "maisbook_credenciais.json"
CRED_SGE = PASTA.parent / "baixa_sge" / "sge_credenciais.json"
PROGRESSO = PASTA / "cpf_progresso.json"
PENDENTES = PASTA / "cpf_pendentes.json"
PARAR = PASTA / "cpf_parar.flag"
RELATORIO = PASTA / "cpf_maisbook.csv"

COR_FUNDO = "#2b2b2b"
COR_CAIXA = "#353535"
COR_TEXTO = "#dddddd"
COR_VERDE = "#2e9e4f"
COR_AZUL = "#2f6db5"
COR_VERMELHO = "#8a2323"
COR_CINZA = "#6b6b6b"


def ler_json(arquivo):
    try:
        return json.loads(arquivo.read_text(encoding="utf-8"))
    except Exception:
        return {}


def gravar_json(arquivo, dados):
    arquivo.write_text(json.dumps(dados, indent=2, ensure_ascii=False), encoding="utf-8")


def python_console():
    """Com pythonw a saída se perde; o robô roda com o python.exe da mesma pasta."""

    exe = Path(sys.executable)
    candidato = exe.with_name("python.exe")
    return str(candidato if candidato.exists() else exe)


class Janela:

    def __init__(self, root):
        self.root = root
        self.processo = None
        self.fila = queue.Queue()
        self.pediu_parar = False
        self.parar_atual = PARAR

        root.title("Robô MaisBook — recadastrar CPF")
        root.configure(background=COR_FUNDO)
        root.geometry("900x860")
        root.minsize(760, 620)
        root.protocol("WM_DELETE_WINDOW", self._fechar)

        self._montar()
        self._carregar()
        self._carregar_curso()
        self._atualizar_progresso()
        self._atualizar_progresso_curso()
        root.after(150, self._ler_saida)

    # ------------------------------------------------------------------ montagem
    def _secao(self, pai, titulo):
        quadro = tk.LabelFrame(pai, text=f" {titulo} ", bg=COR_FUNDO, fg=COR_TEXTO,
                               font=("Segoe UI", 10, "bold"), padx=8, pady=6, bd=1, relief="groove")
        quadro.pack(fill="x", padx=10, pady=(8, 0))
        return quadro

    def _rotulo(self, pai, texto, linha, coluna):
        tk.Label(pai, text=texto, bg=COR_FUNDO, fg=COR_TEXTO, font=("Segoe UI", 9)).grid(
            row=linha, column=coluna, sticky="w", padx=(0, 6), pady=2)

    def _campo(self, pai, linha, coluna, largura=30, senha=False, colspan=1):
        campo = tk.Entry(pai, width=largura, bg=COR_CAIXA, fg=COR_TEXTO, insertbackground=COR_TEXTO,
                         relief="flat", font=("Segoe UI", 9), show="•" if senha else "")
        campo.grid(row=linha, column=coluna, columnspan=colspan, sticky="we", pady=2, padx=(0, 12))
        return campo

    def _botao(self, pai, texto, comando, cor=COR_AZUL):
        return tk.Button(pai, text=texto, command=comando, bg=cor, fg="white", activebackground=cor,
                         activeforeground="white", relief="flat", font=("Segoe UI", 9, "bold"),
                         padx=10, pady=4, cursor="hand2")

    def _montar(self):
        estilo = ttk.Style(self.root)
        estilo.theme_use("clam")
        estilo.configure("TNotebook", background=COR_FUNDO, borderwidth=0)
        estilo.configure("TNotebook.Tab", background=COR_CAIXA, foreground=COR_TEXTO,
                         padding=(14, 5), font=("Segoe UI", 9, "bold"))
        estilo.map("TNotebook.Tab", background=[("selected", COR_AZUL)], foreground=[("selected", "white")])
        self._montar_logins(self.root)
        abas = ttk.Notebook(self.root)
        abas.pack(fill="x", padx=6, pady=(8, 0))
        aba_cpf = tk.Frame(abas, bg=COR_FUNDO, pady=4)
        aba_curso = tk.Frame(abas, bg=COR_FUNDO, pady=4)
        abas.add(aba_cpf, text="Recadastrar CPF")
        abas.add(aba_curso, text="Curso pelo SGE")
        self._montar_cpf(aba_cpf)
        self._montar_curso(aba_curso)

        # log
        self.status = tk.Label(self.root, text="Parado", bg=COR_FUNDO, fg=COR_TEXTO,
                               font=("Segoe UI", 9, "bold"), anchor="w")
        self.status.pack(fill="x", padx=12, pady=(8, 0))
        self.log = scrolledtext.ScrolledText(self.root, bg="#1e1e1e", fg=COR_TEXTO, insertbackground=COR_TEXTO,
                                             font=("Consolas", 9), relief="flat", wrap="word")
        self.log.pack(fill="both", expand=True, padx=10, pady=(4, 10))
        self.log.tag_configure("ok", foreground="#6fd38c")
        self.log.tag_configure("erro", foreground="#ff7b7b")
        self.log.tag_configure("aviso", foreground="#f0c060")
        self.log.tag_configure("aluno", foreground="#8fb8ff")
        self._botoes_rodando(False)

    def _montar_cpf(self, aba):
        # planilha
        sec = self._secao(aba, "Planilha do Google Drive")
        sec.columnconfigure(1, weight=1)
        self._rotulo(sec, "Link:", 0, 0)
        self.link = self._campo(sec, 0, 1, largura=70, colspan=3)
        self._rotulo(sec, "Aba:", 1, 0)
        self.aba = self._campo(sec, 1, 1, largura=25)
        botoes = tk.Frame(sec, bg=COR_FUNDO)
        botoes.grid(row=1, column=2, columnspan=2, sticky="e")
        self._botao(botoes, "Salvar link", self._salvar_planilha).pack(side="left", padx=3)
        self._botao(botoes, "Testar planilha", self._testar_planilha, COR_CINZA).pack(side="left", padx=3)

        # execução
        sec = self._secao(aba, "Execução")
        self.progresso = tk.Label(sec, text="", bg=COR_FUNDO, fg="#f0c060", font=("Segoe UI", 9, "bold"),
                                  anchor="w", justify="left")
        self.progresso.grid(row=0, column=0, columnspan=8, sticky="we", pady=(0, 6))
        self._rotulo(sec, "Começar na linha:", 1, 0)
        self.linha = self._campo(sec, 1, 1, largura=8)
        self._rotulo(sec, "Limite:", 1, 2)
        self.limite = self._campo(sec, 1, 3, largura=6)
        self.simular = tk.BooleanVar(value=False)
        self.refazer = tk.BooleanVar(value=False)
        self.corrigir_curso = tk.BooleanVar(value=True)
        opcoes = tk.Frame(sec, bg=COR_FUNDO)
        opcoes.grid(row=1, column=4, columnspan=4, sticky="w")
        for texto, variavel in (("Simular (não cola o CPF)", self.simular),
                                ("Refazer os já atualizados", self.refazer),
                                ("Corrigir curso pelo SGE", self.corrigir_curso)):
            tk.Checkbutton(opcoes, text=texto, variable=variavel, bg=COR_FUNDO, fg=COR_TEXTO,
                           selectcolor=COR_CAIXA, activebackground=COR_FUNDO,
                           activeforeground=COR_TEXTO).pack(side="left", padx=(0, 10))

        botoes = tk.Frame(sec, bg=COR_FUNDO)
        botoes.grid(row=2, column=0, columnspan=8, sticky="w", pady=(6, 0))
        self.b_iniciar = self._botao(botoes, "▶  Iniciar", self._iniciar, COR_VERDE)
        self.b_continuar = self._botao(botoes, "⏩  Continuar de onde parou", self._continuar, COR_VERDE)
        self.b_conferir = self._botao(botoes, "Conferir fila do SGE", self._conferir)
        self.b_parar = self._botao(botoes, "■  Parar", self._parar, COR_VERMELHO)
        self.b_relatorio = self._botao(botoes, "Abrir relatório", self._abrir_relatorio, COR_CINZA)
        for botao in (self.b_iniciar, self.b_continuar, self.b_conferir, self.b_parar, self.b_relatorio):
            botao.pack(side="left", padx=3)

    def _montar_logins(self, pai):
        sec = self._secao(pai, "Cadastro (logins) — vale para as duas abas")
        self._rotulo(sec, "MaisBook usuário:", 0, 0)
        self.mb_usuario = self._campo(sec, 0, 1, largura=28)
        self._rotulo(sec, "senha:", 0, 2)
        self.mb_senha = self._campo(sec, 0, 3, largura=20, senha=True)
        self.mostrar = tk.BooleanVar(value=False)
        tk.Checkbutton(sec, text="mostrar senhas", variable=self.mostrar, command=self._mostrar_senhas,
                       bg=COR_FUNDO, fg=COR_TEXTO, selectcolor=COR_CAIXA, activebackground=COR_FUNDO,
                       activeforeground=COR_TEXTO).grid(row=0, column=4, padx=6)
        self._botao(sec, "Salvar logins", self._salvar_logins).grid(row=1, column=4, padx=6)

        self._rotulo(sec, "Contas do SGE\n(usa na ordem):", 1, 0)
        self.lista_sge = tk.Listbox(sec, height=4, bg=COR_CAIXA, fg=COR_TEXTO, relief="flat",
                                    selectbackground=COR_AZUL, font=("Segoe UI", 9), exportselection=False)
        self.lista_sge.grid(row=1, column=1, columnspan=3, sticky="we", pady=2, padx=(0, 12))
        self.lista_sge.bind("<<ListboxSelect>>", self._conta_selecionada)
        self._rotulo(sec, "SGE usuário:", 2, 0)
        self.sge_usuario = self._campo(sec, 2, 1, largura=28)
        self._rotulo(sec, "senha:", 2, 2)
        self.sge_senha = self._campo(sec, 2, 3, largura=20, senha=True)
        botoes = tk.Frame(sec, bg=COR_FUNDO)
        botoes.grid(row=3, column=1, columnspan=4, sticky="w", pady=(2, 0))
        self._botao(botoes, "Adicionar conta", self._adicionar_conta).pack(side="left", padx=(0, 3))
        self._botao(botoes, "Alterar selecionada", self._alterar_conta, COR_CINZA).pack(side="left", padx=3)
        self._botao(botoes, "Remover", self._remover_conta, COR_VERMELHO).pack(side="left", padx=3)
        self._botao(botoes, "▲ Subir", self._subir_conta, COR_CINZA).pack(side="left", padx=3)
        tk.Label(sec, text="Se a sessão do SGE cair 2 vezes, alguém está usando a conta e o robô troca para a próxima.",
                 bg=COR_FUNDO, fg="#aaaaaa", font=("Segoe UI", 8)).grid(row=4, column=0, columnspan=5, sticky="w")
        self.contas_sge = []

    def _montar_curso(self, aba):
        sec = self._secao(aba, "Planilha do Google (onde colar o curso)")
        sec.columnconfigure(1, weight=1)
        self._rotulo(sec, "Link:", 0, 0)
        self.c_link = self._campo(sec, 0, 1, largura=70, colspan=3)
        self._rotulo(sec, "Aba:", 1, 0)
        self.c_aba = self._campo(sec, 1, 1, largura=25)
        self._rotulo(sec, "Coluna do nome:", 1, 2)
        self.c_coluna_nome = self._campo(sec, 1, 3, largura=18)
        self._rotulo(sec, "Rótulo do nome no SGE:", 2, 0)
        self.c_rotulo_sge = self._campo(sec, 2, 1, largura=25)
        self._rotulo(sec, "Coluna de destino:", 2, 2)
        self.c_coluna_destino = self._campo(sec, 2, 3, largura=18)
        botoes = tk.Frame(sec, bg=COR_FUNDO)
        botoes.grid(row=3, column=1, columnspan=3, sticky="w", pady=(4, 0))
        self._botao(botoes, "Salvar", self._salvar_curso).pack(side="left", padx=(0, 3))
        self._botao(botoes, "Testar planilha", self._testar_curso, COR_CINZA).pack(side="left", padx=3)
        tk.Label(sec, text="Coluna: título do cabeçalho (ex: CLIENTE, CURSO) ou letra (ex: C, G). "
                           "O curso é colado por cima do que já estiver na célula.",
                 bg=COR_FUNDO, fg="#aaaaaa", font=("Segoe UI", 8)).grid(row=4, column=0, columnspan=4, sticky="w")

        sec = self._secao(aba, "Execução")
        self.c_progresso = tk.Label(sec, text="", bg=COR_FUNDO, fg="#f0c060", font=("Segoe UI", 9, "bold"),
                                    anchor="w", justify="left")
        self.c_progresso.grid(row=0, column=0, columnspan=8, sticky="we", pady=(0, 6))
        self._rotulo(sec, "Começar na linha:", 1, 0)
        self.c_linha = self._campo(sec, 1, 1, largura=8)
        self._rotulo(sec, "Limite:", 1, 2)
        self.c_limite = self._campo(sec, 1, 3, largura=6)
        self.c_simular = tk.BooleanVar(value=False)
        tk.Checkbutton(sec, text="Simular (pesquisa no SGE, não cola)", variable=self.c_simular, bg=COR_FUNDO,
                       fg=COR_TEXTO, selectcolor=COR_CAIXA, activebackground=COR_FUNDO,
                       activeforeground=COR_TEXTO).grid(row=1, column=4, columnspan=4, sticky="w")

        botoes = tk.Frame(sec, bg=COR_FUNDO)
        botoes.grid(row=2, column=0, columnspan=8, sticky="w", pady=(6, 0))
        self.c_iniciar = self._botao(botoes, "▶  Iniciar", self._iniciar_curso, COR_VERDE)
        self.c_continuar = self._botao(botoes, "⏩  Continuar de onde parou", self._continuar_curso, COR_VERDE)
        self.c_parar = self._botao(botoes, "■  Parar", self._parar, COR_VERMELHO)
        self.c_relatorio = self._botao(botoes, "Abrir relatório",
                                       lambda: self._abrir_relatorio(RELATORIO_CURSO), COR_CINZA)
        for botao in (self.c_iniciar, self.c_continuar, self.c_parar, self.c_relatorio):
            botao.pack(side="left", padx=3)
        tk.Label(sec, text="Usa as contas do SGE do Cadastro acima (mesma ordem e troca). Na 1ª vez, entre na sua conta Google "
                           "no navegador que o robô abrir.",
                 bg=COR_FUNDO, fg="#aaaaaa", font=("Segoe UI", 8)).grid(row=3, column=0, columnspan=8, sticky="w",
                                                                        pady=(4, 0))

    # ------------------------------------------------------------------ dados
    def _carregar(self):
        cfg = ler_json(CONFIG)
        self.corrigir_curso.set(cfg.get("corrigir_curso", True))
        self.link.insert(0, cfg.get("planilha_link", ""))
        self.aba.insert(0, cfg.get("aba", "GERAL ENVIOS"))
        mb, sge = ler_json(CRED_MAISBOOK), ler_json(CRED_SGE)
        self.mb_usuario.insert(0, mb.get("usuario", ""))
        self.mb_senha.insert(0, mb.get("senha", ""))
        self.contas_sge = [{"usuario": c.get("usuario", ""), "senha": c.get("senha", "")}
                           for c in sge.get("contas", []) if c.get("usuario")]
        if not self.contas_sge and sge.get("usuario"):
            self.contas_sge = [{"usuario": sge["usuario"], "senha": sge.get("senha", "")}]
        self._mostrar_contas()

    def _mostrar_senhas(self):
        simbolo = "" if self.mostrar.get() else "•"
        self.mb_senha.config(show=simbolo)
        self.sge_senha.config(show=simbolo)
        self._mostrar_contas()

    # ------------------------------------------------------------------ contas do SGE
    def _mostrar_contas(self, selecionar=None):
        self.lista_sge.delete(0, "end")
        for i, conta in enumerate(self.contas_sge, 1):
            senha = conta["senha"] if self.mostrar.get() else "•" * len(conta["senha"])
            self.lista_sge.insert("end", f"{i}. {conta['usuario']}   ({senha})")
        if selecionar is not None and 0 <= selecionar < len(self.contas_sge):
            self.lista_sge.selection_set(selecionar)

    def _indice_conta(self):
        sel = self.lista_sge.curselection()
        return sel[0] if sel else None

    def _conta_selecionada(self, _evento=None):
        i = self._indice_conta()
        if i is None:
            return
        self.sge_usuario.delete(0, "end")
        self.sge_usuario.insert(0, self.contas_sge[i]["usuario"])
        self.sge_senha.delete(0, "end")
        self.sge_senha.insert(0, self.contas_sge[i]["senha"])

    def _conta_dos_campos(self):
        usuario, senha = self.sge_usuario.get().strip(), self.sge_senha.get()
        if not usuario or not senha:
            messagebox.showerror("Conta do SGE", "Preencha usuário e senha do SGE.")
            return None
        return {"usuario": usuario, "senha": senha}

    def _adicionar_conta(self):
        conta = self._conta_dos_campos()
        if conta is None:
            return
        if any(c["usuario"].lower() == conta["usuario"].lower() for c in self.contas_sge):
            messagebox.showerror("Conta do SGE", "Essa conta já está na lista (use 'Alterar selecionada').")
            return
        self.contas_sge.append(conta)
        self._mostrar_contas(len(self.contas_sge) - 1)
        self._salvar_contas()

    def _alterar_conta(self):
        i, conta = self._indice_conta(), self._conta_dos_campos()
        if i is None or conta is None:
            return
        self.contas_sge[i] = conta
        self._mostrar_contas(i)
        self._salvar_contas()

    def _remover_conta(self):
        i = self._indice_conta()
        if i is None:
            return
        if not messagebox.askyesno("Remover conta", f"Remover a conta {self.contas_sge[i]['usuario']}?"):
            return
        del self.contas_sge[i]
        self._mostrar_contas()
        self._salvar_contas()

    def _subir_conta(self):
        i = self._indice_conta()
        if not i:
            return
        self.contas_sge[i - 1], self.contas_sge[i] = self.contas_sge[i], self.contas_sge[i - 1]
        self._mostrar_contas(i - 1)
        self._salvar_contas()

    def _salvar_contas(self):
        dados = ler_json(CRED_SGE)
        dados["contas"] = self.contas_sge
        if self.contas_sge:
            dados["usuario"], dados["senha"] = self.contas_sge[0]["usuario"], self.contas_sge[0]["senha"]
        gravar_json(CRED_SGE, dados)
        self._escrever(f"Contas do SGE salvas ({len(self.contas_sge)}).\n", "ok")

    def _salvar_planilha(self):
        link = self.link.get().strip()
        if not re.search(r"/spreadsheets/d/[\w-]+", link):
            messagebox.showerror("Link inválido", "Cole o link da planilha do Google (docs.google.com/spreadsheets/d/...).")
            return False
        cfg = ler_json(CONFIG)
        if cfg.get("planilha_link") != link:
            PROGRESSO.unlink(missing_ok=True)
        cfg["planilha_link"] = link
        cfg["aba"] = self.aba.get().strip() or "GERAL ENVIOS"
        cfg["corrigir_curso"] = self.corrigir_curso.get()
        gravar_json(CONFIG, cfg)
        self._atualizar_progresso()
        self._escrever("Link e aba salvos.\n", "ok")
        return True

    def _salvar_logins(self):
        gravar_json(CRED_MAISBOOK, {"usuario": self.mb_usuario.get().strip(), "senha": self.mb_senha.get()})
        self._escrever("Login do MaisBook salvo.\n", "ok")
        self._salvar_contas()

    def _atualizar_progresso(self):
        prog = ler_json(PROGRESSO)
        fila = len(ler_json(PENDENTES))
        if prog and not prog.get("concluido") and prog.get("proxima_linha"):
            texto = (f"Parou depois da linha {prog['ultima_linha']} (SGE {prog.get('sge')} — {prog.get('nome')}) "
                     f"em {prog.get('data')}. Próxima: linha {prog['proxima_linha']}.")
            self.b_continuar.config(text=f"⏩  Continuar da linha {prog['proxima_linha']}")
        elif prog.get("concluido"):
            texto = f"Última rodada terminou a planilha toda em {prog.get('data')}."
            self.b_continuar.config(text="⏩  Continuar de onde parou")
        else:
            texto = "Nenhuma rodada interrompida."
            self.b_continuar.config(text="⏩  Continuar de onde parou")
        self.progresso.config(text=texto + (f"   |   Esperando o SGE: {fila}" if fila else ""))
        self.b_conferir.config(text=f"Conferir fila do SGE ({fila})")

    # ------------------------------------------------------------------ ações
    def _opcoes_rodada(self):
        """Limite e Refazer valem para Iniciar e Continuar."""

        argumentos = []
        limite = self.limite.get().strip()
        if limite:
            if not limite.isdigit():
                messagebox.showerror("Limite inválido", "Digite só o número de alunos.")
                return None
            argumentos += ["--limite", limite]
        if self.refazer.get():
            argumentos.append("--refazer")
        return argumentos

    def _iniciar(self):
        argumentos = self._opcoes_rodada()
        if argumentos is None:
            return
        linha = self.linha.get().strip()
        if linha:
            if not linha.isdigit():
                messagebox.showerror("Linha inválida", "Digite só o número da linha.")
                return
            argumentos += ["--linha", linha]
        else:
            prog = ler_json(PROGRESSO)
            if prog.get("proxima_linha") and not prog.get("concluido") and not self.simular.get():
                resposta = messagebox.askyesnocancel(
                    "Rodada interrompida",
                    f"A última rodada parou antes da linha {prog['proxima_linha']}.\n\n"
                    "Sim = continuar dessa linha\nNão = começar do início")
                if resposta is None:
                    return
                if resposta:
                    argumentos.append("--continuar")
        self._rodar(argumentos)

    def _continuar(self):
        prog = ler_json(PROGRESSO)
        if not prog.get("proxima_linha") or prog.get("concluido"):
            messagebox.showinfo("Continuar", "Não há rodada interrompida — use Iniciar.")
            return
        argumentos = self._opcoes_rodada()
        if argumentos is not None:
            self._rodar(argumentos + ["--continuar"])

    def _conferir(self):
        if not ler_json(PENDENTES):
            messagebox.showinfo("Conferir", "Ninguém está esperando a resposta do SGE.")
            return
        self._rodar(["--conferir"], simular=False)

    def _rodar(self, argumentos, simular=None):
        if self.processo:
            return
        if not self._salvar_planilha():
            return
        self._salvar_logins()
        if simular is None:
            simular = self.simular.get()
        if simular:
            argumentos.append("--simular")
        self._executar(ROBO, PARAR, argumentos)

    def _executar(self, robo, parar, argumentos):
        if self.processo:
            return
        parar.unlink(missing_ok=True)
        self.parar_atual = parar
        self.pediu_parar = False
        comando = [python_console(), "-u", str(robo), *argumentos]
        self._escrever(f"\n$ {robo.name} " + " ".join(argumentos) + "\n", "aviso")
        ambiente = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
        self.processo = subprocess.Popen(
            comando, cwd=str(PASTA), stdout=subprocess.PIPE, stderr=subprocess.STDOUT, stdin=subprocess.DEVNULL,
            text=True, encoding="utf-8", errors="replace", env=ambiente,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        threading.Thread(target=self._ler_processo, args=(self.processo,), daemon=True).start()
        self.status.config(text="Rodando...")
        self._botoes_rodando(True)

    def _parar(self):
        if not self.processo:
            return
        if not self.pediu_parar:
            self.parar_atual.write_text("parar", encoding="utf-8")
            self.pediu_parar = True
            self.status.config(text="Parando: termina o aluno atual e salva a linha...")
            for botao in (self.b_parar, self.c_parar):
                botao.config(text="■  Forçar parada")
            self._escrever("\n>>> Pedido para parar: o robô termina o aluno atual e salva a linha.\n", "aviso")
            return
        if messagebox.askyesno("Forçar parada",
                               "Fechar o robô agora?\nO aluno que está no meio será refeito quando continuar."):
            self._matar()

    def _matar(self):
        if self.processo:
            subprocess.run(["taskkill", "/PID", str(self.processo.pid), "/T", "/F"],
                           capture_output=True, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))

    def _abrir_relatorio(self, arquivo=RELATORIO):
        if arquivo.exists():
            os.startfile(str(arquivo))
        else:
            messagebox.showinfo("Relatório", "Ainda não há relatório.")

    def _testar_planilha(self):
        if self._salvar_planilha():
            self._rodar(["--listar"], simular=False)

    # ------------------------------------------------------------------ aba Curso pelo SGE
    def _carregar_curso(self):
        bloco = ler_json(CONFIG).get("curso_sge") or {}
        self.c_link.insert(0, bloco.get("planilha_link", ""))
        self.c_aba.insert(0, bloco.get("aba", ""))
        self.c_coluna_nome.insert(0, bloco.get("coluna_nome", "Nome"))
        self.c_coluna_destino.insert(0, bloco.get("coluna_destino", "CURSO"))
        self.c_rotulo_sge.insert(0, bloco.get("rotulo_nome_sge", "Nome"))

    def _salvar_curso(self):
        link = self.c_link.get().strip()
        if not re.search(r"/spreadsheets/d/[\w-]+", link):
            messagebox.showerror("Link inválido", "Cole o link da planilha do Google (docs.google.com/spreadsheets/d/...).")
            return False
        coluna_nome, coluna_destino = self.c_coluna_nome.get().strip(), self.c_coluna_destino.get().strip()
        if not coluna_nome or not coluna_destino:
            messagebox.showerror("Colunas", "Preencha a coluna do nome e a coluna de destino.")
            return False
        cfg = ler_json(CONFIG)
        bloco = cfg.get("curso_sge") or {}
        if bloco.get("planilha_link") != link or bloco.get("aba") != self.c_aba.get().strip():
            PROGRESSO_CURSO.unlink(missing_ok=True)
        bloco.update({"planilha_link": link, "aba": self.c_aba.get().strip(), "coluna_nome": coluna_nome,
                      "coluna_destino": coluna_destino,
                      "rotulo_nome_sge": self.c_rotulo_sge.get().strip() or "Nome"})
        cfg["curso_sge"] = bloco
        gravar_json(CONFIG, cfg)
        self._atualizar_progresso_curso()
        self._escrever("Planilha do curso salva.\n", "ok")
        return True

    def _atualizar_progresso_curso(self):
        prog = ler_json(PROGRESSO_CURSO)
        if prog and not prog.get("concluido") and prog.get("proxima_linha"):
            texto = (f"Parou depois da linha {prog['ultima_linha']} ({prog.get('nome')}) em {prog.get('data')}. "
                     f"Próxima: linha {prog['proxima_linha']}.")
            self.c_continuar.config(text=f"⏩  Continuar da linha {prog['proxima_linha']}")
        else:
            texto = (f"Última rodada terminou a planilha toda em {prog.get('data')}." if prog.get("concluido")
                     else "Nenhuma rodada interrompida.")
            self.c_continuar.config(text="⏩  Continuar de onde parou")
        self.c_progresso.config(text=texto)

    def _opcoes_curso(self):
        argumentos = []
        limite = self.c_limite.get().strip()
        if limite:
            if not limite.isdigit():
                messagebox.showerror("Limite inválido", "Digite só o número de alunos.")
                return None
            argumentos += ["--limite", limite]
        if self.c_simular.get():
            argumentos.append("--simular")
        return argumentos

    def _iniciar_curso(self):
        argumentos = self._opcoes_curso()
        if argumentos is None:
            return
        linha = self.c_linha.get().strip()
        if linha:
            if not linha.isdigit():
                messagebox.showerror("Linha inválida", "Digite só o número da linha.")
                return
            argumentos += ["--linha", linha]
        else:
            prog = ler_json(PROGRESSO_CURSO)
            if prog.get("proxima_linha") and not prog.get("concluido") and not self.c_simular.get():
                resposta = messagebox.askyesnocancel(
                    "Rodada interrompida",
                    f"A última rodada parou antes da linha {prog['proxima_linha']}.\n\n"
                    "Sim = continuar dessa linha\nNão = começar do início")
                if resposta is None:
                    return
                if resposta:
                    argumentos.append("--continuar")
        self._rodar_curso(argumentos)

    def _continuar_curso(self):
        prog = ler_json(PROGRESSO_CURSO)
        if not prog.get("proxima_linha") or prog.get("concluido"):
            messagebox.showinfo("Continuar", "Não há rodada interrompida — use Iniciar.")
            return
        argumentos = self._opcoes_curso()
        if argumentos is not None:
            self._rodar_curso(argumentos + ["--continuar"])

    def _testar_curso(self):
        self._rodar_curso(["--listar"])

    def _rodar_curso(self, argumentos):
        if self.processo or not self._salvar_curso():
            return
        if "--listar" not in argumentos:
            if not self.contas_sge:
                messagebox.showerror("Contas do SGE", "Cadastre uma conta do SGE no Cadastro (logins).")
                return
            self._salvar_logins()
        self._executar(ROBO_CURSO, PARAR_CURSO, argumentos)

    # ------------------------------------------------------------------ saída do robô
    def _ler_processo(self, processo):
        for linha in processo.stdout:
            self.fila.put(linha)
        processo.wait()
        self.fila.put(None)

    def _ler_saida(self):
        try:
            while True:
                linha = self.fila.get_nowait()
                if linha is None:
                    self._terminou()
                    continue
                self._escrever(linha, self._tag(linha))
                achado = re.match(r"\s*\[(\d+)/(\d+)\]\s*(.*)", linha)
                if achado:
                    self.status.config(text=f"Aluno {achado.group(1)} de {achado.group(2)} — {achado.group(3).strip()}")
        except queue.Empty:
            pass
        self.root.after(150, self._ler_saida)

    def _tag(self, linha):
        texto = linha.upper()
        if re.match(r"\s*\[\d+/\d+\]", linha):
            return "aluno"
        if "ERRO" in texto or "NAO ENCONTRADO" in texto or "CPF INVALIDO" in texto or "TRACEBACK" in texto:
            return "erro"
        if ("ATUALIZADO" in texto or "COLADO" in texto) and "RESUMO" not in texto:
            return "ok"
        if ("AGUARDANDO" in texto or "PARADO" in texto or "ATENÇÃO" in texto or "RESUMO" in texto
                or "VARIOS NO SGE" in texto):
            return "aviso"
        return None

    def _terminou(self):
        codigo = self.processo.returncode if self.processo else None
        self.processo = None
        self.parar_atual.unlink(missing_ok=True)
        self.pediu_parar = False
        self._botoes_rodando(False)
        self._atualizar_progresso()
        self._atualizar_progresso_curso()
        self.status.config(text="Parado" + (f" (saiu com código {codigo})" if codigo else ""))
        self._escrever("\n--- robô parado ---\n", "aviso")

    def _escrever(self, texto, tag=None):
        self.log.insert("end", texto, tag)
        self.log.see("end")

    def _botoes_rodando(self, rodando):
        estado_livre = "disabled" if rodando else "normal"
        for botao in (self.b_iniciar, self.b_continuar, self.b_conferir, self.c_iniciar, self.c_continuar):
            botao.config(state=estado_livre)
        for botao in (self.b_parar, self.c_parar):
            botao.config(state="normal" if rodando else "disabled", text="■  Parar")

    def _fechar(self):
        if self.processo:
            if not messagebox.askyesno("Robô rodando",
                                       "O robô está rodando. Parar agora e fechar?\n"
                                       "A linha onde ele parou fica salva para continuar depois."):
                return
            self._matar()
        self.root.destroy()


def main():
    root = tk.Tk()
    Janela(root)
    root.mainloop()


if __name__ == "__main__":
    main()
