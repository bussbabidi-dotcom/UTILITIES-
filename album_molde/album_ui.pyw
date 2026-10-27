# -*- coding: utf-8 -*-
"""
Janela do álbum.

O molde das 24 páginas já está gravado. Escolha a pasta das fotos:
de cada pose entra uma, e a mesma pose não se repete na página.
A foto de identificação e as de olho fechado ficam de fora, e o álbum
segue um cenário de cada vez, com uma roupa por página. As páginas 1, 2,
12 e 24 ficam no lugar; as outras só trocam com a vizinha.
A aba Fotos deixa mover o corte; a roda do mouse muda o tamanho.
Lá também dá para trocar fotos de lugar, ou pôr uma que ficou de fora.
"""

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from PIL import Image, ImageTk

import album


COR_FUNDO = "#191c22"
COR_SUPERFICIE = "#1e2127"
COR_CAMPO = "#23272f"
COR_BORDA = "#2c313b"
COR_BOTAO = "#2a2f3a"
COR_BOTAO_HOVER = "#343a47"
COR_SELECAO = "#2f3646"
COR_TEXTO = "#d8dee9"
COR_TEXTO_FORTE = "#eceff4"
COR_SECUNDARIO = "#7b88a1"
COR_APAGADO = "#4b5163"
COR_DESTAQUE = "#88c0d0"
COR_DESTAQUE_HOVER = "#98d5e7"
COR_VERDE = "#a3be8c"
COR_AMARELO = "#ebcb8b"
COR_VERMELHO = "#bf616a"
COR_AZUL = "#81a1c1"
LADO_PREVIA = 380
MINI_TROCA = 170
CELULA_L = 190
CELULA_A = 226
ESPACO = 8


class JanelaAlbum:

    def __init__(self, root):
        self.root = root
        self.tk = tk
        self.ttk = ttk
        root.title("Álbum")
        root.minsize(980, 640)
        try:
            root.state("zoomed")
        except tk.TclError:
            root.geometry("1200x760")

        self._configurar_estilo()

        cfg = album.carregar_config()
        self.fotos = tk.StringVar(value=cfg.get("pasta_fotos") or "")
        self.saida = tk.StringVar(value=cfg.get("pasta_saida") or str(album.SAIDA_PADRAO))
        self.status = tk.StringVar(value="Escolha a pasta de fotos. O molde das 24 páginas já está gravado.")
        self.dica_corte = tk.StringVar(value="Arraste para mover. A roda do mouse muda o tamanho.")
        self.parar_evento = threading.Event()
        self.rodando = False
        self.itens = []
        self.selecionado = None
        self.imagem = None
        self.caixa = None
        self.molde_caixa = None
        self.arrastando = False
        self.deslocamento = (0, 0)
        self.escala = 1.0
        self.origem_canvas = (0, 0)
        self._foto_tk = None
        self._previa_tk = None
        self._token_foto = 0
        self._carregando = False
        self._pedido_foto = None
        self.imagem_tela = None
        self.modo_previa = tk.StringVar(value="pagina")
        self._previa_pendente = None
        self._preparando_previa = False
        self._orientacoes = {}
        self._miniaturas = {}
        self._miniaturas_tk = {}
        self._troca_tk = None
        self._token_troca = 0

        self._montar()
        root.protocol("WM_DELETE_WINDOW", self._fechar)

    def _configurar_estilo(self):
        estilo = ttk.Style(self.root)
        try:
            estilo.theme_use("clam")
        except tk.TclError:
            pass

        fonte = ("Segoe UI", 10)
        root = self.root
        root.configure(background=COR_SUPERFICIE)
        for chave_opcao, valor in (
            ("*Text.background", COR_FUNDO),
            ("*Text.foreground", COR_TEXTO),
            ("*Text.insertBackground", COR_DESTAQUE),
            ("*Text.selectBackground", COR_SELECAO),
            ("*Text.selectForeground", COR_TEXTO_FORTE),
            ("*Text.relief", "flat"),
            ("*Text.borderWidth", 0),
            ("*Text.highlightThickness", 1),
            ("*Text.highlightBackground", COR_BORDA),
            ("*Text.highlightColor", COR_APAGADO),
            ("*Text.padX", 8),
            ("*Text.padY", 6),
            ("*Text.font", ("Cascadia Mono", 9)),
            ("*Canvas.background", COR_FUNDO),
            ("*Canvas.highlightThickness", 0),
        ):
            root.option_add(chave_opcao, valor)

        plano = {"bordercolor": COR_BORDA, "lightcolor": COR_BORDA, "darkcolor": COR_BORDA}
        estilo.configure(
            ".", background=COR_SUPERFICIE, foreground=COR_TEXTO, font=fonte,
            focuscolor=COR_SUPERFICIE, troughcolor=COR_FUNDO,
            selectbackground=COR_SELECAO, selectforeground=COR_TEXTO_FORTE, **plano,
        )
        estilo.configure("TFrame", background=COR_SUPERFICIE)
        estilo.configure("TLabel", background=COR_SUPERFICIE, foreground=COR_TEXTO)
        estilo.configure("Dica.TLabel", foreground=COR_SECUNDARIO)
        estilo.configure(
            "TLabelframe", background=COR_SUPERFICIE, borderwidth=1, relief="solid", **plano,
        )
        estilo.configure(
            "TLabelframe.Label", background=COR_SUPERFICIE, foreground=COR_SECUNDARIO,
            font=("Segoe UI Semibold", 9),
        )
        botao = {"bordercolor": COR_BOTAO, "lightcolor": COR_BOTAO, "darkcolor": COR_BOTAO}
        estilo.configure(
            "TButton", background=COR_BOTAO, foreground=COR_TEXTO, padding=(12, 5),
            relief="flat", borderwidth=0, focusthickness=0, **botao,
        )
        estados_botao = [("disabled", COR_SUPERFICIE), ("pressed", COR_SELECAO), ("active", COR_BOTAO_HOVER)]
        estilo.map(
            "TButton",
            background=estados_botao, bordercolor=estados_botao,
            lightcolor=estados_botao, darkcolor=estados_botao,
            foreground=[("disabled", COR_APAGADO)],
        )
        primario = {"bordercolor": COR_DESTAQUE, "lightcolor": COR_DESTAQUE, "darkcolor": COR_DESTAQUE}
        estilo.configure(
            "Primario.TButton", background=COR_DESTAQUE, foreground=COR_FUNDO,
            font=("Segoe UI Semibold", 10), **primario,
        )
        estados_primario = [("disabled", COR_BOTAO), ("pressed", COR_AZUL), ("active", COR_DESTAQUE_HOVER)]
        estilo.map(
            "Primario.TButton",
            background=estados_primario, bordercolor=estados_primario,
            lightcolor=estados_primario, darkcolor=estados_primario,
            foreground=[("disabled", COR_APAGADO)],
        )
        campo = {
            "fieldbackground": COR_CAMPO, "foreground": COR_TEXTO, "insertcolor": COR_DESTAQUE,
            "bordercolor": COR_BORDA, "lightcolor": COR_CAMPO, "darkcolor": COR_CAMPO,
        }
        estilo.configure("TEntry", padding=(8, 5), **campo)
        estilo.map(
            "TEntry", bordercolor=[("focus", COR_APAGADO)],
            fieldbackground=[("disabled", COR_SUPERFICIE)],
            foreground=[("disabled", COR_APAGADO)],
        )
        sem_borda = {"bordercolor": COR_FUNDO, "lightcolor": COR_FUNDO, "darkcolor": COR_FUNDO}
        estilo.configure("TNotebook", background=COR_FUNDO, borderwidth=0, tabmargins=(0, 0, 0, 0), **sem_borda)
        estilo.configure(
            "TNotebook.Tab", background=COR_FUNDO, foreground=COR_SECUNDARIO,
            padding=(18, 9), borderwidth=0, focuscolor=COR_FUNDO, **sem_borda,
        )
        estilo.map(
            "TNotebook.Tab",
            background=[("selected", COR_SUPERFICIE), ("active", COR_FUNDO)],
            foreground=[("selected", COR_TEXTO_FORTE), ("active", COR_TEXTO)],
            bordercolor=[("selected", COR_SUPERFICIE)],
            lightcolor=[("selected", COR_SUPERFICIE)],
            darkcolor=[("selected", COR_SUPERFICIE)],
        )
        lista = {"background": COR_FUNDO, "fieldbackground": COR_FUNDO, "foreground": COR_TEXTO, "borderwidth": 0, **sem_borda}
        estilo.configure("Treeview", rowheight=28, **lista)
        estilo.map("Treeview", background=[("selected", COR_SELECAO)], foreground=[("selected", COR_TEXTO_FORTE)])
        estilo.configure(
            "Treeview.Heading", background=COR_FUNDO, foreground=COR_SECUNDARIO,
            font=("Segoe UI Semibold", 9), padding=(8, 6), relief="flat", borderwidth=0, **sem_borda,
        )
        estilo.map("Treeview.Heading", background=[("active", COR_SUPERFICIE)], foreground=[("active", COR_TEXTO)])
        try:
            estilo.layout("Treeview", [("Treeview.treearea", {"sticky": "nswe"})])
        except tk.TclError:
            pass
        barra = {
            "background": COR_BOTAO, "troughcolor": COR_FUNDO, "arrowsize": 10,
            "borderwidth": 0, "gripcount": 0, "relief": "flat", **sem_borda,
        }
        for orientacao, lado in (("Vertical", "ns"), ("Horizontal", "we")):
            nome = orientacao + ".TScrollbar"
            try:
                estilo.layout(nome, [(orientacao + ".Scrollbar.trough", {
                    "sticky": lado,
                    "children": [(orientacao + ".Scrollbar.thumb", {"expand": "1", "sticky": "nswe"})],
                })])
            except tk.TclError:
                pass
            estilo.configure(nome, **barra)
            estilo.map(nome, background=[("pressed", COR_APAGADO), ("active", COR_BOTAO_HOVER)])

    def _montar(self):
        abas = ttk.Notebook(self.root)
        abas.pack(fill="both", expand=True)

        principal = ttk.Frame(abas, padding=16)
        abas.add(principal, text="Álbum")
        quadro_fotos = ttk.Frame(abas, padding=12)
        abas.add(quadro_fotos, text="Fotos")
        self.abas = abas
        self.quadro_fotos = quadro_fotos
        abas.bind("<<NotebookTabChanged>>", self._aba_mudou)

        principal.columnconfigure(0, weight=1)
        principal.rowconfigure(1, weight=1)
        principal.rowconfigure(3, weight=1)

        origem = ttk.Frame(principal)
        origem.grid(row=0, column=0, sticky="we")
        origem.columnconfigure(1, weight=1)
        ttk.Label(origem, text="Molde").grid(row=0, column=0, sticky="w")
        ttk.Label(origem, text="24 páginas gravadas, buracos já marcados").grid(row=0, column=1, sticky="w", padx=6)
        ttk.Label(origem, text="Fotos").grid(row=1, column=0, sticky="w")
        ttk.Entry(origem, textvariable=self.fotos).grid(row=1, column=1, sticky="we", padx=6, pady=2)
        ttk.Button(origem, text="Escolher...", command=lambda: self._escolher("fotos")).grid(row=1, column=2)
        ttk.Label(origem, text="Saída").grid(row=2, column=0, sticky="w")
        ttk.Entry(origem, textvariable=self.saida).grid(row=2, column=1, sticky="we", padx=6, pady=2)
        ttk.Button(origem, text="Escolher...", command=lambda: self._escolher("saida")).grid(row=2, column=2)
        ttk.Label(
            origem,
            text="Fotos iguais em seguida contam como a mesma pose. Entra a mais nítida olhando para a câmera, "
                 "um cenário de cada vez e uma roupa por página (páginas 1, 2, 12 e 24 fixas; as outras só trocam com a vizinha), "
                 "sem a foto de identificação e sem foto de olho fechado.",
            style="Dica.TLabel",
        ).grid(row=3, column=0, columnspan=3, sticky="w", pady=(6, 0))

        quadro = ttk.LabelFrame(principal, text="Páginas e fotos", padding=6)
        quadro.grid(row=1, column=0, sticky="nsew", pady=(8, 0))
        quadro.columnconfigure(0, weight=1)
        quadro.rowconfigure(0, weight=1)
        colunas = ("pagina", "buraco", "foto", "situacao")
        self.tabela = ttk.Treeview(quadro, columns=colunas, show="headings")
        for col, titulo, largura in (
            ("pagina", "Página", 220),
            ("buraco", "Buraco", 80),
            ("foto", "Foto", 280),
            ("situacao", "Situação", 320),
        ):
            self.tabela.heading(col, text=titulo)
            self.tabela.column(col, width=largura, anchor="w", stretch=(col != "buraco"))
        self.tabela.tag_configure("aviso", foreground=COR_VERMELHO)
        self.tabela.tag_configure("ok", foreground=COR_VERDE)
        self.tabela.tag_configure("nota", foreground=COR_AMARELO)
        self.tabela.grid(row=0, column=0, sticky="nsew")
        barra = ttk.Scrollbar(quadro, orient="vertical", command=self.tabela.yview)
        barra.grid(row=0, column=1, sticky="ns")
        self.tabela.configure(yscrollcommand=barra.set)
        self.tabela.bind("<<TreeviewSelect>>", self._selecionou_tabela)

        acoes = ttk.Frame(principal)
        acoes.grid(row=2, column=0, sticky="we", pady=(8, 0))
        self.botao_atualizar = ttk.Button(acoes, text="Atualizar lista", command=self.atualizar)
        self.botao_atualizar.pack(side="left")
        self.botao_gerar = ttk.Button(
            acoes, text="Gerar páginas", style="Primario.TButton", command=self.gerar,
        )
        self.botao_gerar.pack(side="left", padx=6)
        self.botao_parar = ttk.Button(acoes, text="Parar", command=self.parar, state="disabled")
        self.botao_parar.pack(side="left")
        ttk.Label(acoes, textvariable=self.status).pack(side="left", padx=12)

        saida = ttk.LabelFrame(principal, text="Saída", padding=6)
        saida.grid(row=3, column=0, sticky="nsew", pady=(8, 0))
        saida.columnconfigure(0, weight=1)
        saida.rowconfigure(0, weight=1)
        self.texto = tk.Text(saida, height=10, wrap="word", state="disabled")
        self.texto.grid(row=0, column=0, sticky="nsew")
        barra_saida = ttk.Scrollbar(saida, orient="vertical", command=self.texto.yview)
        barra_saida.grid(row=0, column=1, sticky="ns")
        self.texto.configure(yscrollcommand=barra_saida.set)

        self._montar_fotos(quadro_fotos)

    def _montar_fotos(self, pagina):
        pagina.columnconfigure(1, weight=1)
        pagina.rowconfigure(0, weight=1)

        esquerda = ttk.LabelFrame(pagina, text="Buracos", padding=6)
        esquerda.grid(row=0, column=0, sticky="ns")
        self.lista_fotos = ttk.Treeview(
            esquerda, columns=("rotulo",), show="headings", height=24, selectmode="browse",
        )
        self.lista_fotos.heading("rotulo", text="Página / foto")
        self.lista_fotos.column("rotulo", width=260, anchor="w")
        self.lista_fotos.tag_configure("aviso", foreground=COR_VERMELHO)
        self.lista_fotos.pack(fill="y", expand=True)
        self.lista_fotos.bind("<<TreeviewSelect>>", self._selecionou_lista)

        centro = ttk.LabelFrame(pagina, text="Corte na foto", padding=6)
        centro.grid(row=0, column=1, sticky="nsew", padx=8)
        centro.rowconfigure(0, weight=1)
        centro.columnconfigure(0, weight=1)
        self.canvas = tk.Canvas(centro, background=COR_FUNDO, cursor="fleur")
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.canvas.bind("<Configure>", lambda e: self._desenhar())
        self.canvas.bind("<ButtonPress-1>", self._pressionou)
        self.canvas.bind("<B1-Motion>", self._moveu)
        self.canvas.bind("<ButtonRelease-1>", self._soltou)
        self.canvas.bind("<MouseWheel>", self._roda)
        ttk.Label(centro, textvariable=self.dica_corte, style="Dica.TLabel").grid(row=1, column=0, sticky="w", pady=(6, 0))

        direita = ttk.Frame(pagina)
        direita.grid(row=0, column=2, sticky="ns")
        previa = ttk.LabelFrame(direita, text="Prévia do resultado", padding=6)
        previa.pack(fill="x")
        modos = ttk.Frame(previa)
        modos.pack(fill="x")
        ttk.Radiobutton(
            modos, text="Na página", value="pagina", variable=self.modo_previa,
            command=self._agendar_previa,
        ).pack(side="left")
        ttk.Radiobutton(
            modos, text="Foto final", value="foto", variable=self.modo_previa,
            command=self._agendar_previa,
        ).pack(side="left", padx=8)
        self._previa_tk = ImageTk.PhotoImage(Image.new("RGB", (LADO_PREVIA, LADO_PREVIA), COR_FUNDO))
        self.previa = tk.Label(previa, background=COR_FUNDO, image=self._previa_tk)
        self.previa.pack(pady=(6, 0))
        acoes = ttk.LabelFrame(direita, text="Corte", padding=6)
        acoes.pack(fill="x", pady=(8, 0))
        self.botao_rosto = ttk.Button(acoes, text="Usar rosto", command=self.usar_rosto)
        self.botao_rosto.pack(fill="x")
        ttk.Label(
            acoes,
            text="O ajuste manual fica gravado e entra na hora de gerar.",
            style="Dica.TLabel", wraplength=220,
        ).pack(anchor="w", pady=(8, 0))
        lugar = ttk.LabelFrame(direita, text="Lugar", padding=6)
        lugar.pack(fill="x", pady=(8, 0))
        self.botao_trocar = ttk.Button(lugar, text="Trocar de lugar...", command=self.trocar_lugar)
        self.botao_trocar.pack(fill="x")
        ttk.Label(
            lugar,
            text="Troca esta foto com a de outro buraco ou com uma que ficou de fora.",
            style="Dica.TLabel", wraplength=220,
        ).pack(anchor="w", pady=(8, 0))

    def _escolher(self, qual):
        atual = {"fotos": self.fotos, "saida": self.saida}[qual].get()
        pasta = filedialog.askdirectory(initialdir=atual or str(album.PASTA))
        if not pasta:
            return
        {"fotos": self.fotos, "saida": self.saida}[qual].set(pasta)
        self._gravar_caminhos()
        if qual != "saida":
            self.atualizar()

    def _gravar_caminhos(self):
        album.gravar_config({
            "pasta_fotos": self.fotos.get(),
            "pasta_saida": self.saida.get(),
        })

    def _log(self, texto):
        def ui():
            self.texto.configure(state="normal")
            self.texto.insert("end", texto + "\n")
            self.texto.see("end")
            self.texto.configure(state="disabled")
        self.root.after(0, ui)

    def _limpar_log(self):
        self.texto.configure(state="normal")
        self.texto.delete("1.0", "end")
        self.texto.configure(state="disabled")

    def atualizar(self):
        if self.rodando:
            return
        fotos = self.fotos.get().strip()
        if not fotos or not Path(fotos).is_dir():
            self.status.set("Escolha a pasta das fotos.")
            return
        self._token_foto += 1
        self._travar(True, parar=False)
        self.status.set("Separando as poses...")
        self._gravar_caminhos()

        def trabalho():
            def ao_foto(indice, total, nome):
                self.root.after(0, lambda indice=indice, total=total, nome=nome: self.status.set(
                    "Lendo %s de %s: %s" % (indice, total, nome)
                ))

            try:
                itens = album.montar_lista(fotos, ao_foto=ao_foto)
                erro = None
            except OSError as e:
                itens = []
                erro = str(e)
            self.root.after(0, lambda: self._aplicar_lista(itens, erro))

        threading.Thread(target=trabalho, daemon=True).start()

    def _aplicar_lista(self, itens, erro):
        self._travar(False)
        self.tabela.delete(*self.tabela.get_children())
        self.lista_fotos.delete(*self.lista_fotos.get_children())
        self.itens = itens
        if erro:
            self.status.set(erro)
            messagebox.showerror("Álbum", erro)
            return
        self._preencher_listas()

    def _preencher_listas(self):
        self.tabela.delete(*self.tabela.get_children())
        self.lista_fotos.delete(*self.lista_fotos.get_children())
        buracos = 0
        for n, item in enumerate(self.itens):
            pagina = self._nome_pagina(item)
            buraco = str(item["indice"]) if item["indice"] else "—"
            foto = item["foto"].name if item["foto"] else "—"
            situacao = item["situacao"]
            aviso = (
                situacao in ("falta foto", "erro ao abrir")
                or situacao.startswith("mesma pose")
                or item["tipo"] in ("identificacao", "olho_fechado")
                or any(t in situacao for t in (
                    "cenário misturado", "roupa misturada", "orientação diferente", "orientação trocada",
                    "era identificação",
                    "olho fechado",
                ))
            )
            nota = any(t in situacao for t in (
                "de novo", "repetida", "não olha", "olhos não visíveis", "trocada",
            ))
            tag = "aviso" if aviso else ("nota" if nota else "")
            self.tabela.insert("", "end", iid=str(n), values=(pagina, buraco, foto, item["situacao"]), tags=(tag,))
            if item["tipo"] == "buraco" and item["foto"]:
                buracos += 1
                self.lista_fotos.insert(
                    "", "end", iid=str(n),
                    values=("%s  ·  %s  ·  %s" % (pagina, buraco, foto),),
                )
        sobras = sum(1 for i in self.itens if i["tipo"] == "sobra")
        faltas = sum(1 for i in self.itens if i["situacao"] == "falta foto")
        poses = len({i["pose"] for i in self.itens if i.get("pose") is not None})
        cenarios = len({i["cenario"] for i in self.itens if i.get("cenario") is not None})
        roupas = len({i["roupa"] for i in self.itens if i.get("roupa") is not None})
        identificacoes = sum(1 for i in self.itens if i["tipo"] == "identificacao")
        fechados = sum(1 for i in self.itens if i["tipo"] == "olho_fechado")
        self.status.set(
            "%s buraco(s), %s pose(s), %s cenário(s), %s roupa(s). %s foto(s) ficaram de fora, "
            "%s de identificação, %s de olho fechado. %s falta(m)." % (
                buracos, poses, cenarios, roupas, sobras, identificacoes, fechados, faltas,
            )
        )

    @staticmethod
    def _nome_pagina(item):
        """Posição no álbum e molde, ex. '05 · DGL 2026_12 copiar'."""
        if not item["pagina"]:
            return "—"
        if item.get("ordem"):
            return "%02d · %s" % (item["ordem"], item["pagina"].stem)
        return item["pagina"].name

    def _linha(self, iid):
        try:
            return self.itens[int(iid)]
        except (TypeError, ValueError, IndexError):
            return None

    def _selecionou_tabela(self, _evento=None):
        sel = self.tabela.selection()
        if not sel:
            return
        if self.lista_fotos.exists(sel[0]) and self.lista_fotos.selection() != sel:
            self.lista_fotos.selection_set(sel[0])
            self.lista_fotos.see(sel[0])

    def _aba_mudou(self, _evento=None):
        if self.abas.select() != str(self.quadro_fotos):
            return
        if not self.itens:
            self.atualizar()
        sel = self.lista_fotos.selection() or self.tabela.selection()
        if sel and self.lista_fotos.exists(sel[0]):
            self.lista_fotos.selection_set(sel[0])
            self._abrir_item(self._linha(sel[0]))

    def _selecionou_lista(self, _evento=None):
        sel = self.lista_fotos.selection()
        if not sel:
            return
        if self.tabela.exists(sel[0]) and self.tabela.selection() != sel:
            self.tabela.selection_set(sel[0])
            self.tabela.see(sel[0])
        self._abrir_item(self._linha(sel[0]))

    def _abrir_item(self, item):
        if not item or item["tipo"] != "buraco" or not item["foto"]:
            return
        if item is self.selecionado and (self.imagem is not None or self._carregando):
            return
        self.selecionado = item
        self._token_foto += 1
        self.dica_corte.set("Procurando o rosto...")
        self.canvas.delete("all")
        self._pedido_foto = (self._token_foto, item)
        if not self._carregando:
            self._carregar_proxima()

    def _carregar_proxima(self):
        """Uma foto por vez: só a última escolhida é carregada."""
        pedido, self._pedido_foto = self._pedido_foto, None
        if not pedido:
            self._carregando = False
            return
        self._carregando = True
        token, item = pedido
        caminho = item["foto"]
        pagina = item["pagina"]
        indice = item["indice"]
        buraco = item["buraco"]

        def trabalho():
            try:
                img, _icc, _dpi = album.abrir_imagem(caminho)
                caixa, origem = album.caixa_da_foto(caminho, pagina, indice, (buraco[2], buraco[3]), img)
                tela = img.copy()
                tela.thumbnail((1600, 1600), Image.BILINEAR)
            except OSError as e:
                def falhou(erro=str(e)):
                    self.dica_corte.set("Erro ao abrir a foto: " + erro)
                    self._carregar_proxima()
                self.root.after(0, falhou)
                return
            self.root.after(0, lambda: self._aplicar_foto(token, item, img, tela, caixa, origem))

        threading.Thread(target=trabalho, daemon=True).start()

    def _aplicar_foto(self, token, item, img, tela, caixa, origem):
        self._carregar_proxima()
        if token != self._token_foto:
            img.close()
            tela.close()
            return
        if self.imagem is not None:
            self.imagem.close()
        if self.imagem_tela is not None:
            self.imagem_tela.close()
        self.imagem = img
        self.imagem_tela = tela
        self.caixa = caixa
        self.molde_caixa = (item["buraco"][2], item["buraco"][3])
        self.dica_corte.set(
            "%s  ·  %s. Arraste para mover. A roda do mouse muda o tamanho."
            % (item["foto"].name, origem)
        )
        self._desenhar()
        self._atualizar_previa()

    def _desenhar(self):
        if self.imagem is None or self.caixa is None:
            return
        cw = max(self.canvas.winfo_width(), 20)
        ch = max(self.canvas.winfo_height(), 20)
        escala = min(cw / self.imagem.width, ch / self.imagem.height)
        tw, th = max(1, round(self.imagem.width * escala)), max(1, round(self.imagem.height * escala))
        ox = (cw - tw) // 2
        oy = (ch - th) // 2
        self.escala = escala
        self.origem_canvas = (ox, oy)
        mostra = (self.imagem_tela or self.imagem).resize((tw, th), Image.BILINEAR)
        self._foto_tk = ImageTk.PhotoImage(mostra)
        self.canvas.delete("all")
        self.canvas.create_image(ox, oy, image=self._foto_tk, anchor="nw")
        l, t, r, b = self.caixa
        self.canvas.create_rectangle(
            ox + l * escala, oy + t * escala, ox + r * escala, oy + b * escala,
            outline=COR_DESTAQUE, width=2,
        )

    def _ponto_imagem(self, x, y):
        ox, oy = self.origem_canvas
        return (x - ox) / self.escala, (y - oy) / self.escala

    def _pressionou(self, evento):
        if not self.caixa:
            return
        ix, iy = self._ponto_imagem(evento.x, evento.y)
        l, t, r, b = self.caixa
        if l <= ix <= r and t <= iy <= b:
            self.arrastando = True
            self.deslocamento = (ix - (l + r) / 2, iy - (t + b) / 2)

    def _mover_caixa(self, cx, cy, altura=None):
        if not self.caixa or self.imagem is None:
            return
        l, t, r, b = self.caixa
        h = altura if altura is not None else (b - t)
        w = h * self.molde_caixa[0] / float(self.molde_caixa[1])
        self.caixa = album.ajustar_caixa(
            (cx - w / 2, cy - h / 2, cx + w / 2, cy + h / 2),
            self.imagem.size, self.molde_caixa,
        )
        self._desenhar()

    def _moveu(self, evento):
        if not self.arrastando:
            return
        ix, iy = self._ponto_imagem(evento.x, evento.y)
        dx, dy = self.deslocamento
        self._mover_caixa(ix - dx, iy - dy)
        self._agendar_previa()

    def _soltou(self, _evento):
        if not self.arrastando:
            return
        self.arrastando = False
        self._gravar_caixa("manual")

    def _roda(self, evento):
        if not self.caixa or self.imagem is None:
            return
        l, t, r, b = self.caixa
        fator = 1.08 if evento.delta > 0 else 1 / 1.08
        self._mover_caixa((l + r) / 2, (t + b) / 2, altura=(b - t) * fator)
        self._gravar_caixa("manual")

    def _gravar_caixa(self, origem):
        item = self.selecionado
        if not item or not self.caixa:
            return
        album.gravar_manual(
            item["foto"], item["pagina"], item["indice"], self.molde_caixa,
            self.caixa, self.imagem.size,
        )
        album.esquecer_recorte_rapido(item["foto"], item["pagina"], item["indice"], self.molde_caixa)
        self.dica_corte.set(
            "%s  ·  %s. Arraste para mover. A roda do mouse muda o tamanho."
            % (item["foto"].name, origem)
        )
        self._atualizar_previa()

    def _agendar_previa(self):
        if self._previa_pendente:
            self.root.after_cancel(self._previa_pendente)
        self._previa_pendente = self.root.after(80, self._atualizar_previa)

    def _recorte_atual(self, tamanho):
        """Corte desta foto tirado da cópia de tela, sem tocar na foto grande."""
        fator = self.imagem_tela.width / float(self.imagem.width)
        caixa = tuple(round(c * fator) for c in self.caixa)
        return self.imagem_tela.crop(caixa).resize(tamanho, Image.BILINEAR)

    def _outras_da_pagina(self, item):
        return [
            outro for outro in self.itens
            if outro is not item and outro["tipo"] == "buraco" and outro["foto"]
            and outro["pagina"] == item["pagina"]
        ]

    def _preparar_outras(self, item, escala):
        """Corta as outras fotos da página em segundo plano, uma por vez."""
        faltando = []
        for outro in self._outras_da_pagina(item):
            _x, _y, w, h = album.buraco_reduzido(outro["buraco"], escala)
            molde = (outro["buraco"][2], outro["buraco"][3])
            if not album.tem_recorte_rapido(outro["foto"], outro["pagina"], outro["indice"], molde, (w, h)):
                faltando.append((outro, molde, (w, h)))
        if not faltando or self._preparando_previa:
            return
        self._preparando_previa = True

        def trabalho():
            for outro, molde, tamanho in faltando:
                try:
                    album.recorte_rapido(outro["foto"], outro["pagina"], outro["indice"], molde, tamanho)
                except OSError:
                    pass

            def pronto():
                self._preparando_previa = False
                self._agendar_previa()
            self.root.after(0, pronto)

        threading.Thread(target=trabalho, daemon=True).start()

    def _atualizar_previa(self):
        self._previa_pendente = None
        item = self.selecionado
        if self.imagem is None or self.imagem_tela is None or self.caixa is None or not item:
            return
        if self.modo_previa.get() == "pagina":
            _base, escala = album.pagina_reduzida(item["pagina"])
            colocadas = []
            for outro in self._outras_da_pagina(item):
                _x, _y, w, h = album.buraco_reduzido(outro["buraco"], escala)
                molde = (outro["buraco"][2], outro["buraco"][3])
                foto = None
                if album.tem_recorte_rapido(outro["foto"], outro["pagina"], outro["indice"], molde, (w, h)):
                    foto = album.recorte_rapido(outro["foto"], outro["pagina"], outro["indice"], molde, (w, h))
                colocadas.append((outro["buraco"], foto))
            _x, _y, w, h = album.buraco_reduzido(item["buraco"], escala)
            colocadas.append((item["buraco"], self._recorte_atual((w, h))))
            imagem = album.montar_previa_pagina(item["pagina"], colocadas)
            self._preparar_outras(item, escala)
        else:
            mw, mh = self.molde_caixa
            fator = min(1.0, LADO_PREVIA * 2 / float(max(mw, mh)))
            imagem = self._recorte_atual((max(1, round(mw * fator)), max(1, round(mh * fator))))
        imagem.thumbnail((LADO_PREVIA, LADO_PREVIA), Image.LANCZOS)
        fundo = Image.new("RGB", (LADO_PREVIA, LADO_PREVIA), COR_FUNDO)
        fundo.paste(imagem, ((LADO_PREVIA - imagem.width) // 2, (LADO_PREVIA - imagem.height) // 2))
        imagem.close()
        self._previa_tk = ImageTk.PhotoImage(fundo)
        self.previa.configure(image=self._previa_tk)

    def usar_rosto(self):
        item = self.selecionado
        if not item or self.imagem is None:
            return
        album.apagar_manual(item["foto"], item["pagina"], item["indice"], self.molde_caixa)
        album.esquecer_recorte_rapido(item["foto"], item["pagina"], item["indice"], self.molde_caixa)
        self.caixa, origem = album.caixa_da_foto(
            item["foto"], item["pagina"], item["indice"], self.molde_caixa, self.imagem,
        )
        self.dica_corte.set(
            "%s  ·  %s. Arraste para mover. A roda do mouse muda o tamanho."
            % (item["foto"].name, origem)
        )
        self._desenhar()
        self._atualizar_previa()

    def _miniatura(self, foto):
        chave = (foto, MINI_TROCA)
        if chave not in self._miniaturas:
            img = album.miniatura_leve(foto, MINI_TROCA)
            self._miniaturas[chave] = img
            self._orientacoes[foto] = album.orientacao(*img.size)
        return self._miniaturas[chave]

    def _orientacao_foto(self, foto):
        if foto not in self._orientacoes:
            self._miniatura(foto)
        return self._orientacoes[foto]

    def _numeros_pagina(self):
        numeros = {}
        for atual in self.itens:
            if atual["pagina"] and atual["pagina"] not in numeros:
                numeros[atual["pagina"]] = atual.get("ordem") or len(numeros) + 1
        return numeros

    def _rotulo_curto(self, item, numeros):
        if item["tipo"] == "buraco":
            texto = "pág %02d · buraco %s" % (numeros.get(item["pagina"], 0), item["indice"])
            return texto if item["foto"] else texto + " (vazio)"
        return {
            "sobra": "fora (sobra)",
            "identificacao": "identificação",
            "olho_fechado": "olho fechado",
        }.get(item["tipo"], "—")

    def _lugar_texto(self, item):
        if item["tipo"] == "buraco":
            return "%s  ·  buraco %s" % (self._nome_pagina(item), item["indice"])
        return {
            "sobra": "fora (sobra)",
            "identificacao": "fora (identificação)",
            "olho_fechado": "fora (olho fechado)",
        }.get(item["tipo"], "—")

    def trocar_lugar(self):
        if self.rodando:
            return
        sel = self.lista_fotos.selection()
        item = self._linha(sel[0]) if sel else self.selecionado
        if not item or item["tipo"] != "buraco" or not item["foto"]:
            messagebox.showinfo("Álbum", "Escolha uma foto na lista de buracos.")
            return
        self._dialogo_troca(item)

    def _dialogo_troca(self, item):
        janela = tk.Toplevel(self.root)
        janela.title("Trocar de lugar")
        janela.configure(background=COR_SUPERFICIE)
        janela.transient(self.root)
        janela.geometry("1180x740")
        janela.minsize(820, 520)
        quadro = ttk.Frame(janela, padding=12)
        quadro.pack(fill="both", expand=True)
        quadro.columnconfigure(0, weight=1)
        quadro.rowconfigure(2, weight=1)

        numeros = self._numeros_pagina()
        o_origem = album.orientacao(item["buraco"][2], item["buraco"][3])
        ttk.Label(
            quadro,
            text="Trocar %s (%s, buraco %s) com:" % (
                item["foto"].name, self._rotulo_curto(item, numeros), o_origem,
            ),
        ).grid(row=0, column=0, columnspan=2, sticky="w")

        filtro = tk.StringVar(value="todos")
        so_compativel = tk.BooleanVar(value=False)
        filtros = ttk.Frame(quadro)
        filtros.grid(row=1, column=0, columnspan=2, sticky="we", pady=(6, 6))

        caixa_grade = ttk.Frame(quadro)
        caixa_grade.grid(row=2, column=0, sticky="nsew")
        caixa_grade.columnconfigure(0, weight=1)
        caixa_grade.rowconfigure(0, weight=1)
        tela = tk.Canvas(caixa_grade, background=COR_FUNDO, yscrollincrement=24, cursor="hand2")
        tela.grid(row=0, column=0, sticky="nsew")
        barra = ttk.Scrollbar(caixa_grade, orient="vertical", command=tela.yview)
        barra.grid(row=0, column=1, sticky="ns")

        lado = ttk.Frame(quadro, width=320)
        lado.grid(row=2, column=1, sticky="n", padx=(12, 0))
        vazia = ImageTk.PhotoImage(Image.new("RGB", (300, 300), COR_FUNDO))
        detalhe = tk.Label(lado, background=COR_FUNDO, image=vazia)
        detalhe.image = vazia
        detalhe.pack()
        info = tk.StringVar(value="Clique numa foto para escolher. Duplo clique troca.")
        ttk.Label(lado, textvariable=info, style="Dica.TLabel", wraplength=300, justify="left").pack(
            anchor="w", pady=(6, 0),
        )

        parar = threading.Event()
        trava = threading.Lock()
        estado = {
            "celulas": [], "colunas": 0, "sel": None, "visiveis": [],
            "faltam": {}, "por_foto": {}, "refazer": None, "total": 0,
        }
        progresso = tk.StringVar()

        def cor_tipo(outro):
            if outro["tipo"] == "buraco":
                return COR_BORDA if outro["foto"] else COR_APAGADO
            return COR_AMARELO if outro["tipo"] == "sobra" else COR_VERMELHO

        def incompativel(outro):
            o_foto = self._orientacoes.get(outro["foto"]) if outro["foto"] else None
            if o_foto and not album.aceita(o_origem, o_foto):
                return True
            o_minha = self._orientacoes.get(item["foto"])
            if outro["tipo"] == "buraco" and o_minha:
                o_buraco = album.orientacao(outro["buraco"][2], outro["buraco"][3])
                if not album.aceita(o_buraco, o_minha):
                    return True
            return False

        def candidatos():
            modo = filtro.get()
            saida = []
            for n, outro in enumerate(self.itens):
                if outro is item:
                    continue
                if outro["tipo"] == "buraco":
                    if modo == "fora":
                        continue
                elif outro["tipo"] in album.FORA_DO_ALBUM and outro["foto"]:
                    if modo == "album":
                        continue
                else:
                    continue
                if so_compativel.get() and incompativel(outro):
                    continue
                saida.append((n, outro))
            return saida

        def marcar(i):
            celula = estado["celulas"][i]
            outro = celula["item"]
            foto = outro["foto"]
            o_foto = self._orientacoes.get(foto) if foto else None
            ruim = incompativel(outro)
            if o_foto:
                tela.itemconfigure("ot%d" % i, text={"vertical": "V", "horizontal": "H"}.get(o_foto, "Q"))
                tela.itemconfigure("or%d" % i, fill=COR_VERMELHO if ruim else COR_BOTAO)
                tela.itemconfigure("o%d" % i, state="normal")
            tela.itemconfigure("a%d" % i, state="normal" if ruim else "hidden")
            selecionada = estado["sel"] == i
            tela.itemconfigure(
                "r%d" % i,
                outline=COR_DESTAQUE if selecionada else (COR_VERMELHO if ruim else cor_tipo(outro)),
                width=3 if selecionada else (2 if ruim else 1),
            )
            foto_tk = self._miniaturas_tk.get(foto) if foto else None
            if foto_tk is not None:
                tela.itemconfigure("i%d" % i, image=foto_tk)
                tela.itemconfigure("p%d" % i, state="hidden")

        def desenhar():
            largura = max(tela.winfo_width(), CELULA_L + 2 * ESPACO)
            colunas = max(1, (largura - ESPACO) // (CELULA_L + ESPACO))
            estado["colunas"] = colunas
            tela.delete("all")
            lista = candidatos()
            sel_n = estado["celulas"][estado["sel"]]["n"] if estado["sel"] is not None else None
            estado["celulas"] = []
            estado["por_foto"] = {}
            estado["sel"] = None
            margem = (largura - colunas * (CELULA_L + ESPACO) + ESPACO) // 2
            for i, (n, outro) in enumerate(lista):
                linha, coluna = divmod(i, colunas)
                x0 = margem + coluna * (CELULA_L + ESPACO)
                y0 = ESPACO + linha * (CELULA_A + ESPACO)
                cor = cor_tipo(outro)
                tela.create_rectangle(
                    x0, y0, x0 + CELULA_L, y0 + CELULA_A,
                    fill=COR_SUPERFICIE, outline=cor, tags=("r%d" % i,),
                )
                cx, cy = x0 + CELULA_L // 2, y0 + 8 + MINI_TROCA // 2
                tela.create_rectangle(
                    cx - MINI_TROCA // 2, cy - MINI_TROCA // 2, cx + MINI_TROCA // 2, cy + MINI_TROCA // 2,
                    fill=COR_CAMPO, outline="",
                )
                tela.create_text(
                    cx, cy, text="carregando..." if outro["foto"] else "(vazio)",
                    fill=COR_APAGADO, font=("Segoe UI", 9), tags=("p%d" % i,),
                )
                tela.create_image(cx, cy, tags=("i%d" % i,))
                tela.create_rectangle(
                    x0 + 6, y0 + 6, x0 + 26, y0 + 24, fill=COR_BOTAO, outline="",
                    tags=("o%d" % i, "or%d" % i), state="hidden",
                )
                tela.create_text(
                    x0 + 16, y0 + 15, text="", fill=COR_TEXTO_FORTE, font=("Segoe UI Semibold", 9),
                    tags=("o%d" % i, "ot%d" % i), state="hidden",
                )
                tela.create_rectangle(
                    x0 + CELULA_L - 92, y0 + 6, x0 + CELULA_L - 6, y0 + 24, fill=COR_VERMELHO, outline="",
                    tags=("a%d" % i,), state="hidden",
                )
                tela.create_text(
                    x0 + CELULA_L - 49, y0 + 15, text="orientação ≠", fill=COR_FUNDO,
                    font=("Segoe UI Semibold", 8), tags=("a%d" % i,), state="hidden",
                )
                texto_cor = COR_SECUNDARIO if cor in (COR_BORDA, COR_APAGADO) else cor
                if outro["tipo"] == "buraco" and outro["foto"]:
                    texto_cor = COR_TEXTO
                tela.create_text(
                    x0 + CELULA_L // 2, y0 + 16 + MINI_TROCA + 10, text=self._rotulo_curto(outro, numeros),
                    fill=texto_cor, font=("Segoe UI Semibold", 9),
                )
                nome = outro["foto"].name if outro["foto"] else ""
                if len(nome) > 28:
                    nome = nome[:25] + "..."
                tela.create_text(
                    x0 + CELULA_L // 2, y0 + 16 + MINI_TROCA + 28, text=nome,
                    fill=COR_SECUNDARIO, font=("Segoe UI", 8),
                )
                estado["celulas"].append({"n": n, "item": outro, "x": x0, "y": y0})
                if outro["foto"]:
                    estado["por_foto"].setdefault(outro["foto"], []).append(i)
                if n == sel_n:
                    estado["sel"] = i
                marcar(i)
            linhas = (len(lista) + colunas - 1) // colunas
            tela.configure(scrollregion=(0, 0, largura, ESPACO + linhas * (CELULA_A + ESPACO)))
            ver_visiveis()

        def ver_visiveis():
            if not estado["colunas"]:
                return
            topo = tela.canvasy(0)
            fundo = topo + tela.winfo_height()
            passo = CELULA_A + ESPACO
            primeira = max(0, int(topo // passo) - 1) * estado["colunas"]
            ultima = (int(fundo // passo) + 2) * estado["colunas"]
            visiveis = [
                c["item"]["foto"] for c in estado["celulas"][primeira:ultima] if c["item"]["foto"]
            ]
            with trava:
                estado["visiveis"] = visiveis

        def rolou(primeiro, ultimo):
            barra.set(primeiro, ultimo)
            ver_visiveis()

        tela.configure(yscrollcommand=rolou)

        def proxima():
            with trava:
                faltam = estado["faltam"]
                for foto in estado["visiveis"]:
                    if foto in faltam:
                        del faltam[foto]
                        return foto
                if faltam:
                    foto = next(iter(faltam))
                    del faltam[foto]
                    return foto
            return None

        def trabalhador():
            while not parar.is_set():
                foto = proxima()
                if foto is None:
                    return
                try:
                    self._miniatura(foto)
                except OSError:
                    pass
                if parar.is_set():
                    return
                try:
                    self.root.after(0, lambda foto=foto: chegou(foto))
                except RuntimeError:
                    return

        def chegou(foto):
            if parar.is_set() or not janela.winfo_exists():
                return
            img = self._miniaturas.get((foto, MINI_TROCA))
            if img is not None and foto not in self._miniaturas_tk:
                self._miniaturas_tk[foto] = ImageTk.PhotoImage(img)
            estado["prontas"] = estado.get("prontas", 0) + 1
            restam = estado["total"] - estado["prontas"]
            progresso.set("Carregando miniaturas: faltam %s" % restam if restam > 0 else "")
            if foto == item["foto"]:
                for i in range(len(estado["celulas"])):
                    marcar(i)
            else:
                for i in estado["por_foto"].get(foto, []):
                    marcar(i)
            if so_compativel.get():
                agendar_desenho()

        def agendar_desenho(_evento=None):
            if estado["refazer"]:
                janela.after_cancel(estado["refazer"])
            estado["refazer"] = janela.after(150, refazer)

        def refazer():
            estado["refazer"] = None
            desenhar()

        def mudou_tamanho(evento):
            colunas = max(1, (max(evento.width, CELULA_L + 2 * ESPACO) - ESPACO) // (CELULA_L + ESPACO))
            if colunas != estado["colunas"]:
                agendar_desenho()
            else:
                ver_visiveis()

        def celula_em(evento):
            x, y = tela.canvasx(evento.x), tela.canvasy(evento.y)
            for i, c in enumerate(estado["celulas"]):
                if c["x"] <= x <= c["x"] + CELULA_L and c["y"] <= y <= c["y"] + CELULA_A:
                    return i
            return None

        def escolher(i):
            antiga, estado["sel"] = estado["sel"], i
            if antiga is not None and antiga < len(estado["celulas"]):
                marcar(antiga)
            marcar(i)
            mostrar(estado["celulas"][i]["item"])

        def clicou(evento):
            i = celula_em(evento)
            if i is not None:
                escolher(i)

        def duplo(evento):
            i = celula_em(evento)
            if i is not None:
                escolher(i)
                confirmar()

        def mostrar(destino):
            self._token_troca += 1
            token = self._token_troca
            partes = [self._rotulo_curto(destino, numeros)]
            if destino["foto"]:
                partes.append(destino["foto"].name)
            partes.append(destino["situacao"])
            if destino["tipo"] == "buraco":
                partes.append("buraco %s" % album.orientacao(destino["buraco"][2], destino["buraco"][3]))
            if destino["foto"] and destino["foto"] in self._orientacoes:
                partes.append("foto %s" % self._orientacoes[destino["foto"]])
            if incompativel(destino):
                partes.append("A orientação não combina.")
            if not destino["foto"]:
                partes.append("A foto escolhida vai para lá e este buraco fica vazio.")
            info.set("\n".join(partes))
            if not destino["foto"]:
                detalhe.configure(image=vazia)
                return
            foto = destino["foto"]

            def trabalho():
                try:
                    img = album.abrir_miniatura(foto, 300)
                except OSError:
                    img = None
                try:
                    self.root.after(0, lambda: pronto(img))
                except RuntimeError:
                    pass

            def pronto(img):
                if token != self._token_troca or not janela.winfo_exists():
                    if img is not None:
                        img.close()
                    return
                if img is None:
                    detalhe.configure(image=vazia)
                    return
                self._troca_tk = ImageTk.PhotoImage(img)
                img.close()
                detalhe.configure(image=self._troca_tk)

            threading.Thread(target=trabalho, daemon=True).start()

        def confirmar(_evento=None):
            if estado["sel"] is None:
                return
            self._confirmar_troca(item, estado["celulas"][estado["sel"]]["item"], janela, botao)

        def roda(evento):
            tela.yview_scroll(int(-evento.delta / 40), "units")

        def fechou(evento):
            if evento.widget is janela:
                parar.set()

        for valor, texto in (("todos", "Todos"), ("album", "No álbum"), ("fora", "Fora do álbum")):
            ttk.Radiobutton(
                filtros, text=texto, value=valor, variable=filtro, command=agendar_desenho,
            ).pack(side="left", padx=(0, 12))
        ttk.Checkbutton(
            filtros, text="Só orientação compatível", variable=so_compativel, command=agendar_desenho,
        ).pack(side="left", padx=(12, 0))
        ttk.Label(filtros, textvariable=progresso, style="Dica.TLabel").pack(side="right")

        tela.bind("<Configure>", mudou_tamanho)
        tela.bind("<ButtonPress-1>", clicou)
        tela.bind("<Double-1>", duplo)
        janela.bind("<MouseWheel>", roda)
        janela.bind("<Destroy>", fechou)
        janela.bind("<Return>", confirmar)
        janela.bind("<Escape>", lambda _e: janela.destroy())

        acoes = ttk.Frame(quadro)
        acoes.grid(row=3, column=0, columnspan=2, sticky="e", pady=(10, 0))
        ttk.Button(acoes, text="Cancelar", command=janela.destroy).pack(side="right")
        botao = ttk.Button(acoes, text="Trocar", style="Primario.TButton", command=confirmar)
        botao.pack(side="right", padx=6)

        fotos = [item["foto"]] + [o["foto"] for _n, o in candidatos() if o["foto"]]
        for foto in fotos:
            if (foto, MINI_TROCA) in self._miniaturas:
                if foto not in self._miniaturas_tk:
                    self._miniaturas_tk[foto] = ImageTk.PhotoImage(self._miniaturas[(foto, MINI_TROCA)])
            else:
                estado["faltam"][foto] = True
        estado["total"] = len(estado["faltam"])
        if estado["total"]:
            progresso.set("Carregando miniaturas: faltam %s" % estado["total"])
            for _ in range(3):
                threading.Thread(target=trabalhador, daemon=True).start()

        janela.update_idletasks()
        desenhar()
        janela.grab_set()
        tela.focus_set()

    def _confirmar_troca(self, item, destino, janela, botao):
        if self.rodando:
            return
        botao.configure(state="disabled")
        fotos = [f for f in (item["foto"], destino["foto"]) if f]

        def trabalho():
            orient = {}
            for foto in fotos:
                try:
                    orient[foto] = self._orientacao_foto(foto)
                except OSError:
                    orient[foto] = None
            self.root.after(0, lambda: self._fazer_troca(item, destino, orient, janela, botao))

        threading.Thread(target=trabalho, daemon=True).start()

    def _fazer_troca(self, item, destino, orient, janela, botao):
        if not janela.winfo_exists():
            return
        botao.configure(state="normal")
        nomes = {"vertical": "vertical", "horizontal": "horizontal", "quadrado": "quadrada"}
        avisos = []
        diferentes = []
        for foto, lugar in ((item["foto"], destino), (destino["foto"], item)):
            if not foto or lugar["tipo"] != "buraco":
                continue
            o_buraco = album.orientacao(lugar["buraco"][2], lugar["buraco"][3])
            o_foto = orient.get(foto)
            if o_foto and not album.aceita(o_buraco, o_foto):
                diferentes.append(lugar)
                avisos.append("A foto %s é %s e o buraco (%s) é %s." % (
                    foto.name, nomes.get(o_foto, o_foto), self._lugar_texto(lugar), o_buraco,
                ))
        if destino["tipo"] == "identificacao":
            avisos.append("A foto %s é a de identificação." % destino["foto"].name)
        elif destino["tipo"] == "olho_fechado":
            avisos.append("A foto %s está de olho fechado." % destino["foto"].name)
        if avisos and not messagebox.askyesno(
            "Álbum", "\n".join(avisos) + "\n\nTrocar mesmo assim?", parent=janela,
        ):
            return
        album.trocar_fotos(item, destino)
        for lugar in diferentes:
            if lugar["foto"]:
                lugar["situacao"] += ", orientação diferente"
        janela.destroy()

        self.selecionado = None
        self._token_foto += 1
        if self.imagem is not None:
            self.imagem.close()
        if self.imagem_tela is not None:
            self.imagem_tela.close()
        self.imagem = None
        self.imagem_tela = None
        self.caixa = None
        self.canvas.delete("all")
        self._preencher_listas()
        alvo = None
        for candidato in (item, destino):
            if candidato["tipo"] == "buraco" and candidato["foto"]:
                alvo = candidato
                break
        if alvo is None:
            self._previa_tk = ImageTk.PhotoImage(Image.new("RGB", (LADO_PREVIA, LADO_PREVIA), COR_FUNDO))
            self.previa.configure(image=self._previa_tk)
            self.dica_corte.set("O buraco ficou sem foto.")
            return
        iid = str(next(n for n, atual in enumerate(self.itens) if atual is alvo))
        self.lista_fotos.selection_set(iid)
        self.lista_fotos.see(iid)
        if self.tabela.exists(iid):
            self.tabela.selection_set(iid)
            self.tabela.see(iid)
        self._abrir_item(alvo)
        self.status.set("Trocou: %s agora está em %s." % (alvo["foto"].name, self._lugar_texto(alvo)))

    def _travar(self, rodando, parar=True):
        self.rodando = rodando
        estado = "disabled" if rodando else "normal"
        self.botao_gerar.configure(state=estado)
        self.botao_atualizar.configure(state=estado)
        self.botao_trocar.configure(state=estado)
        self.botao_parar.configure(state="normal" if rodando and parar else "disabled")

    def gerar(self):
        if self.rodando:
            return
        if not any(i["tipo"] == "buraco" and i["foto"] for i in self.itens):
            messagebox.showwarning("Álbum", "Atualize a lista das fotos antes de gerar.")
            return
        self._gravar_caminhos()
        self._limpar_log()
        self.parar_evento.clear()
        self._travar(True)
        self.status.set("Gerando...")
        itens = self.itens
        saida = self.saida.get().strip() or str(album.SAIDA_PADRAO)

        def ao_item(item):
            def ui():
                for n, atual in enumerate(self.itens):
                    if atual is item and self.tabela.exists(str(n)):
                        aviso = item["situacao"] in ("falta foto", "sobra, não entra", "erro ao abrir")
                        nota = item["situacao"].startswith("sem rosto")
                        tag = "aviso" if aviso else ("nota" if nota else "ok")
                        self.tabela.item(str(n), values=(
                            self._nome_pagina(item), item["indice"], item["foto"].name, item["situacao"],
                        ), tags=(tag,))
            self.root.after(0, ui)

        def trabalho():
            try:
                album.gerar(itens, saida, aviso=self._log, parar=self.parar_evento.is_set, ao_item=ao_item)
            except Exception as e:
                self._log("ERRO: " + str(e))
            self.root.after(0, lambda: self._terminou())

        threading.Thread(target=trabalho, daemon=True).start()

    def _terminou(self):
        self._travar(False)
        self.status.set("Concluído." if not self.parar_evento.is_set() else "Parado.")

    def parar(self):
        self.parar_evento.set()
        self.status.set("Parando...")

    def _fechar(self):
        self.parar_evento.set()
        self._gravar_caminhos()
        self.root.destroy()


def main():
    root = tk.Tk()
    JanelaAlbum(root)
    root.mainloop()


if __name__ == "__main__":
    main()
