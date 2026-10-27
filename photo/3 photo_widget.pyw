# -*- coding: utf-8 -*-
"""
Botão flutuante do exportador de JPG (photo.py).

Fica sempre por cima das outras janelas. Clique em LIGADO/DESLIGADO
para ativar a tecla de atalho, e na pasta para trocar onde salvar.
Botão direito abre o menu (abrir pasta, trocar tecla, tamanho, fechar).
Arraste pela área cinza para mudar de lugar.
"""

import os
import queue
import threading
import tkinter as tk
from tkinter import filedialog

import keyboard

import photo


COR_FUNDO = "#2b2b2b"
COR_TEXTO = "#dddddd"
COR_LIGADO = "#2e9e4f"
COR_DESLIGADO = "#6b6b6b"
COR_OK = "#1f6f3a"
COR_ERRO = "#8a2323"

TECLAS = ["F6", "F7", "F8", "F9", "F10", "F11", "F12"]


class Dica:
    """Tooltip simples ao passar o mouse."""

    def __init__(self, widget, texto_func):
        self.widget = widget
        self.texto_func = texto_func
        self.janela = None
        widget.bind("<Enter>", self._mostrar, add="+")
        widget.bind("<Leave>", self._esconder, add="+")

    def _mostrar(self, _evento):
        texto = self.texto_func()
        if not texto:
            return
        x = self.widget.winfo_rootx()
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
        self.janela = tk.Toplevel(self.widget)
        self.janela.overrideredirect(True)
        self.janela.attributes("-topmost", True)
        self.janela.geometry(f"+{x}+{y}")
        tk.Label(
            self.janela, text=texto, justify="left", wraplength=480,
            background="#ffffe0", relief="solid", borderwidth=1, padx=6, pady=3,
        ).pack()

    def _esconder(self, _evento):
        if self.janela:
            self.janela.destroy()
            self.janela = None


class WidgetExportador:

    def __init__(self, root):
        self.root = root
        self.config = photo.carregar_config()
        self.ligado = False
        self.atalho = None
        self.exportando = False
        self.contador = 0
        self.ultimo_erro = ""
        self.ultimo_caminho = ""
        self.fila = queue.Queue()
        self._arraste = (0, 0)

        root.overrideredirect(True)
        root.attributes("-topmost", True)
        root.configure(background=COR_FUNDO)

        x = self.config.get("x", root.winfo_screenwidth() - 300)
        y = self.config.get("y", 80)
        root.geometry(f"+{x}+{y}")

        self._montar()
        self._aplicar_compacto()

        if self.config.get("ligado"):
            self._ligar()
        else:
            self._atualizar_botao()

        root.after(100, self._ler_fila)

    # --------------------------------------------------------
    # Montagem
    # --------------------------------------------------------

    def _montar(self):
        root = self.root

        self.quadro = tk.Frame(root, background=COR_FUNDO, padx=6, pady=6)
        self.quadro.pack(fill="both", expand=True)

        topo = tk.Frame(self.quadro, background=COR_FUNDO)
        topo.pack(fill="x")

        self.botao = tk.Button(
            topo, width=14, font=("Segoe UI", 10, "bold"), foreground="white",
            relief="flat", activeforeground="white", command=self.alternar, cursor="hand2",
        )
        self.botao.pack(side="left")

        self.botao_pasta = tk.Button(
            topo, text="📁", font=("Segoe UI", 10), relief="flat", cursor="hand2",
            background="#444444", foreground="white", activebackground="#555555",
            command=self.trocar_pasta,
        )
        self.botao_pasta.pack(side="left", padx=(4, 0))
        Dica(self.botao_pasta, lambda: "Trocar pasta de salvamento")

        fechar = tk.Button(
            topo, text="✕", font=("Segoe UI", 9), relief="flat", cursor="hand2",
            background=COR_FUNDO, foreground="#aaaaaa", activebackground="#553333",
            command=self.fechar,
        )
        fechar.pack(side="right")

        self.detalhes = tk.Frame(self.quadro, background=COR_FUNDO)
        self.detalhes.pack(fill="x")

        self.rotulo_pasta = tk.Label(
            self.detalhes, anchor="w", background=COR_FUNDO, foreground="#aaaaaa",
            font=("Segoe UI", 8),
        )
        self.rotulo_pasta.pack(fill="x", pady=(4, 0))
        Dica(self.rotulo_pasta, lambda: self.config["pasta_saida"])

        self.rotulo_status = tk.Label(
            self.detalhes, anchor="w", background=COR_FUNDO, foreground=COR_TEXTO,
            font=("Segoe UI", 8), text="Nenhuma exportação ainda.",
        )
        self.rotulo_status.pack(fill="x", pady=(2, 0))
        Dica(self.rotulo_status, lambda: self.ultimo_erro or self.ultimo_caminho)

        self._atualizar_pasta()

        # Menu do botão direito
        self.menu = tk.Menu(root, tearoff=0)
        self.menu.add_command(label="Abrir pasta", command=self.abrir_pasta)
        self.menu.add_command(label="Trocar pasta...", command=self.trocar_pasta)

        self.var_tecla = tk.StringVar(value=self.config["tecla"])
        menu_teclas = tk.Menu(self.menu, tearoff=0)
        for tecla in TECLAS:
            menu_teclas.add_radiobutton(
                label=tecla, value=tecla, variable=self.var_tecla,
                command=lambda t=tecla: self.trocar_tecla(t),
            )
        self.menu.add_cascade(label="Tecla de atalho", menu=menu_teclas)

        self.var_compacto = tk.BooleanVar(value=bool(self.config.get("compacto")))
        self.menu.add_checkbutton(
            label="Compacto (só o botão)", variable=self.var_compacto,
            command=self._alternar_compacto,
        )
        self.menu.add_separator()
        self.menu.add_command(label="Fechar", command=self.fechar)

        # Arrastar e menu em toda a área que não é botão
        for w in (self.quadro, topo, self.detalhes, self.rotulo_pasta, self.rotulo_status):
            w.bind("<ButtonPress-1>", self._iniciar_arraste)
            w.bind("<B1-Motion>", self._arrastar)
            w.bind("<ButtonRelease-1>", lambda e: self._salvar_posicao())
        for w in (self.quadro, topo, self.detalhes, self.rotulo_pasta,
                  self.rotulo_status, self.botao, self.botao_pasta):
            w.bind("<Button-3>", self._abrir_menu)

    # --------------------------------------------------------
    # Liga / desliga
    # --------------------------------------------------------

    def alternar(self):
        if self.ligado:
            self._desligar()
        else:
            self._ligar()
        self.config["ligado"] = self.ligado
        self._salvar()

    def _ligar(self):
        try:
            self.atalho = keyboard.add_hotkey(self.config["tecla"], self._ao_apertar, suppress=False)
            self.ligado = True
        except Exception as e:
            self.ligado = False
            self._mostrar_status(False, f"Não foi possível ativar {self.config['tecla']}: {e}")
        self._atualizar_botao()

    def _desligar(self):
        if self.atalho is not None:
            try:
                keyboard.remove_hotkey(self.atalho)
            except Exception:
                pass
        self.atalho = None
        self.ligado = False
        self._atualizar_botao()

    def _atualizar_botao(self):
        tecla = self.config["tecla"]
        if self.ligado:
            self.botao.configure(text=f"LIGADO  ({tecla})", background=COR_LIGADO, activebackground=COR_LIGADO)
        else:
            self.botao.configure(text="DESLIGADO", background=COR_DESLIGADO, activebackground=COR_DESLIGADO)

    def trocar_tecla(self, tecla):
        estava_ligado = self.ligado
        if estava_ligado:
            self._desligar()
        self.config["tecla"] = tecla
        if estava_ligado:
            self._ligar()
        self._atualizar_botao()
        self._salvar()

    # --------------------------------------------------------
    # Exportação (roda fora da thread da janela)
    # --------------------------------------------------------

    def _ao_apertar(self):
        if self.exportando:
            return
        self.exportando = True
        pasta = self.config["pasta_saida"]
        threading.Thread(target=self._exportar, args=(pasta,), daemon=True).start()

    def _exportar(self, pasta):
        try:
            resultado = photo.exportar_jpg(pasta)
        except Exception as e:
            resultado = (False, f"{type(e).__name__}: {e}", None)
        self.fila.put(resultado)

    def _ler_fila(self):
        try:
            while True:
                ok, mensagem, caminho = self.fila.get_nowait()
                self.exportando = False
                if ok:
                    self.contador += 1
                    self.ultimo_caminho = caminho or ""
                    self.ultimo_erro = ""
                    self._mostrar_status(True, f"✓ {self.contador}  {mensagem}")
                else:
                    self._mostrar_status(False, "✗ " + mensagem)
        except queue.Empty:
            pass
        self.root.after(100, self._ler_fila)

    def _mostrar_status(self, ok, texto):
        if not ok:
            self.ultimo_erro = texto
        curto = texto if len(texto) <= 42 else texto[:40] + "…"
        self.rotulo_status.configure(text=curto, background=COR_OK if ok else COR_ERRO)
        self.root.after(900, lambda: self.rotulo_status.configure(background=COR_FUNDO))
        if not ok and self.var_compacto.get():
            self.botao.configure(background=COR_ERRO)
            self.root.after(900, self._atualizar_botao)

    # --------------------------------------------------------
    # Pasta
    # --------------------------------------------------------

    def trocar_pasta(self):
        self.root.attributes("-topmost", False)
        try:
            nova = filedialog.askdirectory(
                title="Pasta de salvamento dos JPG",
                initialdir=self.config["pasta_saida"],
                parent=self.root,
            )
        finally:
            self.root.attributes("-topmost", True)
        if nova:
            self.config["pasta_saida"] = os.path.normpath(nova)
            self._atualizar_pasta()
            self._salvar()

    def abrir_pasta(self):
        try:
            os.startfile(self.config["pasta_saida"])
        except Exception as e:
            self._mostrar_status(False, f"Não abriu a pasta: {e}")

    def _atualizar_pasta(self):
        pasta = self.config["pasta_saida"]
        partes = pasta.rstrip("\\/").replace("/", "\\").split("\\")
        curto = "\\".join(partes[-2:]) if len(partes) > 2 else pasta
        self.rotulo_pasta.configure(text="📂 …\\" + curto if len(partes) > 2 else "📂 " + curto)

    # --------------------------------------------------------
    # Janela
    # --------------------------------------------------------

    def _alternar_compacto(self):
        self.config["compacto"] = self.var_compacto.get()
        self._aplicar_compacto()
        self._salvar()

    def _aplicar_compacto(self):
        if self.config.get("compacto"):
            self.detalhes.pack_forget()
        elif not self.detalhes.winfo_ismapped():
            self.detalhes.pack(fill="x")

    def _abrir_menu(self, evento):
        self.menu.tk_popup(evento.x_root, evento.y_root)

    def _iniciar_arraste(self, evento):
        self._arraste = (evento.x_root - self.root.winfo_x(), evento.y_root - self.root.winfo_y())

    def _arrastar(self, evento):
        dx, dy = self._arraste
        self.root.geometry(f"+{evento.x_root - dx}+{evento.y_root - dy}")

    def _salvar_posicao(self):
        self.config["x"] = self.root.winfo_x()
        self.config["y"] = self.root.winfo_y()
        self._salvar()

    def _salvar(self):
        try:
            photo.salvar_config(self.config)
        except Exception:
            pass

    def fechar(self):
        self._desligar()
        self.config["x"] = self.root.winfo_x()
        self.config["y"] = self.root.winfo_y()
        self._salvar()
        self.root.destroy()


def ja_esta_aberto():
    """Evita dois botões abertos (o atalho exportaria duas vezes)."""

    import ctypes

    kernel32 = ctypes.windll.kernel32
    main.mutex = kernel32.CreateMutexW(None, False, "PapaExportadorJPG_Widget")
    return kernel32.GetLastError() == 183  # ERROR_ALREADY_EXISTS


def main():
    if ja_esta_aberto():
        return
    root = tk.Tk()
    root.title("Exportador JPG")
    WidgetExportador(root)
    root.mainloop()


if __name__ == "__main__":
    main()
