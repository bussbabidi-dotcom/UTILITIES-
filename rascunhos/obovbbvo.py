"""Calculadora simples com interface gráfica (tkinter)."""

import tkinter as tk
from tkinter import font as tkfont


class Calculadora:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Calculadora")
        self.root.resizable(False, False)
        self.root.configure(bg="#f0f0f0")

        self.expressao = ""
        self.fonte_display = tkfont.Font(family="Segoe UI", size=20)
        self.fonte_botao = tkfont.Font(family="Segoe UI", size=14)

        self._criar_display()
        self._criar_botoes()

    def _criar_display(self) -> None:
        # Campo que mostra a expressão e o resultado
        self.display = tk.Entry(
            self.root,
            font=self.fonte_display,
            justify="right",
            bd=8,
            relief="sunken",
            bg="#ffffff",
            fg="#222222",
            insertbackground="#222222",
        )
        self.display.grid(row=0, column=0, columnspan=4, sticky="nsew", padx=10, pady=10)
        self.display.insert(0, "0")
        self.display.config(state="readonly")

    def _criar_botoes(self) -> None:
        # Layout: linhas de botões numéricos e operações
        botoes = [
            ("C", 1, 0), ("/", 1, 1), ("*", 1, 2), ("-", 1, 3),
            ("7", 2, 0), ("8", 2, 1), ("9", 2, 2), ("+", 2, 3),
            ("4", 3, 0), ("5", 3, 1), ("6", 3, 2), ("=", 3, 3),
            ("1", 4, 0), ("2", 4, 1), ("3", 4, 2),
            ("0", 5, 0), (".", 5, 1),
        ]

        for texto, linha, coluna in botoes:
            if texto == "=":
                # O botão "=" ocupa duas linhas
                btn = tk.Button(
                    self.root,
                    text=texto,
                    font=self.fonte_botao,
                    command=self._calcular,
                    bg="#4a90d9",
                    fg="white",
                    activebackground="#3a7bc8",
                    activeforeground="white",
                    width=5,
                    height=3,
                )
                btn.grid(row=linha, column=coluna, rowspan=2, sticky="nsew", padx=4, pady=4)
            elif texto == "C":
                btn = tk.Button(
                    self.root,
                    text=texto,
                    font=self.fonte_botao,
                    command=self._limpar,
                    bg="#e74c3c",
                    fg="white",
                    activebackground="#c0392b",
                    activeforeground="white",
                    width=5,
                    height=2,
                )
                btn.grid(row=linha, column=coluna, sticky="nsew", padx=4, pady=4)
            elif texto == "0":
                btn = tk.Button(
                    self.root,
                    text=texto,
                    font=self.fonte_botao,
                    command=lambda t=texto: self._adicionar(t),
                    bg="#ffffff",
                    fg="#222222",
                    width=5,
                    height=2,
                )
                btn.grid(row=linha, column=coluna, columnspan=2, sticky="nsew", padx=4, pady=4)
            elif texto == ".":
                btn = tk.Button(
                    self.root,
                    text=texto,
                    font=self.fonte_botao,
                    command=lambda t=texto: self._adicionar(t),
                    bg="#ffffff",
                    fg="#222222",
                    width=5,
                    height=2,
                )
                btn.grid(row=linha, column=2, sticky="nsew", padx=4, pady=4)
            else:
                # Números e operadores
                cor_bg = "#e8e8e8" if texto in "+-*/" else "#ffffff"
                btn = tk.Button(
                    self.root,
                    text=texto,
                    font=self.fonte_botao,
                    command=lambda t=texto: self._adicionar(t),
                    bg=cor_bg,
                    fg="#222222",
                    width=5,
                    height=2,
                )
                btn.grid(row=linha, column=coluna, sticky="nsew", padx=4, pady=4)

        for i in range(6):
            self.root.grid_rowconfigure(i, weight=1)
        for i in range(4):
            self.root.grid_columnconfigure(i, weight=1)

    def _atualizar_display(self, texto: str) -> None:
        self.display.config(state="normal")
        self.display.delete(0, tk.END)
        self.display.insert(0, texto)
        self.display.config(state="readonly")

    def _adicionar(self, valor: str) -> None:
        # Evita começar com operador (exceto menos para números negativos simples via expressão)
        if self.expressao in ("Erro", "Divisão por zero"):
            self.expressao = ""

        if valor == ".":
            # Impede mais de um ponto no número atual
            ultimo_numero = self._ultimo_numero()
            if "." in ultimo_numero:
                return
            if not ultimo_numero:
                self.expressao += "0"

        if valor in "+-*/" and self.expressao and self.expressao[-1] in "+-*/":
            # Substitui o operador anterior
            self.expressao = self.expressao[:-1] + valor
        else:
            self.expressao += valor

        self._atualizar_display(self.expressao)

    def _ultimo_numero(self) -> str:
        for i in range(len(self.expressao) - 1, -1, -1):
            if self.expressao[i] in "+-*/":
                return self.expressao[i + 1 :]
        return self.expressao

    def _limpar(self) -> None:
        self.expressao = ""
        self._atualizar_display("0")

    def _calcular(self) -> None:
        if not self.expressao or self.expressao in ("Erro", "Divisão por zero"):
            return

        # Remove operador pendente no final
        expressao = self.expressao
        while expressao and expressao[-1] in "+-*/":
            expressao = expressao[:-1]

        if not expressao:
            return

        try:
            # Avalia apenas expressões numéricas simples (+ - * /)
            resultado = eval(expressao, {"__builtins__": {}}, {})
            if isinstance(resultado, float) and resultado.is_integer():
                resultado = int(resultado)
            self.expressao = str(resultado)
            self._atualizar_display(self.expressao)
        except ZeroDivisionError:
            self.expressao = "Divisão por zero"
            self._atualizar_display(self.expressao)
        except Exception:
            self.expressao = "Erro"
            self._atualizar_display(self.expressao)


def main() -> None:
    root = tk.Tk()
    Calculadora(root)
    root.mainloop()


if __name__ == "__main__":
    main()
