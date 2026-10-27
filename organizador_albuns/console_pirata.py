# -*- coding: utf-8 -*-
"""
CMD da separação do lote.

Roda o comando que vem depois de "--", grava cada linha no log
(a janela do organizador acompanha esse arquivo) e redesenha a tela
com a saída em verde por cima de uma caveira pirata em cinza.
"""

import argparse
import ctypes
import os
import shutil
import subprocess
import sys
import textwrap
import threading
import time
from collections import deque


CAVEIRA = textwrap.dedent(r'''
           ______
        .-"      "-.
       /            \
      |              |
      |,  .-.  .-.  ,|
      | )(_ /  \_ )( |
      |/     /\     \|
      (_     ^^     _)
       \__|IIIIII|__/
        | \IIIIII/ |
        \          /
     jgs `--------`
    .-.            .-.
   ((  `-.      .-`  ))
    \ `-. `.__.` .-` /
     `-. `------` .-`
        `--------`
''').strip("\n").splitlines()

OSSOS = 5


def ampliar(desenho, fator):
    if fator <= 1:
        return desenho
    grande = []
    for linha in desenho:
        larga = "".join(letra * fator for letra in linha)
        grande.extend([larga] * fator)
    return grande


def escolher_desenho(colunas, altura):
    for desenho in (CAVEIRA, CAVEIRA[:-OSSOS]):
        largura = max(len(linha) for linha in desenho)
        for fator in (3, 2, 1):
            if len(desenho) * fator <= altura and largura * fator <= colunas:
                return ampliar(desenho, fator)
    return CAVEIRA[:-OSSOS]

VERDE = "\x1b[92m"
CINZA = "\x1b[90m"
AMARELO = "\x1b[93m"
NORMAL = "\x1b[0m"
INICIO = "\x1b[H"
LIMPAR = "\x1b[2J"
SEM_CURSOR = "\x1b[?25l"
COM_CURSOR = "\x1b[?25h"

INTERVALO = 0.05


def ativar_cores():
    if os.name != "nt":
        return
    kernel32 = ctypes.windll.kernel32
    kernel32.SetConsoleOutputCP(65001)
    saida = kernel32.GetStdHandle(-11)
    modo = ctypes.c_uint32()
    if kernel32.GetConsoleMode(saida, ctypes.byref(modo)):
        kernel32.SetConsoleMode(saida, modo.value | 0x0004)


def posicionar(titulo, x, y, largura, altura):
    if os.name != "nt":
        return
    kernel32 = ctypes.windll.kernel32
    kernel32.SetConsoleTitleW(titulo)
    if largura <= 0 or altura <= 0:
        return
    janela = kernel32.GetConsoleWindow()
    if not janela:
        return
    user32 = ctypes.windll.user32
    for _ in range(3):
        user32.ShowWindow(janela, 9)
        user32.MoveWindow(janela, x, y, largura, altura, True)
        time.sleep(0.1)


def limpar_linha(texto):
    partes = []
    for letra in texto:
        codigo = ord(letra)
        if codigo in (0xFE0F, 0x200D):
            continue
        if codigo >= 0x1F000 or 0x2600 <= codigo <= 0x27BF or 0x2B00 <= codigo <= 0x2BFF:
            partes.append("*")
        elif letra == "\t":
            partes.append("    ")
        elif codigo < 32:
            continue
        else:
            partes.append(letra)
    return "".join(partes)


class Tela:

    def __init__(self):
        self.linhas = deque(maxlen=1000)
        self.trava = threading.Lock()
        self.mudou = True
        self.total = 0

    def adicionar(self, linha):
        with self.trava:
            self.linhas.append(limpar_linha(linha))
            self.total += 1
            self.mudou = True

    def desenhar(self, rodape):
        with self.trava:
            self.mudou = False
            tamanho = shutil.get_terminal_size((100, 30))
            colunas = max(tamanho.columns - 1, 20)
            altura = max(tamanho.lines - 1, 5)
            visiveis = list(self.linhas)[-altura:]

        desenho = escolher_desenho(colunas, altura)
        largura_desenho = max(len(linha) for linha in desenho)
        esquerda = max((colunas - largura_desenho) // 2, 0)
        topo = max((altura - len(desenho)) // 2, 0)

        partes = [INICIO]
        for numero in range(altura):
            indice = numero - topo
            fundo = desenho[indice] if 0 <= indice < len(desenho) else ""
            fundo = (" " * esquerda + fundo).ljust(colunas)[:colunas]
            texto = visiveis[numero][:colunas] if numero < len(visiveis) else ""
            partes.append(VERDE + texto + CINZA + fundo[len(texto):] + NORMAL + "\r\n")
        partes.append(AMARELO + rodape[:colunas].ljust(colunas) + NORMAL)
        sys.stdout.write("".join(partes))
        sys.stdout.flush()


def ler_saida(processo, tela, log):
    for bruto in iter(processo.stdout.readline, b""):
        linha = bruto.decode("utf-8", errors="replace").rstrip("\r\n")
        log.write(linha + "\n")
        log.flush()
        tela.adicionar(linha)


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    parser = argparse.ArgumentParser(description="CMD pirata da separação do lote")
    parser.add_argument("--log", required=True)
    parser.add_argument("--x", type=int, default=0)
    parser.add_argument("--y", type=int, default=0)
    parser.add_argument("--largura", type=int, default=0)
    parser.add_argument("--altura", type=int, default=0)
    parser.add_argument("--titulo", default="Separando o lote")
    parser.add_argument("comando", nargs=argparse.REMAINDER)
    args = parser.parse_args()

    comando = args.comando[1:] if args.comando[:1] == ["--"] else args.comando

    ativar_cores()
    posicionar(args.titulo, args.x, args.y, args.largura, args.altura)
    sys.stdout.write(SEM_CURSOR + LIMPAR)

    tela = Tela()
    log = open(args.log, "w", encoding="utf-8")
    codigo = 1

    try:
        ambiente = dict(os.environ)
        ambiente["PYTHONIOENCODING"] = "utf-8"
        ambiente["PYTHONUNBUFFERED"] = "1"
        processo = subprocess.Popen(
            comando,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            stdin=subprocess.DEVNULL,
            env=ambiente,
        )
        leitor = threading.Thread(target=ler_saida, args=(processo, tela, log), daemon=True)
        leitor.start()

        inicio = time.time()
        while leitor.is_alive():
            if tela.mudou:
                segundos = int(time.time() - inicio)
                tela.desenhar(f" SEPARANDO O LOTE...  linhas: {tela.total}  tempo: {segundos // 60:02d}:{segundos % 60:02d}")
            time.sleep(INTERVALO)
        leitor.join()
        codigo = processo.wait()
    except Exception as erro:
        mensagem = f"Não foi possível iniciar o processo: {erro}"
        log.write(mensagem + "\n")
        tela.adicionar(mensagem)
    finally:
        log.write(f"__FIM__ {codigo}\n")
        log.close()

    situacao = "TERMINOU" if codigo == 0 else f"TERMINOU COM ERRO ({codigo})"
    tela.desenhar(f" {situacao}  |  Pressione ENTER para fechar")
    sys.stdout.write(COM_CURSOR)
    sys.stdout.flush()
    try:
        input()
    except EOFError:
        pass
    return codigo


if __name__ == "__main__":
    sys.exit(main())
