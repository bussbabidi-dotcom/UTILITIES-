# -*- coding: utf-8 -*-
"""
Monta as páginas do álbum em JPG.

O molde das 24 páginas fica na pasta molde/, em PNG com os buracos
transparentes. A foto entra por baixo do PNG, então texto e moldura ficam
por cima. De cada pose da criança entra uma foto; a mesma pose não se
repete na mesma página. Foto vertical só entra em buraco vertical e foto
horizontal só em buraco horizontal; buraco quadrado aceita as duas.
A foto de identificação (com a ficha e o número) não entra, nem a foto com
olho fechado (sem o reflexo do flash na íris e com a abertura pequena).
O álbum segue um cenário (fundo) de cada vez e, dentro dele, as roupas na
ordem da sessão, uma roupa por página. As páginas 1, 2, 12 e 24 ficam no
lugar; as outras só trocam com a vizinha (no máximo uma posição do molde),
e o JPG sai com a posição na frente do nome. Fotos em que a criança não olha para a
câmera ficam por último.
Com óculos escuros não dá para ver os olhos: a foto entra normalmente, mas
na mesma pose a de olhos à mostra vem primeiro.
"""

import hashlib
import itertools
import json
import os
import re
import sys
import threading
from collections import Counter
from pathlib import Path

import cv2
import numpy as np
from PIL import Image, ImageOps


def pasta_programa():
    """Pasta do programa: ao lado do exe, ou ao lado deste arquivo."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _pasta_modelo(nome="face_detection_yunet_2023mar.onnx"):
    ao_lado = pasta_programa() / "modelos" / nome
    if ao_lado.is_file():
        return ao_lado
    interno = getattr(sys, "_MEIPASS", None)
    if interno:
        embutido = Path(interno) / "modelos" / nome
        if embutido.is_file():
            return embutido
    return ao_lado


PASTA = pasta_programa()
MODELO_YUNET = _pasta_modelo()
FICHA_REFERENCIA = _pasta_modelo("ficha.png")
MANUAL_FILE = PASTA / "recorte_manual.json"
CONFIG_FILE = PASTA / "album_config.json"
SAIDA_PADRAO = PASTA / "saida"
PASTA_MINIATURAS = Path(os.environ.get("LOCALAPPDATA") or Path.home()) / "album_molde" / "miniaturas"

EXTENSOES = {".jpg", ".jpeg", ".png", ".tif", ".tiff"}
LADO_DETECCAO = 1600
LIMIAR_ALPHA = 128
FOLGA_BURACO = 2
MIN_LADO = 80
MIN_AREA = 12000
PROPORCAO_MAXIMA = 6.0
QUALIDADE_JPG = 95
LIMIAR_POSE = 0.82
LIMIAR_CORPO = 0.82
PAGINA_GRANDE = 3
PAGINAS_FIXAS = (1, 2, 12, 24)
DISTANCIA_MAXIMA = 1
PESO_FALTA = 1e6
PESO_FOTO_DUPLICADA = 1e5
PESO_ROUPA_AUSENTE = 1e4
PESO_ABAIXO_MINIMO = 1e3
PESO_ORIENTACAO = 1e3
PESO_POSE = 100
PESO_REPETIDA = 10
PESO_NAO_OLHA = 0.1
PESO_TROCA = 0.01
TOLERANCIA_QUADRADO = 0.1
LADO_ANALISE = 900
MIN_PONTOS_FICHA = 12
LIMIAR_CENARIO = 60.0
LIMIAR_ROUPA = 0.8
CONFIRMA_ROUPA = 3
SALTO_NUMERACAO = 5
LIMIAR_OLHANDO = 0.25
INCLINACAO_MAXIMA = 0.35
TAMANHO_OLHO = (96, 66)
LIMIAR_BRILHO_OLHO = 0.5
ABERTURA_MINIMA = 0.29
LIMIAR_OCULOS = 0.6

CALIBRACAO_PADRAO = {
    "centro_x": 0.5,
    "olhos_y": 0.33,
    "altura_rosto": 0.16,
}
ROSTO_MAXIMO = 0.4

_detector = None
_haar = None
_rostos = {}
_trava_manual = threading.Lock()
_trava_rosto = threading.Lock()


def chave_natural(caminho):
    partes = re.split(r"(\d+)", Path(caminho).name)
    return [int(p) if p.isdigit() else p.casefold() for p in partes]


def listar_imagens(pasta):
    pasta = Path(pasta)
    if not pasta.is_dir():
        return []
    arquivos = [
        f for f in pasta.iterdir()
        if f.is_file() and f.suffix.lower() in EXTENSOES and not f.name.startswith(".")
    ]
    return sorted(arquivos, key=chave_natural)


def carregar_config():
    dados = {}
    if CONFIG_FILE.is_file():
        try:
            dados = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            dados = {}
    calib = dict(CALIBRACAO_PADRAO)
    calib.update(dados.get("calibracao") or {})
    dados["calibracao"] = calib
    return dados


def gravar_config(alteracoes):
    dados = {}
    if CONFIG_FILE.is_file():
        try:
            dados = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            dados = {}
    dados.update(alteracoes)
    CONFIG_FILE.write_text(json.dumps(dados, indent=2, ensure_ascii=False), encoding="utf-8")


def calibracao():
    return carregar_config()["calibracao"]


def _pngs(pasta):
    if not pasta.is_dir():
        return []
    return sorted(
        (f for f in pasta.iterdir()
         if f.is_file() and f.suffix.lower() == ".png" and not f.name.startswith(".")),
        key=chave_natural,
    )


def pasta_molde():
    """Pasta com os PNGs do molde: ao lado do programa ou dentro do exe."""
    ao_lado = pasta_programa() / "molde"
    if _pngs(ao_lado):
        return ao_lado
    interno = getattr(sys, "_MEIPASS", None)
    if interno:
        embutido = Path(interno) / "molde"
        if _pngs(embutido):
            return embutido
    return ao_lado


def carregar_molde_fixo():
    pasta = pasta_molde()
    paginas = _pngs(pasta)
    if not paginas:
        raise FileNotFoundError("Nenhum molde PNG em " + str(pasta))
    return [{"arquivo": caminho, "buracos": buracos_da_pagina(caminho)} for caminho in paginas]


# ============================================================
# BURACOS TRANSPARENTES
# ============================================================

def ordenar_buracos(buracos):
    """De cima para baixo e, na mesma faixa, da esquerda para a direita."""
    if not buracos:
        return []
    alturas = sorted(b[3] for b in buracos)
    tolerancia = max(20, alturas[len(alturas) // 2] * 0.25)
    linhas = []
    for buraco in sorted(buracos, key=lambda b: (b[1], b[0])):
        if linhas and abs(buraco[1] - linhas[-1][0][1]) <= tolerancia:
            linhas[-1].append(buraco)
        else:
            linhas.append([buraco])
    saida = []
    for linha in linhas:
        saida.extend(sorted(linha, key=lambda b: b[0]))
    return saida


def achar_buracos_alpha(img):
    """Áreas transparentes do PNG, como caixas (x, y, w, h)."""
    alpha = np.asarray(img.getchannel("A"))
    altura, largura = alpha.shape[:2]
    vazio = (alpha < LIMIAR_ALPHA).astype(np.uint8)
    n, _labels, stats, _ = cv2.connectedComponentsWithStats(vazio, connectivity=4)
    buracos = []
    for i in range(1, n):
        x, y, w, h, area = (int(v) for v in stats[i])
        if w < MIN_LADO or h < MIN_LADO or area < MIN_AREA:
            continue
        if max(w, h) / float(min(w, h)) > PROPORCAO_MAXIMA:
            continue
        x = max(0, x - FOLGA_BURACO)
        y = max(0, y - FOLGA_BURACO)
        w = min(largura - x, w + 2 * FOLGA_BURACO)
        h = min(altura - y, h + 2 * FOLGA_BURACO)
        buracos.append((x, y, w, h))
    return ordenar_buracos(buracos)


_buracos_cache = {}


def buracos_da_pagina(caminho):
    chave = (str(Path(caminho).resolve()), Path(caminho).stat().st_mtime_ns)
    if chave not in _buracos_cache:
        img, _icc, _dpi = abrir_molde(caminho)
        try:
            _buracos_cache[chave] = achar_buracos_alpha(img)
        finally:
            img.close()
    return _buracos_cache[chave]


# ============================================================
# ROSTO
# ============================================================

def detector_yunet():
    global _detector
    if _detector is None:
        if not MODELO_YUNET.is_file():
            raise FileNotFoundError(
                "Modelo de rosto não encontrado: " + str(MODELO_YUNET)
            )
        _detector = cv2.FaceDetectorYN.create(
            str(MODELO_YUNET), "", (320, 320), 0.7, 0.3, 5000
        )
    return _detector


def detectar_rosto(img):
    """Maior rosto em coordenadas da imagem original, ou None."""
    w, h = img.size
    escala = min(1.0, LADO_DETECCAO / max(w, h))
    pequena = img.convert("RGB")
    if escala < 1.0:
        pequena = pequena.resize((round(w * escala), round(h * escala)), Image.BILINEAR)

    bgr = cv2.cvtColor(np.asarray(pequena), cv2.COLOR_RGB2BGR)
    ph, pw = bgr.shape[:2]

    with _trava_rosto:
        det = detector_yunet()
        det.setInputSize((pw, ph))
        _ok, faces = det.detect(bgr)

    if faces is not None and len(faces):
        f = max(faces, key=lambda r: r[2] * r[3])
        x, y, fw, fh = (float(v) / escala for v in f[:4])
        olho_d_x, olho_d_y, olho_e_x, olho_e_y, nariz_x, nariz_y = (
            float(v) / escala for v in f[4:10]
        )
        return {
            "x": x, "y": y, "w": fw, "h": fh,
            "cx": x + fw / 2,
            "olhos_y": (olho_d_y + olho_e_y) / 2,
            "olho_d": (olho_d_x, olho_d_y),
            "olho_e": (olho_e_x, olho_e_y),
            "nariz": (nariz_x, nariz_y),
            "metodo": "yunet",
        }

    global _haar
    if not hasattr(cv2, "CascadeClassifier"):
        return None
    if _haar is None:
        _haar = cv2.CascadeClassifier(
            cv2.data.haarcascades + "haarcascade_frontalface_default.xml"
        )
    cinza = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
    caixas = _haar.detectMultiScale(cinza, 1.1, 6, minSize=(40, 40))
    if len(caixas):
        x, y, fw, fh = (float(v) / escala for v in max(caixas, key=lambda r: r[2] * r[3]))
        return {
            "x": x, "y": y, "w": fw, "h": fh * 0.9,
            "cx": x + fw / 2,
            "olhos_y": y + fh * 0.4,
            "metodo": "haar",
        }
    return None


def rosto_da_foto(caminho, img):
    st = Path(caminho).stat()
    chave = (str(Path(caminho).resolve()), st.st_mtime_ns)
    if chave not in _rostos:
        _rostos[chave] = detectar_rosto(img)
    return _rostos[chave]


def abrir_imagem(caminho):
    img = Image.open(caminho)
    icc = img.info.get("icc_profile")
    dpi = img.info.get("dpi")
    img = ImageOps.exif_transpose(img)
    if img.mode != "RGB":
        img = img.convert("RGB")
    return img, icc, dpi


def abrir_molde(caminho):
    """PNG do molde em RGBA, com a transparência dos buracos."""
    img = Image.open(caminho)
    icc = img.info.get("icc_profile")
    dpi = img.info.get("dpi")
    if img.mode != "RGBA":
        img = img.convert("RGBA")
    return img, icc, dpi


def compor_pagina(molde_rgba, colocadas):
    """Fotos numa base branca e o molde por cima. colocadas: (x, y, foto RGB)."""
    base = Image.new("RGBA", molde_rgba.size, (255, 255, 255, 255))
    for x, y, foto in colocadas:
        base.paste(foto, (x, y))
    base.alpha_composite(molde_rgba)
    return base.convert("RGB")


def calcular_recorte(img_size, rosto, molde, calib):
    """(left, top, right, bottom) na foto original."""
    W, H = img_size
    mw, mh = molde
    olhos = float(calib["olhos_y"])
    desejado = rosto["h"] / float(calib["altura_rosto"])
    crop_h = min(desejado, H, W * mh / float(mw))
    cabe = min(rosto["olhos_y"] / olhos, (H - rosto["olhos_y"]) / (1.0 - olhos))
    crop_h = max(min(crop_h, cabe), min(crop_h, rosto["h"] / ROSTO_MAXIMO))
    crop_w = crop_h * mw / mh
    left = rosto["cx"] - float(calib["centro_x"]) * crop_w
    top = rosto["olhos_y"] - olhos * crop_h
    left = min(max(0.0, left), W - crop_w)
    top = min(max(0.0, top), H - crop_h)
    return (
        int(round(left)),
        int(round(top)),
        int(round(left + crop_w)),
        int(round(top + crop_h)),
    )


def ajustar_caixa(caixa, tamanho_fonte, molde):
    """Força a proporção do buraco e mantém a caixa dentro da foto."""
    W, H = tamanho_fonte
    mw, mh = molde
    l, t, r, b = caixa
    cx, cy = (l + r) / 2, (t + b) / 2
    h = max(10.0, b - t)
    w = h * mw / float(mh)
    fator = min(1.0, W / w, H / h)
    w *= fator
    h *= fator
    l = min(max(0.0, cx - w / 2), W - w)
    t = min(max(0.0, cy - h / 2), H - h)
    return (int(round(l)), int(round(t)), int(round(l + w)), int(round(t + h)))


def recortar_imagem(img, caixa, molde):
    recortada = img.crop(caixa).resize(molde, Image.LANCZOS)
    if recortada.mode != "RGB":
        recortada = recortada.convert("RGB")
    return recortada


def chave_ajuste(foto, pagina, indice, molde):
    return "|".join([
        os.path.normcase(os.path.abspath(str(foto))),
        os.path.normcase(os.path.abspath(str(pagina))),
        str(indice),
        f"{molde[0]}x{molde[1]}",
    ])


def ler_manual():
    if not MANUAL_FILE.is_file():
        return {}
    try:
        return json.loads(MANUAL_FILE.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}


def _gravar_manual(dados):
    MANUAL_FILE.parent.mkdir(parents=True, exist_ok=True)
    temp = MANUAL_FILE.with_suffix(".tmp")
    temp.write_text(json.dumps(dados, indent=2, ensure_ascii=False), encoding="utf-8")
    os.replace(temp, MANUAL_FILE)


def gravar_manual(foto, pagina, indice, molde, caixa, tamanho_fonte):
    W, H = tamanho_fonte
    chave = chave_ajuste(foto, pagina, indice, molde)
    with _trava_manual:
        dados = ler_manual()
        dados[chave] = {
            "caixa": [
                round(caixa[0] / W, 6), round(caixa[1] / H, 6),
                round(caixa[2] / W, 6), round(caixa[3] / H, 6),
            ],
            "origem": [W, H],
        }
        _gravar_manual(dados)


def apagar_manual(foto, pagina, indice, molde):
    chave = chave_ajuste(foto, pagina, indice, molde)
    with _trava_manual:
        dados = ler_manual()
        if dados.pop(chave, None) is not None:
            _gravar_manual(dados)


def caixa_manual(foto, pagina, indice, molde, tamanho_fonte):
    chave = chave_ajuste(foto, pagina, indice, molde)
    with _trava_manual:
        reg = ler_manual().get(chave)
    if not reg:
        return None
    W, H = tamanho_fonte
    l, t, r, b = reg["caixa"]
    return ajustar_caixa((l * W, t * H, r * W, b * H), tamanho_fonte, molde)


def caixa_da_foto(foto, pagina, indice, molde, img):
    """
    (caixa, origem) com origem em manual, rosto ou centro.
    """
    manual = caixa_manual(foto, pagina, indice, molde, img.size)
    if manual:
        return manual, "manual"
    rosto = rosto_da_foto(foto, img)
    if rosto:
        return calcular_recorte(img.size, rosto, molde, calibracao()), "rosto " + rosto["metodo"]
    return ajustar_caixa((0, 0, img.width, img.height), img.size, molde), "sem rosto, centro"


# ============================================================
# PRÉVIA RÁPIDA DA PÁGINA
# ============================================================

LADO_PREVIA_PAGINA = 900
_paginas_reduzidas = {}
_recortes_rapidos = {}
_trava_previa = threading.Lock()


def pagina_reduzida(caminho):
    """(imagem, escala) da página do molde já reduzida, guardada na memória."""
    chave = str(caminho)
    with _trava_previa:
        if chave in _paginas_reduzidas:
            return _paginas_reduzidas[chave]
    original, _icc, _dpi = abrir_molde(caminho)
    try:
        escala = min(1.0, LADO_PREVIA_PAGINA / float(max(original.size)))
        img = original.resize(
            (max(1, round(original.width * escala)), max(1, round(original.height * escala))),
            Image.BILINEAR,
        )
    finally:
        original.close()
    resultado = (img, escala)
    with _trava_previa:
        _paginas_reduzidas[chave] = resultado
    return resultado


def recorte_rapido(foto, pagina, indice, molde, tamanho):
    """
    Foto já cortada para a prévia, a partir da miniatura.
    O rosto daqui não entra no cache da foto grande.
    """
    chave = chave_ajuste(foto, pagina, indice, molde)
    with _trava_previa:
        guardado = _recortes_rapidos.get(chave)
    if guardado is not None and guardado.size == tamanho:
        return guardado
    img = abrir_miniatura(foto, 900)
    try:
        caixa = caixa_manual(foto, pagina, indice, molde, img.size)
        if not caixa:
            rosto = detectar_rosto(img)
            if rosto:
                caixa = calcular_recorte(img.size, rosto, molde, calibracao())
            else:
                caixa = ajustar_caixa((0, 0, img.width, img.height), img.size, molde)
        recorte = img.crop(caixa).resize(tamanho, Image.BILINEAR)
    finally:
        img.close()
    with _trava_previa:
        _recortes_rapidos[chave] = recorte
    return recorte


def tem_recorte_rapido(foto, pagina, indice, molde, tamanho):
    with _trava_previa:
        guardado = _recortes_rapidos.get(chave_ajuste(foto, pagina, indice, molde))
    return guardado is not None and guardado.size == tamanho


def esquecer_recorte_rapido(foto, pagina, indice, molde):
    with _trava_previa:
        _recortes_rapidos.pop(chave_ajuste(foto, pagina, indice, molde), None)


def buraco_reduzido(buraco, escala):
    x, y, w, h = buraco
    return (
        round(x * escala), round(y * escala),
        max(1, round(w * escala)), max(1, round(h * escala)),
    )


def montar_previa_pagina(pagina, colocadas):
    """
    Página reduzida com as fotos nos buracos.
    colocadas: lista de (buraco em pixels do molde, foto já cortada ou None).
    """
    base, escala = pagina_reduzida(pagina)
    fotos = []
    for buraco, foto in colocadas:
        if foto is None:
            continue
        x, y, w, h = buraco_reduzido(buraco, escala)
        if foto.size != (w, h):
            foto = foto.resize((w, h), Image.BILINEAR)
        fotos.append((x, y, foto))
    return compor_pagina(base, fotos)


# ============================================================
# POSES
# ============================================================

def _abrir_draft(caminho, lado):
    """Foto decodificada direto em escala reduzida (JPEG), com pelo menos lado px no menor lado."""
    img = Image.open(caminho)
    try:
        img.draft("RGB", (lado, lado))
    except Exception:
        pass
    img = ImageOps.exif_transpose(img)
    if img.mode != "RGB":
        img = img.convert("RGB")
    return img


def _reduzir(img, lado):
    escala = lado / max(img.size)
    return img.resize(
        (max(1, round(img.width * escala)), max(1, round(img.height * escala))),
        Image.BILINEAR,
    )


def abrir_miniatura(caminho, lado=480):
    """Lê uma foto de cada vez, já reduzida. A pose não precisa do arquivo inteiro."""
    img = _abrir_draft(caminho, lado)
    if max(img.size) > lado:
        img = _reduzir(img, lado)
    return img


def miniatura_leve(caminho, lado=180):
    """Miniatura guardada em disco (fora da pasta das fotos) para abrir de novo na hora."""
    caminho = Path(caminho)
    destino = None
    try:
        st = caminho.stat()
        chave = "%s|%s|%s|%s" % (
            os.path.normcase(os.path.abspath(str(caminho))), st.st_mtime_ns, st.st_size, lado,
        )
        destino = PASTA_MINIATURAS / (hashlib.sha1(chave.encode("utf-8")).hexdigest() + ".jpg")
        if destino.is_file():
            with Image.open(destino) as guardada:
                return guardada.convert("RGB")
    except OSError:
        pass
    img = abrir_miniatura(caminho, lado)
    if destino is not None:
        try:
            PASTA_MINIATURAS.mkdir(parents=True, exist_ok=True)
            temp = destino.with_name("%s.%s.tmp" % (destino.stem, threading.get_ident()))
            img.save(temp, "JPEG", quality=85)
            os.replace(temp, destino)
        except OSError:
            pass
    return img


def _vetor_pose(img):
    """Assinatura do centro da foto, para juntar os cliques da mesma pose."""
    w, h = img.size
    recorte = img.crop((int(w * 0.2), int(h * 0.15), int(w * 0.8), int(h * 0.85)))
    recorte = recorte.resize((16, 16), Image.BILINEAR)
    valores = np.asarray(recorte, dtype=np.float32).reshape(-1)
    valores -= valores.mean()
    norma = np.linalg.norm(valores)
    return valores / (norma or 1.0)


def _vetor_corpo(img, rosto):
    """
    Miniatura colorida do corpo, alinhada e na escala do rosto: mostra onde
    estão braços e mãos, para achar a mesma pose em fotos fora de sequência.
    """
    if not rosto:
        return None
    lab = cv2.cvtColor(np.asarray(img.convert("RGB")), cv2.COLOR_RGB2LAB).astype(np.float32)
    h = rosto["h"]
    x0, y0 = rosto["cx"] - 3 * h, rosto["y"] + 0.9 * h
    ex, ey = 16 / (6.0 * h), 16 / (4.5 * h)
    transformacao = np.float32([[ex, 0, -x0 * ex], [0, ey, -y0 * ey]])
    valores = cv2.warpAffine(
        lab, transformacao, (16, 16), flags=cv2.INTER_AREA, borderMode=cv2.BORDER_REPLICATE
    ).reshape(-1, 3)
    valores = (valores - valores.mean(axis=0)).reshape(-1)
    norma = np.linalg.norm(valores)
    return valores / norma if norma else None


def _nitidez(img):
    pequena = img.convert("L")
    escala = min(1.0, 480 / max(pequena.size))
    if escala < 1.0:
        pequena = pequena.resize(
            (max(1, round(pequena.width * escala)), max(1, round(pequena.height * escala))),
            Image.BILINEAR,
        )
    return float(cv2.Laplacian(np.asarray(pequena), cv2.CV_64F).var())


def orientacao(w, h):
    proporcao = w / float(h)
    if abs(proporcao - 1.0) <= TOLERANCIA_QUADRADO:
        return "quadrado"
    return "horizontal" if w > h else "vertical"


def aceita(orient_buraco, orient_foto):
    """Buraco quadrado aceita qualquer foto; os outros, só a mesma orientação."""
    return orient_buraco == "quadrado" or orient_buraco == orient_foto


# ============================================================
# FICHA DE IDENTIFICAÇÃO, CENÁRIO E OLHAR
# ============================================================

_orb = None
_ficha = None


def _ficha_referencia():
    """(pontos, descritores) da ficha impressa, ou None sem o arquivo."""
    global _orb, _ficha
    if _orb is None:
        _orb = cv2.ORB_create(2000)
    if _ficha is None:
        if not FICHA_REFERENCIA.is_file():
            _ficha = False
        else:
            with Image.open(FICHA_REFERENCIA) as ref:
                cinza = np.asarray(ref.convert("L"))
            pontos, descritores = _orb.detectAndCompute(cinza, None)
            _ficha = (pontos, descritores) if descritores is not None else False
    return _ficha or None


def _pontos_ficha(img):
    """Quantos pontos da ficha impressa aparecem na foto, após a homografia."""
    ref = _ficha_referencia()
    if ref is None:
        return None
    pontos_ref, desc_ref = ref
    cinza = np.asarray(img.convert("L"))
    pontos, desc = _orb.detectAndCompute(cinza, None)
    if desc is None or len(pontos) < 10:
        return 0
    pares = cv2.BFMatcher(cv2.NORM_HAMMING).knnMatch(desc_ref, desc, k=2)
    bons = [p[0] for p in pares if len(p) == 2 and p[0].distance < 0.75 * p[1].distance]
    if len(bons) < 8:
        return 0
    origem = np.float32([pontos_ref[m.queryIdx].pt for m in bons])
    destino = np.float32([pontos[m.trainIdx].pt for m in bons])
    _h, mascara = cv2.findHomography(origem, destino, cv2.RANSAC, 5.0)
    return int(mascara.sum()) if mascara is not None else 0


def _papel_branco(img, rosto):
    """Sem a ficha de referência: retângulo branco grande abaixo do rosto."""
    hsv = cv2.cvtColor(np.asarray(img.convert("RGB")), cv2.COLOR_RGB2HSV)
    altura, largura = hsv.shape[:2]
    branco = ((hsv[..., 1] < 50) & (hsv[..., 2] > 170)).astype(np.uint8)
    n, _labels, stats, _ = cv2.connectedComponentsWithStats(branco, connectivity=8)
    abaixo = (rosto["y"] + rosto["h"]) if rosto else altura * 0.4
    for i in range(1, n):
        x, y, w, h, area = (int(v) for v in stats[i])
        if area < 0.04 * largura * altura or y < abaixo:
            continue
        if area / float(w * h) >= 0.8:
            return True
    return False


def tem_ficha(img, rosto=None):
    pontos = _pontos_ficha(img)
    if pontos is None:
        return _papel_branco(img, rosto)
    return pontos >= MIN_PONTOS_FICHA


def _vetor_cenario(img):
    """Cor mediana (Lab) das bordas de cima e dos lados, onde quase só aparece o fundo."""
    lab = cv2.cvtColor(np.asarray(img.convert("RGB")), cv2.COLOR_RGB2LAB)
    altura, largura = lab.shape[:2]
    mascara = np.zeros((altura, largura), bool)
    mascara[: int(altura * 0.12), :] = True
    mascara[: int(altura * 0.7), : int(largura * 0.08)] = True
    mascara[: int(altura * 0.7), largura - int(largura * 0.08):] = True
    return np.median(lab[mascara].astype(np.float32), axis=0)


def agrupar_cenarios(vetores):
    """
    Índice de cenário para cada vetor, por ligação média: junta os grupos
    mais parecidos até a distância média passar de LIMIAR_CENARIO.
    O cenário 0 é o da primeira foto.
    """
    n = len(vetores)
    if not n:
        return []
    pontos = np.asarray(vetores, dtype=np.float64)
    dist = np.linalg.norm(pontos[:, None, :] - pontos[None, :, :], axis=2)
    np.fill_diagonal(dist, np.inf)
    tamanhos = np.ones(n)
    ativo = np.ones(n, bool)
    rotulo = np.arange(n)
    while ativo.sum() > 1:
        i, j = np.unravel_index(np.argmin(dist), dist.shape)
        if dist[i, j] > LIMIAR_CENARIO:
            break
        novo = (tamanhos[i] * dist[i] + tamanhos[j] * dist[j]) / (tamanhos[i] + tamanhos[j])
        dist[i, :] = novo
        dist[:, i] = novo
        dist[i, i] = np.inf
        dist[j, :] = np.inf
        dist[:, j] = np.inf
        tamanhos[i] += tamanhos[j]
        ativo[j] = False
        rotulo[rotulo == j] = i
    ordem = {}
    for r in rotulo:
        ordem.setdefault(int(r), len(ordem))
    return [ordem[int(r)] for r in rotulo]


def _vetor_roupa(img, rosto):
    """
    Cor do tronco, logo abaixo do queixo: mediana Lab e histograma grosso de
    matiz e saturação. Sem rosto, usa o terço central de baixo da foto.
    """
    largura, altura = img.size
    if rosto:
        cx, base, fw, fh = rosto["cx"], rosto["y"] + rosto["h"], rosto["w"], rosto["h"]
        caixa = (cx - fw, base + fh * 0.15, cx + fw, base + fh * 2.65)
    else:
        caixa = (largura / 3.0, altura * 2 / 3.0, largura * 2 / 3.0, altura)
    l, t, r, b = (
        int(max(0, caixa[0])), int(max(0, caixa[1])),
        int(min(largura, caixa[2])), int(min(altura, caixa[3])),
    )
    if r - l < 8 or b - t < 8:
        l, t, r, b = int(largura / 3.0), int(altura * 2 / 3.0), int(largura * 2 / 3.0), altura
    rgb = np.asarray(img.convert("RGB").crop((l, t, r, b)))
    lab = cv2.cvtColor(rgb, cv2.COLOR_RGB2LAB).reshape(-1, 3).astype(np.float32)
    hsv = cv2.cvtColor(rgb, cv2.COLOR_RGB2HSV)
    hist = cv2.calcHist([hsv], [0, 1], None, [8, 3], [0, 180, 0, 256]).reshape(-1).astype(np.float32)
    return {"lab": np.median(lab, axis=0), "hist": hist / (hist.sum() or 1.0)}


def distancia_roupa(a, b):
    """Diferença de cor Lab (em unidades de 50) mais a distância dos histogramas."""
    return (
        float(np.linalg.norm(a["lab"] - b["lab"])) / 50.0
        + float(cv2.compareHist(a["hist"], b["hist"], cv2.HISTCMP_BHATTACHARYYA))
    )


def _media_roupa(vetores):
    return {
        "lab": np.mean([v["lab"] for v in vetores], axis=0).astype(np.float32),
        "hist": np.mean([v["hist"] for v in vetores], axis=0).astype(np.float32),
    }


def _numero_arquivo(caminho):
    numeros = re.findall(r"\d+", Path(caminho).stem)
    return int(numeros[-1]) if numeros else None


def _sequencias(fotos):
    """Trechos de numeração contínua (saltos de até SALTO_NUMERACAO), na ordem dos arquivos."""
    trechos = []
    for foto in sorted(fotos, key=chave_natural):
        if trechos:
            antes, agora = _numero_arquivo(trechos[-1][-1]), _numero_arquivo(foto)
            if antes is None or agora is None or agora - antes <= SALTO_NUMERACAO:
                trechos[-1].append(foto)
                continue
        trechos.append([foto])
    return trechos


def suavizar_cenarios(fotos, info):
    """
    Fotos tiradas em sequência estão no mesmo fundo: cada trecho de numeração
    contínua fica com o cenário da maioria. Depois renumera os cenários pela
    ordem em que aparecem.
    """
    if any(_numero_arquivo(f) is None for f in fotos):
        return
    for trecho in _sequencias(fotos):
        contagem = {}
        for foto in trecho:
            contagem[info[foto]["cenario"]] = contagem.get(info[foto]["cenario"], 0) + 1
        maioria = max(contagem, key=lambda c: (contagem[c], -min(
            i for i, f in enumerate(trecho) if info[f]["cenario"] == c
        )))
        for foto in trecho:
            info[foto]["cenario"] = maioria
    novos = {}
    for foto in sorted(fotos, key=chave_natural):
        novos.setdefault(info[foto]["cenario"], len(novos))
    for foto in fotos:
        info[foto]["cenario"] = novos[info[foto]["cenario"]]


def agrupar_roupas(fotos, info):
    """
    Dentro de cada trecho de numeração contínua, começa roupa nova só quando
    CONFIRMA_ROUPA fotos seguidas fogem da média da roupa atual (uma foto com
    o braço na frente não quebra a roupa). Na dúvida divide: roupa dividida
    em dois trechos não mistura página, roupas juntas misturam. Grava
    info[foto]["roupa"], numerada por cenário e depois pela ordem da sessão.
    """
    vetor = lambda f: info[f]["roupa_vetor"]
    trechos = []
    for sequencia in _sequencias(fotos):
        atuais = [[sequencia[0]]]
        for i in range(1, len(sequencia)):
            proximas = sequencia[i:i + CONFIRMA_ROUPA]
            if len(atuais[-1]) >= CONFIRMA_ROUPA and len(proximas) == CONFIRMA_ROUPA:
                media = _media_roupa([vetor(f) for f in atuais[-1]])
                if all(distancia_roupa(media, vetor(f)) > LIMIAR_ROUPA for f in proximas):
                    atuais.append([sequencia[i]])
                    continue
            atuais[-1].append(sequencia[i])
        trechos.extend(atuais)
    trechos.sort(key=lambda t: (
        Counter(info[f]["cenario"] for f in t).most_common(1)[0][0], chave_natural(t[0])
    ))
    for roupa, trecho in enumerate(trechos):
        for foto in trecho:
            info[foto]["roupa"] = roupa


def esta_olhando(rosto):
    """Cabeça de frente: nariz no meio dos olhos e olhos quase nivelados."""
    if not rosto or "nariz" not in rosto:
        return False
    (dx, dy), (ex, ey), (nx, _ny) = rosto["olho_d"], rosto["olho_e"], rosto["nariz"]
    distancia = float(np.hypot(ex - dx, ey - dy))
    if distancia <= 0 or abs(ey - dy) > INCLINACAO_MAXIMA * abs(ex - dx):
        return False
    return abs(nx - (dx + ex) / 2) / distancia < LIMIAR_OLHANDO


def _brilho_pupila(cinza):
    """
    Reflexo do flash na íris: ponto claro e pequeno cercado de escuro,
    em proporção à pele. Olho fechado não tem íris à mostra, então não tem reflexo.
    """
    h, w = cinza.shape
    ref = max(1.0, float(np.percentile(cinza, 75)))
    escuro = (cinza < 0.4 * ref).astype(np.float32)
    disco = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (13, 13)).astype(np.float32)
    vizinhanca = cv2.filter2D(escuro, -1, disco / disco.sum())
    topo = cv2.morphologyEx(cinza, cv2.MORPH_TOPHAT, cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (9, 9)))
    topo = topo * (vizinhanca > 0.5)
    return float(topo[:, int(w * 0.15): int(w * 0.85)].max() / ref)


def _abertura(cinza):
    """Altura da parte escura (íris e cílios) no meio do olho, em fração do recorte."""
    h, w = cinza.shape
    ref = float(np.percentile(cinza, 75))
    suave = cv2.GaussianBlur(cinza, (3, 3), 0)
    escuro = (suave < 0.4 * ref).astype(np.uint8)
    escuro = cv2.morphologyEx(escuro, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (1, 5)))
    colunas = escuro[:, int(w * 0.25): int(w * 0.75)].sum(axis=0)
    return float(np.percentile(colunas, 90)) / h


def medir_olhos(img, rosto, escala=1.0):
    """
    Recorta os dois olhos (pelos pontos do YuNet) e mede o reflexo na íris,
    a abertura e quanto o olho está escuro em relação à pele do rosto
    (óculos escuros). escala leva as coordenadas do rosto para img.
    Sem os pontos dos olhos devolve None.
    """
    if not rosto or "olho_d" not in rosto:
        return None
    (dx, dy), (ex, ey), (nx, ny) = rosto["olho_d"], rosto["olho_e"], rosto["nariz"]
    distancia = float(np.hypot(ex - dx, ey - dy)) * escala
    if distancia < 8:
        return None
    cinza = img.convert("L")
    meia = distancia * 0.3
    pele = np.asarray(cinza.crop((
        int(nx * escala - meia), int(ny * escala - meia * 0.5),
        int(nx * escala + meia), int(ny * escala + meia * 0.5),
    )), dtype=np.float32)
    pele = float(np.median(pele)) if pele.size else 0.0
    brilhos, aberturas, escuros = [], [], []
    for x, y in ((dx, dy), (ex, ey)):
        cx, cy = x * escala, y * escala
        caixa = (
            int(round(cx - distancia * 0.32)), int(round(cy - distancia * 0.22)),
            int(round(cx + distancia * 0.32)), int(round(cy + distancia * 0.22)),
        )
        olho = np.asarray(cinza.crop(caixa).resize(TAMANHO_OLHO, Image.BICUBIC), dtype=np.float32)
        brilhos.append(_brilho_pupila(olho))
        aberturas.append(_abertura(olho))
        escuros.append(float((olho < 0.3 * pele).mean()) if pele > 0 else 0.0)
    return {
        "brilho": max(brilhos),
        "abertura": max(aberturas),
        "oculto": min(escuros) >= LIMIAR_OCULOS,
    }


def marcar_olhos_fechados(info):
    """
    Olho fechado: pouco reflexo na íris em relação às outras fotos da pasta
    (a luz do estúdio é a mesma) e abertura pequena. Os dois sinais juntos,
    para não tirar foto boa só porque a luz não fez reflexo. Com os olhos
    escondidos (óculos escuros) não dá para saber: fica como aberto.
    """
    medidas = [d["olhos"] for d in info.values() if d.get("olhos") and not d["olhos"]["oculto"]]
    referencia = float(np.median([m["brilho"] for m in medidas])) if medidas else 0.0
    for dados in info.values():
        olhos = dados.get("olhos")
        dados["olhos_ocultos"] = bool(olhos and olhos["oculto"])
        dados["olho_fechado"] = bool(
            olhos and not olhos["oculto"] and referencia > 0
            and olhos["brilho"] < LIMIAR_BRILHO_OLHO * referencia
            and olhos["abertura"] < ABERTURA_MINIMA
        )


def analisar_fotos(fotos, ao_foto=None):
    """
    Abre cada foto uma vez e devolve {foto: dados} com orientação, pose,
    nitidez, cenário, olhando, olho fechado e identificação.
    """
    info = {}
    total = len(fotos)
    for indice, caminho in enumerate(fotos, start=1):
        if ao_foto:
            ao_foto(indice, total, caminho.name)
        grande = _abrir_draft(caminho, LADO_ANALISE)
        img = _reduzir(grande, LADO_ANALISE) if max(grande.size) > LADO_ANALISE else grande
        try:
            rosto = detectar_rosto(img)
            info[caminho] = {
                "orientacao": orientacao(*img.size),
                "pose": _vetor_pose(img),
                "corpo": _vetor_corpo(img, rosto),
                "nitidez": _nitidez(img),
                "cenario_vetor": _vetor_cenario(img),
                "roupa_vetor": _vetor_roupa(img, rosto),
                "olhando": esta_olhando(rosto),
                "olhos": medir_olhos(grande, rosto, grande.width / float(img.width)),
                "identificacao": tem_ficha(img, rosto),
            }
        finally:
            img.close()
            grande.close()
    marcar_olhos_fechados(info)
    validas = [f for f in fotos if not info[f]["identificacao"] and not info[f]["olho_fechado"]]
    for foto, cenario in zip(validas, agrupar_cenarios([info[f]["cenario_vetor"] for f in validas])):
        info[foto]["cenario"] = cenario
    suavizar_cenarios(validas, info)
    agrupar_roupas(validas, info)
    return info


def agrupar_poses(fotos, info):
    """
    Fotos seguidas e parecidas são a mesma pose.
    Em cada grupo vêm primeiro as que olham para a câmera, depois as de olhos
    à mostra, depois as mais nítidas.
    """
    grupos = []
    atual = []
    anterior = None
    for caminho in fotos:
        vetor = info[caminho]["pose"]
        if anterior is not None and float(anterior @ vetor) < LIMIAR_POSE:
            grupos.append(atual)
            atual = []
        atual.append(caminho)
        anterior = vetor
    if atual:
        grupos.append(atual)
    return [
        sorted(grupo, key=lambda f: (not info[f]["olhando"], info[f]["olhos_ocultos"], -info[f]["nitidez"]))
        for grupo in grupos
    ]


def _parecida(foto, fotos_na_pagina, info):
    """A foto repete a pose de alguma que já está na página."""
    vetor, corpo = info[foto]["pose"], info[foto].get("corpo")
    for f in fotos_na_pagina:
        if float(vetor @ info[f]["pose"]) >= LIMIAR_POSE:
            return True
        outro = info[f].get("corpo")
        if corpo is not None and outro is not None and float(corpo @ outro) >= LIMIAR_CORPO:
            return True
    return False


def _melhor(candidatas, info, poses_na_pagina, fotos_na_pagina=()):
    """
    Primeiro evita pose repetida na página (mesmo grupo ou foto parecida com
    uma que já está lá); depois olhando ganha de não olhando; depois a ordem
    das rodadas.
    """
    return min(
        candidatas,
        key=lambda r: (
            r[0] in poses_na_pagina,
            _parecida(r[1], fotos_na_pagina, info),
            not info[r[1]]["olhando"],
        ),
        default=None,
    )


def _tentar_pagina(buracos, rodadas, info, usos, cenarios):
    """
    Preenche os buracos só com fotos novas dos cenários dados, na ordem de preferência.
    Em página grande, não enche buraco com pose já usada se ainda há pose nova.
    Não altera usos. Devolve a lista de escolhas, com None onde não coube.
    """
    tipos = [orientacao(b[2], b[3]) for b in buracos]
    ordem = sorted(range(len(buracos)), key=lambda i: tipos[i] == "quadrado")
    da_pagina = [None] * len(buracos)
    poses_na_pagina = set()
    fotos_na_pagina = set()
    grande = len(buracos) >= PAGINA_GRANDE
    for i in ordem:
        for posicao, cenario in enumerate(cenarios):
            candidatas = [
                r for r in rodadas
                if r[1] not in usos and r[1] not in fotos_na_pagina
                and info[r[1]]["cenario"] == cenario
                and aceita(tipos[i], info[r[1]]["orientacao"])
            ]
            if grande and candidatas:
                novas = [
                    r for r in candidatas
                    if r[0] not in poses_na_pagina and not _parecida(r[1], fotos_na_pagina, info)
                ]
                if novas:
                    candidatas = novas
                elif poses_na_pagina:
                    # Prefere deixar o buraco vazio a repetir pose na página grande.
                    continue
            escolha = _melhor(candidatas, info, poses_na_pagina, fotos_na_pagina)
            if escolha:
                indice, foto, rodada = escolha
                da_pagina[i] = (indice, foto, rodada, False, posicao > 0)
                poses_na_pagina.add(indice)
                fotos_na_pagina.add(foto)
                break
    return da_pagina


def _completar_com_repeticao(buracos, da_pagina, rodadas, info, usos, atual):
    """
    Buracos vazios: repete foto do cenário atual; senão foto nova de outro
    cenário; senão repete de qualquer cenário. Por último, aceita a mesma
    foto outra vez na página — buraco vazio não pode sobrar.
    """
    tipos = [orientacao(b[2], b[3]) for b in buracos]
    poses_na_pagina = {e[0] for e in da_pagina if e}
    fotos_na_pagina = {e[1] for e in da_pagina if e}
    for i, escolha in enumerate(da_pagina):
        if escolha:
            continue
        servem = [
            r for r in rodadas
            if r[1] not in fotos_na_pagina and aceita(tipos[i], info[r[1]]["orientacao"])
        ]
        passos = [
            ([r for r in servem if r[1] in usos and info[r[1]]["cenario"] == atual], True, False),
            ([r for r in servem if r[1] not in usos], False, True),
            ([r for r in servem if r[1] in usos], True, True),
            # Último recurso: mesma foto já na página (melhor que buraco vazio).
            ([r for r in rodadas if aceita(tipos[i], info[r[1]]["orientacao"])], True, False),
        ]
        for candidatas, repetida, misturado in passos:
            if not candidatas:
                continue
            novas = [
                r for r in candidatas
                if r[0] not in poses_na_pagina and not _parecida(r[1], fotos_na_pagina, info)
            ]
            candidatas = novas or candidatas
            menos = min(usos.get(r[1], 0) for r in candidatas)
            indice, foto, rodada = _melhor(
                [r for r in candidatas if usos.get(r[1], 0) == menos], info, poses_na_pagina, fotos_na_pagina
            )
            da_pagina[i] = (indice, foto, rodada, repetida, misturado and info[foto]["cenario"] != atual)
            poses_na_pagina.add(indice)
            fotos_na_pagina.add(foto)
            break
        if da_pagina[i] is None and rodadas:
            # A roupa não tem foto nessa orientação: entra cortada.
            candidatas = [r for r in rodadas if r[1] not in fotos_na_pagina] or list(rodadas)
            menos = min(usos.get(r[1], 0) for r in candidatas)
            indice, foto, rodada = _melhor(
                [r for r in candidatas if usos.get(r[1], 0) == menos], info, poses_na_pagina, fotos_na_pagina
            )
            da_pagina[i] = (indice, foto, rodada, foto in usos or foto in fotos_na_pagina, False)
            poses_na_pagina.add(indice)
            fotos_na_pagina.add(foto)
    return da_pagina


def _rodadas(grupos, indices):
    """Uma foto de cada pose antes de repetir alguma: (pose, foto, rodada)."""
    rodadas = []
    maior = max((len(grupos[i]) for i in indices), default=0)
    for rodada in range(maior):
        for i in indices:
            if rodada < len(grupos[i]):
                rodadas.append((i, grupos[i][rodada], rodada))
    return rodadas


def _completa(da_pagina):
    return None not in da_pagina


def _variedade(rodadas, info):
    """Quantas fotos da roupa cabem numa página sem repetir pose."""
    poses, fotos = set(), []
    for pose, foto, _rodada in rodadas:
        if pose not in poses and not _parecida(foto, fotos, info):
            poses.add(pose)
            fotos.append(foto)
    return len(fotos)


def _poses_repetidas(da_pagina, info):
    """Fotos da página que repetem a pose (ou o corpo) de outra anterior."""
    poses, fotos, repetidas = set(), [], 0
    for e in da_pagina:
        if not e:
            continue
        if e[0] in poses or _parecida(e[1], fotos, info):
            repetidas += 1
        poses.add(e[0])
        fotos.append(e[1])
    return repetidas


def _custo_poses(da_pagina, info):
    """
    Pose repetida pesa mais em página com vários buracos: uma pose em 4–6
    buracos deixa as outras poses de fora e estraga a página.
    """
    repetidas = _poses_repetidas(da_pagina, info)
    buracos = sum(1 for e in da_pagina if e)
    if buracos >= PAGINA_GRANDE and repetidas:
        return repetidas * buracos
    return repetidas


def _fotos_duplicadas(da_pagina):
    """Quantas vezes a mesma foto aparece de novo na página."""
    fotos = [e[1] for e in da_pagina if e]
    return len(fotos) - len(set(fotos))


def _minimos_roupa(roupas, por_roupa, variedade, n_paginas):
    """
    Mínimo de páginas por roupa: 1 sempre. Sobe para 2 nas roupas do fim da
    sessão que têm fotos de sobra, mas deixa pelo menos metade das páginas
    extras para as cotas proporcionais (não equaliza nem engole o álbum).
    """
    minimo = {r: 1 for r in roupas}
    resto = n_paginas - len(roupas)
    if resto <= 0:
        return minimo
    extras = max(1, resto // 2) if resto >= 2 else resto
    candidatos = [
        r for r in roupas
        if len(por_roupa[r]) >= 4 and variedade[r] >= 1
    ]
    # Mais fotos primeiro; no empate, roupa mais tarde na sessão.
    candidatos.sort(key=lambda r: (-len(por_roupa[r]), -roupas.index(r)))
    for r in candidatos:
        if extras <= 0:
            break
        minimo[r] = 2
        extras -= 1
    return minimo


def _custo_trecho(moldes, rodadas, cota, minimo_roupa, paginas_buracos, info):
    """
    Preenche as páginas `moldes` só com a roupa (maiores primeiro, sem
    repetir foto entre elas enquanto der). Devolve (custo, escolhas por
    molde, usos). O custo soma, com pesos: buraco vazio, mesma foto duas vezes
    na página, pose repetida, foto repetida de outra página, distância da cota
    de buracos, foto que não olha e roupa abaixo do mínimo de páginas.
    """
    cenario = info[rodadas[0][1]]["cenario"]
    usos = {}
    escolhas = {}
    for p in sorted(moldes, key=lambda p: -len(paginas_buracos[p])):
        buracos = paginas_buracos[p]
        if not buracos:
            escolhas[p] = []
            continue
        da_pagina = _tentar_pagina(buracos, rodadas, info, usos, [cenario])
        if not _completa(da_pagina):
            da_pagina = _completar_com_repeticao(buracos, da_pagina, rodadas, info, usos, cenario)
        for escolha in da_pagina:
            if escolha:
                usos[escolha[1]] = usos.get(escolha[1], 0) + 1
        escolhas[p] = da_pagina

    paginas = buracos = faltas = duplicadas = poses = repetidas = nao_olha = cortadas = 0
    for p, da_pagina in escolhas.items():
        if not paginas_buracos[p]:
            continue
        paginas += 1
        buracos += len(da_pagina)
        faltas += sum(1 for e in da_pagina if e is None)
        cortadas += sum(
            1 for b, e in zip(paginas_buracos[p], da_pagina)
            if e and not aceita(orientacao(b[2], b[3]), info[e[1]]["orientacao"])
        )
        duplicadas += _fotos_duplicadas(da_pagina)
        poses += _custo_poses(da_pagina, info)
        repetidas += sum(1 for e in da_pagina if e and e[3])
        nao_olha += sum(1 for e in da_pagina if e and not info[e[1]]["olhando"])
    custo = (
        PESO_FALTA * faltas
        + PESO_FOTO_DUPLICADA * duplicadas
        + PESO_ORIENTACAO * cortadas
        + PESO_POSE * poses
        + PESO_REPETIDA * repetidas
        + abs(buracos - cota)
        + PESO_NAO_OLHA * nao_olha
    )
    if paginas < minimo_roupa:
        custo += PESO_ABAIXO_MINIMO
    return custo, escolhas, usos


def _indices_fixos(n):
    """Índices (0-based) das páginas que não saem do lugar; a última sempre."""
    fixas = {p - 1 for p in PAGINAS_FIXAS if 1 <= p <= n}
    if n:
        fixas.add(n - 1)
    return fixas


def _respeita_ordem(ordem, n):
    """Fixas no lugar e nenhuma página a mais de DISTANCIA_MAXIMA da posição do molde."""
    fixas = _indices_fixos(n)
    return sorted(ordem) == list(range(n)) and all(
        abs(p - posicao) <= DISTANCIA_MAXIMA and (p not in fixas or p == posicao)
        for posicao, p in enumerate(ordem)
    )


def _planejar_album(roupas, cotas, minimo, paginas_buracos, por_roupa, info):
    """
    Divide o álbum em trechos contíguos, um por roupa e na ordem dada. As
    páginas ficam na ordem do molde; só na fronteira entre duas roupas a
    última página de um trecho pode trocar com a primeira do seguinte (nunca
    uma página fixa, e cada página troca no máximo uma vez). Programação
    dinâmica sobre as posições. Devolve (custo, [(molde, roupa, escolhas)]
    na ordem do álbum, usos).
    """
    n = len(paginas_buracos)
    total = len(roupas)
    fixas = _indices_fixos(n)
    pode_trocar = [False] * (n + 1)
    for b in range(1, n):
        pode_trocar[b] = (b - 1) not in fixas and b not in fixas
    # Deixar uma roupa de fora custa menos que um buraco vazio.
    pular = True
    maior_trecho = n if total <= 1 else max(3, min(n, 2 * n // total + 4))
    infinito = float("inf")

    cache = {}

    def trecho(moldes, k):
        chave = (tuple(sorted(moldes)), k)
        if chave not in cache:
            r = roupas[k]
            cache[chave] = _custo_trecho(
                moldes, por_roupa[r], cotas[r], minimo[r], paginas_buracos, info
            )
        return cache[chave]

    def moldes_do_trecho(i, j, troca_esq, troca_dir):
        moldes = list(range(i, j))
        if troca_esq:
            moldes[0] = i - 1
        if troca_dir:
            moldes[-1] = j
        return moldes

    memo = {}

    def melhor(i, k, troca_esq):
        if i == n:
            return PESO_ROUPA_AUSENTE * (total - k), None
        if k == total:
            return infinito, None
        chave = (i, k, troca_esq)
        if chave in memo:
            return memo[chave]
        resultado = (infinito, None)
        if not pular and total - k > n - i:
            memo[chave] = resultado
            return resultado
        if pular:
            custo = PESO_ROUPA_AUSENTE + melhor(i, k + 1, troca_esq)[0]
            if custo < resultado[0]:
                resultado = (custo, "pula")
        ultima = k == total - 1
        fim = n if ultima else min(n, i + maior_trecho)
        for j in range(n if ultima else i + 1, fim + 1):
            if not pular and total - (k + 1) > n - j:
                break
            for troca_dir in (False, True):
                if troca_dir and (j == n or not pode_trocar[j] or (troca_esq and j - i == 1)):
                    continue
                custo = trecho(moldes_do_trecho(i, j, troca_esq, troca_dir), k)[0]
                custo += melhor(j, k + 1, troca_dir)[0] + PESO_TROCA * troca_dir
                if custo < resultado[0]:
                    resultado = (custo, (j, troca_dir))
        memo[chave] = resultado
        return resultado

    custo_total = melhor(0, 0, False)[0]
    album_paginas = []
    usos = {}
    i, k, troca_esq = 0, 0, False
    while i < n and k < total:
        decisao = melhor(i, k, troca_esq)[1]
        if decisao is None:
            break
        if decisao == "pula":
            k += 1
            continue
        j, troca_dir = decisao
        moldes = moldes_do_trecho(i, j, troca_esq, troca_dir)
        _custo, escolhas, usos_trecho = trecho(moldes, k)
        for p in moldes:
            album_paginas.append((p, roupas[k], escolhas[p]))
        for foto, vezes in usos_trecho.items():
            usos[foto] = usos.get(foto, 0) + vezes
        i, k, troca_esq = j, k + 1, troca_dir
    return custo_total, album_paginas, usos


def distribuir_roupas(grupos, info, paginas_buracos):
    """
    Uma roupa por página, sem misturar. As roupas seguem a ordem da sessão
    dentro de cada cenário, e os cenários vêm um de cada vez, cada roupa num
    trecho seguido de páginas. As páginas PAGINAS_FIXAS ficam no lugar; as
    outras só trocam com a vizinha (no máximo DISTANCIA_MAXIMA do molde).
    Cada roupa ganha pelo menos uma página (e duas quando couber), e o resto
    segue a cota proporcional às fotos, sem deixar cenário/roupa de fora.
    Devolve (lista de (índice do molde, roupa, escolhas) na ordem do álbum, sobras).
    Cada escolha é (pose, foto, rodada, repetida, misturado) ou None.
    """
    indices_roupa = {}
    for i, grupo in enumerate(grupos):
        indices_roupa.setdefault(info[grupo[0]]["roupa"], []).append(i)
    por_roupa = {r: _rodadas(grupos, idx) for r, idx in indices_roupa.items()}
    por_roupa = {r: rod for r, rod in por_roupa.items() if rod}
    if not por_roupa:
        vazias = [(p, None, [None] * len(b)) for p, b in enumerate(paginas_buracos)]
        return vazias, []

    total_fotos = sum(len(rod) for rod in por_roupa.values())
    total_buracos = sum(len(b) for b in paginas_buracos)
    cotas = {r: total_buracos * len(rod) / float(total_fotos) for r, rod in por_roupa.items()}
    cenario_de = {r: info[rod[0][1]]["cenario"] for r, rod in por_roupa.items()}
    fotos_cenario = {}
    for r, rod in por_roupa.items():
        fotos_cenario[cenario_de[r]] = fotos_cenario.get(cenario_de[r], 0) + len(rod)
    variedade = {r: _variedade(rod, info) for r, rod in por_roupa.items()}
    n_com_buraco = sum(1 for b in paginas_buracos if b)
    minimo = _minimos_roupa(list(por_roupa), por_roupa, variedade, n_com_buraco)
    cenarios = sorted(fotos_cenario, key=lambda c: -fotos_cenario[c])
    ordens = list(itertools.permutations(cenarios)) if len(cenarios) <= 5 else [tuple(cenarios)]

    melhor = None
    for ordem_cenarios in ordens:
        roupas = sorted(por_roupa, key=lambda r: (ordem_cenarios.index(cenario_de[r]), r))
        custo, album_paginas, usos = _planejar_album(
            roupas, cotas, minimo, paginas_buracos, por_roupa, info
        )
        if not _respeita_ordem([p for p, _r, _e in album_paginas], len(paginas_buracos)):
            continue
        if melhor is None or custo < melhor[0]:
            melhor = (custo, album_paginas, usos)
    if melhor is None:
        raise RuntimeError("não foi possível montar o álbum com as páginas fixas")
    _custo, album_paginas, usos = melhor
    sobras = [foto for grupo in grupos for foto in grupo if foto not in usos]
    return album_paginas, sobras


# ============================================================
# LISTA E MONTAGEM
# ============================================================

def montar_lista(pasta_fotos, pasta_molde=None, ao_foto=None):
    """
    Uma linha por buraco do molde gravado.
    A foto é a melhor de cada pose, sem repetir a pose na mesma página.
    """
    paginas = carregar_molde_fixo()
    fotos = listar_imagens(pasta_fotos)
    info = analisar_fotos(fotos, ao_foto=ao_foto)
    identificacoes = [f for f in fotos if info[f]["identificacao"]]
    fechados = [f for f in fotos if not info[f]["identificacao"] and info[f]["olho_fechado"]]
    validas = [f for f in fotos if not info[f]["identificacao"] and not info[f]["olho_fechado"]]
    grupos = []
    for roupa in sorted({info[f]["roupa"] for f in validas}):
        grupos.extend(agrupar_poses(
            sorted((f for f in validas if info[f]["roupa"] == roupa), key=chave_natural), info
        ))
    album_paginas, sobras = distribuir_roupas(grupos, info, [pagina["buracos"] for pagina in paginas])
    itens = []
    for ordem, (p, _roupa, da_pagina) in enumerate(album_paginas, start=1):
        pagina = paginas[p]
        buracos = pagina["buracos"]
        if not buracos:
            itens.append({
                "tipo": "vazia",
                "pagina": pagina["arquivo"],
                "ordem": ordem,
                "indice": 0,
                "buraco": None,
                "foto": None,
                "pose": None,
                "roupa": None,
                "situacao": "sem buraco",
            })
            continue
        for indice, (buraco, escolha) in enumerate(zip(buracos, da_pagina), start=1):
            cenario = None
            roupa = None
            if escolha:
                pose, foto, rodada, repetida, misturado = escolha
                cenario = info[foto]["cenario"]
                roupa = info[foto]["roupa"]
                partes = [
                    "cenário %s" % (cenario + 1), "roupa %s" % (roupa + 1), "pose %s" % (pose + 1),
                ]
                if repetida:
                    partes.append("repetida")
                elif rodada:
                    partes.append("de novo")
                if misturado:
                    partes.append("cenário misturado")
                if not aceita(orientacao(buraco[2], buraco[3]), info[foto]["orientacao"]):
                    partes.append("orientação trocada")
                if not info[foto]["olhando"]:
                    partes.append("não olha")
                if info[foto]["olhos_ocultos"]:
                    partes.append("olhos não visíveis")
                situacao = ", ".join(partes)
            else:
                foto = None
                pose = None
                situacao = "falta foto"
            itens.append({
                "tipo": "buraco",
                "pagina": pagina["arquivo"],
                "ordem": ordem,
                "indice": indice,
                "buraco": buraco,
                "foto": foto,
                "pose": pose,
                "cenario": cenario,
                "roupa": roupa,
                "situacao": situacao,
            })
    for foto in sobras:
        itens.append({
            "tipo": "sobra",
            "pagina": None,
            "indice": 0,
            "buraco": None,
            "foto": foto,
            "pose": None,
            "cenario": info[foto]["cenario"],
            "roupa": info[foto]["roupa"],
            "situacao": "mesma pose ou orientação sem buraco, não entra",
        })
    for foto in identificacoes:
        itens.append({
            "tipo": "identificacao",
            "pagina": None,
            "indice": 0,
            "buraco": None,
            "foto": foto,
            "pose": None,
            "cenario": None,
            "situacao": "identificação, não entra",
        })
    for foto in fechados:
        itens.append({
            "tipo": "olho_fechado",
            "pagina": None,
            "indice": 0,
            "buraco": None,
            "foto": foto,
            "pose": None,
            "cenario": None,
            "situacao": "olho fechado, não entra",
        })
    return itens


FORA_DO_ALBUM = ("sobra", "identificacao", "olho_fechado")


def _situacao_trocada(item, situacao_antiga, tipo_antigo):
    partes = ["trocada manualmente"]
    if tipo_antigo == "sobra":
        partes.append("veio de fora")
    elif tipo_antigo == "identificacao":
        partes.append("era identificação")
    elif tipo_antigo == "olho_fechado":
        partes.append("olho fechado")
    if item.get("cenario") is not None:
        partes.append("cenário %s" % (item["cenario"] + 1))
    if item.get("roupa") is not None:
        partes.append("roupa %s" % (item["roupa"] + 1))
    if item.get("pose") is not None:
        partes.append("pose %s" % (item["pose"] + 1))
    for marca in ("veio de fora", "era identificação", "olho fechado", "não olha", "olhos não visíveis"):
        if marca in situacao_antiga and marca not in partes:
            partes.append(marca)
    return ", ".join(partes)


def trocar_fotos(a, b):
    """
    Troca a foto entre dois itens, de buraco com buraco ou com uma foto de fora.
    A foto que sai do álbum vira sobra. O corte manual do novo lugar é apagado,
    para o rosto ser procurado de novo.
    """
    if not a["foto"] and not b["foto"]:
        raise ValueError("nenhuma das duas tem foto")
    if a["tipo"] != "buraco" and b["tipo"] != "buraco":
        raise ValueError("pelo menos um dos dois precisa ser buraco")
    antes = {}
    for item in (a, b):
        antes[id(item)] = (item["tipo"], item["situacao"])
        if item["tipo"] == "buraco" and item["foto"]:
            molde = (item["buraco"][2], item["buraco"][3])
            esquecer_recorte_rapido(item["foto"], item["pagina"], item["indice"], molde)
    roupa_diferente = (
        a.get("roupa") is not None and b.get("roupa") is not None and a["roupa"] != b["roupa"]
    )
    for chave in ("foto", "pose", "cenario", "roupa"):
        a[chave], b[chave] = b.get(chave), a.get(chave)
    for item, outro in ((a, b), (b, a)):
        tipo_antigo, situacao_antiga = antes[id(outro)]
        if item["tipo"] == "buraco":
            if not item["foto"]:
                item["situacao"] = "falta foto"
                continue
            molde = (item["buraco"][2], item["buraco"][3])
            apagar_manual(item["foto"], item["pagina"], item["indice"], molde)
            esquecer_recorte_rapido(item["foto"], item["pagina"], item["indice"], molde)
            item["situacao"] = _situacao_trocada(item, situacao_antiga, tipo_antigo)
            if roupa_diferente:
                item["situacao"] += ", roupa misturada"
        else:
            item["tipo"] = "sobra"
            item["situacao"] = "trocada manualmente, saiu do álbum, não entra"


def salvar_jpg(img, destino, icc, dpi):
    destino.parent.mkdir(parents=True, exist_ok=True)
    opcoes = {"quality": QUALIDADE_JPG, "subsampling": 0, "optimize": True}
    if icc:
        opcoes["icc_profile"] = icc
    if dpi:
        opcoes["dpi"] = dpi
    temp = destino.with_name(destino.stem + ".album_tmp.jpg")
    img.save(temp, "JPEG", **opcoes)
    os.replace(temp, destino)


def _grupos(itens):
    grupos = []
    for item in itens:
        if item["tipo"] in ("sobra", "identificacao", "olho_fechado"):
            continue
        if not grupos or grupos[-1][0] != item["pagina"]:
            grupos.append((item["pagina"], []))
        grupos[-1][1].append(item)
    return grupos


def nome_saida(pagina, ordem=None):
    """Nome do JPG da página: com a posição no álbum na frente, para imprimir na ordem."""
    if ordem is None:
        return pagina.stem + ".jpg"
    return "%02d - %s.jpg" % (ordem, pagina.stem)


def gerar(itens, pasta_saida, aviso=print, parar=lambda: False, ao_item=None):
    """Grava uma JPG por página do molde. Devolve False se parou no meio."""
    pasta_saida = Path(pasta_saida)
    pasta_saida.mkdir(parents=True, exist_ok=True)
    grupos = _grupos(itens)
    feitas = 0
    for pagina, grupo in grupos:
        if parar():
            aviso("Parado.")
            return False
        aviso(pagina.name)
        base, icc, dpi = abrir_molde(pagina)
        colocadas = []
        try:
            for item in grupo:
                if parar():
                    aviso("Parado.")
                    return False
                if item["tipo"] != "buraco" or not item["foto"] or not item["buraco"]:
                    if item["situacao"] == "falta foto":
                        aviso("  buraco %s: falta foto" % item["indice"])
                    continue
                x, y, w, h = item["buraco"]
                try:
                    foto_img, _ficc, _fdpi = abrir_imagem(item["foto"])
                except OSError as e:
                    item["situacao"] = "erro ao abrir"
                    aviso("  buraco %s: erro ao abrir %s (%s)" % (item["indice"], item["foto"].name, e))
                    if ao_item:
                        ao_item(item)
                    continue
                try:
                    caixa, origem = caixa_da_foto(
                        item["foto"], pagina, item["indice"], (w, h), foto_img
                    )
                    colocadas.append((x, y, recortar_imagem(foto_img, caixa, (w, h))))
                    item["situacao"] = origem
                    aviso("  buraco %s: %s (%s)" % (item["indice"], item["foto"].name, origem))
                finally:
                    foto_img.close()
                if ao_item:
                    ao_item(item)
            destino = pasta_saida / nome_saida(pagina, grupo[0].get("ordem"))
            pronta = compor_pagina(base, colocadas)
            try:
                salvar_jpg(pronta, destino, icc, dpi)
            finally:
                pronta.close()
            feitas += 1
            aviso("  gravou " + destino.name)
        finally:
            base.close()
    sobras = sum(1 for i in itens if i["tipo"] == "sobra")
    aviso("Pronto: %s página(s) em %s" % (feitas, pasta_saida))
    if sobras:
        aviso("%s foto(s) sobraram e não entraram." % sobras)
    return True
