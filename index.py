import fitz  # PyMuPDF
from PIL import Image, ImageChops
import customtkinter as ctk
from tkinter import filedialog, messagebox, simpledialog
import re
from collections import Counter, defaultdict
import os
import subprocess
import platform
import time
import threading

# Importar sistema de auto-update
try:
    from updater import check_and_update
    from version import get_version
    UPDATE_AVAILABLE = True
except ImportError:
    UPDATE_AVAILABLE = False

# Importar tela de carregamento
try:
    from loading_screen import show_loading_screen
    LOADING_AVAILABLE = True
except ImportError:
    LOADING_AVAILABLE = False

# Importar sistema de rastreamento de SKUs
try:
    from sku_tracker import atualizar_planilha_skus, coletar_skus_de_ocorrencias, coletar_skus_de_registros_ml, get_tracker_status, carregar_planilha_em_dict
    SKU_TRACKER_AVAILABLE = True
except ImportError:
    SKU_TRACKER_AVAILABLE = False

# =========================
# Verificação de dependências e configuração de DPI
# =========================


def verificar_dependencias():
    """Verifica se todas as dependências necessárias estão disponíveis"""
    try:
        import fitz
        import PIL
        import customtkinter
        return True
    except ImportError as e:
        print(f"Erro de dependência: {e}")
        return False


def configurar_escalonamento_dpi():
    """Configura o escalonamento para tamanhos consistentes"""
    try:
        # Desabilitar escalonamento automático do CustomTkinter
        ctk.set_widget_scaling(1.0)
        ctk.set_window_scaling(1.0)
        return 1.0
    except Exception as e:
        print(f"Erro ao configurar DPI: {e}")
        ctk.set_widget_scaling(1.0)
        ctk.set_window_scaling(1.0)
        return 1.0


def configurar_dpi_awareness():
    """Configura awareness de DPI para Windows - VERSÃO SEGURA"""
    # Completamente desabilitado para evitar comandos problemáticos
    pass


def configurar_ambiente_consistente():
    """Configura o ambiente para tamanhos consistentes"""
    try:
        # Configurar CustomTkinter para não usar escalonamento automático
        ctk.set_widget_scaling(1.0)
        ctk.set_window_scaling(1.0)
    except Exception:
        pass


# =========================
# Configurações visuais / constantes
# =========================
DESLOC_INDICE_X = 5
DESLOC_INDICE_Y = 2
TAMANHO_INDICE = 10
AUMENTO_TEXTO_FACTOR = 1.5
HIGHLIGHT_OFFSET = 12
HIGHLIGHT_MARGIN_LEFT = 10  # Margem mínima para linha bem mais larga
HIGHLIGHT_MARGIN_RIGHT = 10  # Margem mínima para linha bem mais larga

INDICE_POSICAO_FIXA_X = 400
INDICE_OFFSET_Y = 25

SKU_PATTERN = re.compile(r"^([A-Z0-9\/\-_\.\s]+)\s*-\s*(.*)", re.IGNORECASE)
FRACAO_PATTERN = re.compile(r"\b(\d{1,2})/(\d{1,2})\b")


def normalizar_sku(sku):
    """Normaliza SKU para garantir consistência, especialmente para códigos com barra"""
    if not sku:
        return ""
    # Converter para maiúsculas primeiro
    sku_normalizado = sku.upper().strip()
    # Normalizar espaços ao redor da barra: "18200 / SM501001", "18200/ SM501001", "18200 /SM501001" -> "18200/SM501001"
    # Isso garante que todos os formatos sejam tratados como o mesmo SKU
    sku_normalizado = re.sub(r'\s*/\s*', '/', sku_normalizado)
    # Normalizar múltiplas barras consecutivas
    sku_normalizado = re.sub(r'/+', '/', sku_normalizado)
    # Remover espaços extras em outras partes (mas manter o código limpo)
    sku_normalizado = re.sub(r'\s+', ' ', sku_normalizado)
    return sku_normalizado.strip()

# =========================
# Funções auxiliares (copiadas do arquivo original)
# =========================

# =========================
# Funções do modo ML (integradas do pedidos.py)
# =========================


# Configurações para modo ML
FONTE_ML = "helv"      # fonte interna do PyMuPDF (Helvetica)
TAMANHO_ML = 8
DESLOC_X_ML = 5        # desloc. horizontal a partir do fim do código
DESLOC_Y_ML = 7        # desloc. vertical levemente abaixo da base

# Tokens que podem compor o "código" após SKU (números, letras, /, -, _)
REGEX_TOKEN_ML = re.compile(r"^[A-Z0-9]+$", re.IGNORECASE)
TOKENS_LIVRES_ML = {"/", "-", "_"}


def is_code_token_ml(tok: str) -> bool:
    return bool(REGEX_TOKEN_ML.match(tok)) or tok in TOKENS_LIVRES_ML


def extrair_sku_numerico(codigo_completo: str) -> str:
    """
    Extrai apenas a parte numérica do SKU, ignorando descrições.

    Exemplos:
    - "80465 - KIT RELACAO..." -> "80465"
    - "18200/SM501001 - Descrição" -> "18200/SM501001"
    - "KIT123-456 - Algo" -> "KIT123-456"

    Retorna apenas a parte antes do primeiro "-" seguido de espaço,
    ou se não houver, retorna tudo até encontrar uma palavra descritiva.
    """
    if not codigo_completo:
        return ""

    codigo_str = codigo_completo.strip()

    # Se houver " - " (hífen com espaços), pegar apenas a parte antes
    if " - " in codigo_str:
        return codigo_str.split(" - ", 1)[0].strip()

    # Se houver espaço seguido de letras maiúsculas múltiplas (descrição), cortar aí
    # Mas primeiro, verificar se é um código com múltiplas partes (como "123 / ABC456")
    partes = codigo_str.split()

    if len(partes) == 1:
        # Sem espaços, retorna como está
        return codigo_str

    # Retornar apenas a primeira parte ( assumindo que é o SKU)
    # Exceto se for um padrão tipo "123 / ABC456" (código com barra)
    sku_parte = partes[0]
    if len(partes) > 1 and partes[1] in ["/", "-"]:
        # Código tipo "123 / 456" ou "123 - 456"
        if len(partes) > 2:
            sku_parte = f"{partes[0]} {partes[1]} {partes[2]}"
        else:
            sku_parte = f"{partes[0]} {partes[1]}"
        # Verificar se há descrição após
        resto = " ".join(partes[3:]) if len(partes) > 3 else ""
        if resto and " - " in resto:
            return sku_parte.strip()

    return sku_parte.strip()


def coletar_ocorrencias_por_pagina_ml(page):
    """Coleta ocorrências de SKU no formato ML (SKU: XXXXX)"""
    words = page.get_text("words")
    if not words:
        return []

    words.sort(key=lambda w: (w[5], w[6], w[7]))
    ocorrencias = []
    n = len(words)
    i = 0

    def get(idx): return words[idx][4].strip() if 0 <= idx < n else ""

    def bbox(
        idx): return words[idx][0], words[idx][1], words[idx][2], words[idx][3]
    def same_line(a, b): return (
        words[a][5], words[a][6]) == (words[b][5], words[b][6])

    while i < n:
        t = get(i).upper()
        start_code_idx = None
        if t == "SKU:":
            if i + 1 < n and same_line(i, i + 1):
                start_code_idx = i + 1
        elif t == "SKU":
            if i + 2 < n and get(i + 1) == ":" and same_line(i, i + 2):
                start_code_idx = i + 2
            elif i + 1 < n and same_line(i, i + 1):
                start_code_idx = i + 1

        if start_code_idx is not None:
            j = start_code_idx
            code_tokens = []
            while j < n and same_line(i, j) and is_code_token_ml(get(j)):
                code_tokens.append(get(j))
                j += 1

            if code_tokens:
                codigo_str = " ".join(code_tokens).upper()
                # Extrair apenas a parte numérica do SKU (ignorar descrições)
                sku_numerico = extrair_sku_numerico(codigo_str)
                x0, y0, x1, y1 = bbox(j - 1)
                ocorrencias.append((sku_numerico, (x0, y0, x1, y1)))
                i = j
                continue
        i += 1
    return ocorrencias


def inserir_indice_ml(page, bbox_final_token, texto_indice):
    """Insere índice no formato ML"""
    x0, y0, x1, y1 = bbox_final_token
    pos_x = x1 + DESLOC_X_ML
    pos_y = y0 + DESLOC_Y_ML
    page.insert_text((pos_x, pos_y), texto_indice,
                     fontsize=TAMANHO_ML, fontname=FONTE_ML, fill=(0, 0, 0))


def desenhar_highlight_ml(page, sku_bbox, quantidade=1):
    """Desenha uma linha grande de destaque abaixo do bloco (modo ML)"""
    x0, y0, x1, y1 = sku_bbox
    largura_pagina = page.rect.width
    y_linha = y1 + 12  # um pouco abaixo do SKU/quantidade

    # Ajustar espessura da linha baseado na quantidade
    # Se quantidade > 2, linha bem mais grossa e visível
    if quantidade > 2:
        line_width = 5.0  # Linha bem mais grossa e visível
    elif quantidade > 1:
        line_width = 3.0  # Linha média
    else:
        line_width = 1.5  # Linha padrão

    page.draw_line(
        (10, y_linha),              # margem mínima para linha bem mais larga
        # margem mínima para linha bem mais larga
        (largura_pagina - 10, y_linha),
        color=(0, 0, 0),
        width=line_width
    )


def detectar_tipo_folha(doc):
    """Detecta automaticamente se é folha Bling ou ML"""
    try:
        # Verificar primeira página
        primeira_pagina = doc[0]
        texto = primeira_pagina.get_text().upper()

        # Padrões que indicam DANFE (Bling)
        padroes_danfe = ["DANFE", "NOTA FISCAL",
                         "VL. ITEM", "QTD. TOTAL DE ITENS"]
        # Padrões que indicam ML
        padroes_ml = ["SKU:", "SKU "]

        score_danfe = sum(1 for padrao in padroes_danfe if padrao in texto)
        score_ml = sum(1 for padrao in padroes_ml if padrao in texto)

        # Se encontrar mais padrões ML, é ML
        if score_ml > score_danfe:
            return "ML"
        # Se encontrar padrões DANFE, é Bling
        elif score_danfe > 0:
            return "Bling"
        # Se não conseguir detectar, verificar se há SKUs no formato ML
        else:
            ocorrencias = coletar_ocorrencias_por_pagina_ml(primeira_pagina)
            if ocorrencias:
                return "ML"
            else:
                return "Bling"  # Padrão por defeito
    except Exception:
        return "Bling"  # Padrão por defeito em caso de erro


def processar_pdf_ml(caminho_pdf, modo="global"):
    """Processa PDF no formato ML (integradas do pedidos.py)"""
    base = os.path.splitext(caminho_pdf)[0]
    pdf_saida = base + "_Indice.pdf"
    csv_saida = base + "_Indice.csv"

    doc = fitz.open(caminho_pdf)
    registros = []

    todas_ocorrencias = []
    ocorr_por_pagina = []
    for pidx in range(len(doc)):
        page = doc[pidx]
        occ = coletar_ocorrencias_por_pagina_ml(page)
        ocorr_por_pagina.append(occ)
        for codigo, bb in occ:
            todas_ocorrencias.append((pidx, codigo, bb))

    if modo == "global":
        totais = Counter([codigo for _, codigo, _ in todas_ocorrencias])
        usados = defaultdict(int)
    else:
        totais_por_pagina = []
        for occ in ocorr_por_pagina:
            totais_por_pagina.append(Counter([c for c, _ in occ]))

    for pidx in range(len(doc)):
        page = doc[pidx]
        occ = ocorr_por_pagina[pidx]

        if modo == "global":
            for codigo, bb in occ:
                usados[codigo] += 1
                idx_atual = usados[codigo]
                idx_total = totais[codigo]
                texto = f"[{idx_atual}/{idx_total}]"
                inserir_indice_ml(page, bb, texto)
                registros.append(
                    [codigo, f"{idx_atual}/{idx_total}", pidx + 1])

                # Procurar "Quantidade:" logo abaixo do SKU
                blocks = page.get_text("blocks")
                for blk in blocks:
                    bx0, by0, bx1, by1, btext, *_ = blk
                    if "Quantidade:" in btext and abs(bx0 - bb[0]) < 50 and by0 > bb[1]:
                        try:
                            qtd = int(btext.split(":")[1].strip())
                            if qtd > 1:
                                desenhar_highlight_ml(page, bb, qtd)
                        except:
                            pass
                        break

        else:  # por página
            totais_local = totais_por_pagina[pidx]
            usados_local = defaultdict(int)
            for codigo, bb in occ:
                usados_local[codigo] += 1
                idx_atual = usados_local[codigo]
                idx_total = totais_local[codigo]
                texto = f"[{idx_atual}/{idx_total}]"
                inserir_indice_ml(page, bb, texto)
                registros.append(
                    [codigo, f"{idx_atual}/{idx_total}", pidx + 1])

                blocks = page.get_text("blocks")
                for blk in blocks:
                    bx0, by0, bx1, by1, btext, *_ = blk
                    if "Quantidade:" in btext and abs(bx0 - bb[0]) < 50 and by0 > bb[1]:
                        try:
                            qtd = int(btext.split(":")[1].strip())
                            if qtd > 1:
                                desenhar_highlight_ml(page, bb, qtd)
                        except:
                            pass
                        break

    # Adicionar marcadores de página antes de salvar
    adicionar_marcadores_pagina(doc)

    doc.save(pdf_saida)
    doc.close()

    # Salvar CSV apenas se houver registros
    if registros:
        import csv
        with open(csv_saida, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["SKU (normalizado)", "INDICE", "PAGINA"])
            w.writerows(registros)

    # Atualizar planilha de SKUs (rastreamento automático)
    skus_updated = False
    try:
        if SKU_TRACKER_AVAILABLE and registros:
            skus_quantidades = coletar_skus_de_registros_ml(registros)
            if skus_quantidades:
                try:
                    skus_updated = atualizar_planilha_skus(skus_quantidades)
                except Exception as e:
                    print(
                        f"⚠️  Aviso: Erro ao atualizar planilha de SKUs: {e}")
                    skus_updated = False
    except Exception as e:
        print(f"⚠️  Aviso: Erro ao coletar/atualizar SKUs (ML): {e}")
        skus_updated = False

    # Garantir retorno consistente: (pdf_saida, csv_saida, skus_updated)
    if registros:
        return pdf_saida, csv_saida, skus_updated
    else:
        return pdf_saida, None, False


def detectar_posicao_ideal_indice(pagina):
    try:
        data = pagina.get_text("dict")
        for bloco in data.get("blocks", []):
            if "lines" in bloco:
                for linha in bloco["lines"]:
                    texto = " ".join([s["text"]
                                     for s in linha["spans"]]).strip()
                    if "VL. ITEM" in texto.upper():
                        if linha["spans"]:
                            return linha["spans"][0]["bbox"][0] - 20
    except Exception:
        pass
    return INDICE_POSICAO_FIXA_X


def abrir_arquivo_para_impressao(caminho_arquivo):
    try:
        if not os.path.exists(caminho_arquivo):
            return False

        sistema = platform.system().lower()
        if sistema == "windows":
            try:
                os.startfile(caminho_arquivo)
                return True
            except Exception:
                try:
                    subprocess.run(["start", "", caminho_arquivo],
                                   shell=True, check=True)
                    return True
                except Exception:
                    return False
        elif sistema == "darwin":
            try:
                subprocess.run(
                    ["open", "-a", "Preview", caminho_arquivo], check=True)
                return True
            except Exception:
                try:
                    subprocess.run(["open", caminho_arquivo], check=True)
                    return True
                except Exception:
                    return False
        else:
            try:
                subprocess.run(["xdg-open", caminho_arquivo], check=True)
                return True
            except Exception:
                for cmd in ["evince", "okular", "firefox", "chrome"]:
                    try:
                        subprocess.run([cmd, caminho_arquivo], check=True)
                        return True
                    except Exception:
                        continue
                return False
    except Exception as e:
        print(f"Erro ao abrir arquivo: {e}")
        return False


def detectar_bbox(pagina):
    pix = pagina.get_pixmap(dpi=150)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    bg = Image.new("RGB", img.size, (255, 255, 255))
    diff = ImageChops.difference(img, bg)
    bbox = diff.getbbox()
    if not bbox:
        return pagina.rect
    x0, y0, x1, y1 = bbox
    fator_x = pagina.rect.width / pix.width if pix.width != 0 else 1
    fator_y = pagina.rect.height / pix.height if pix.height != 0 else 1
    return fitz.Rect(x0 * fator_x, y0 * fator_y, x1 * fator_x, y1 * fator_y)


def adicionar_texto_personalizado_pagina(pagina, texto_personalizado, posicao_y):
    if not texto_personalizado or not texto_personalizado.strip():
        return

    try:
        largura_pagina = pagina.rect.width
        pos_x = largura_pagina - 120
        pos_y = posicao_y + 50

        pagina.insert_text(
            (pos_x, pos_y),
            texto_personalizado.strip(),
            fontsize=15,
            fontname="helv",
            fill=(0, 0, 0)
        )
    except Exception as e:
        print(f"Erro ao adicionar texto personalizado: {e}")


def adicionar_marcadores_pagina(doc):
    """Adiciona marcadores de página no canto inferior direito de cada página do PDF em negrito"""
    try:
        total_paginas = len(doc)

        for num_pagina in range(total_paginas):
            pagina = doc[num_pagina]
            largura_pagina = pagina.rect.width
            altura_pagina = pagina.rect.height

            # Formato: "folha 1/x", "2/x", etc.
            texto_marcador = f"Pagina {num_pagina + 1}/{total_paginas}"

            # Posição no canto inferior direito
            # Margem de 100 pixels da direita e 20 pixels do fundo
            pos_x = largura_pagina - 70
            pos_y = altura_pagina - 20

            # Simular negrito usando duas inserções com pequeno deslocamento
            # Primeira inserção (base)
            pagina.insert_text(
                (pos_x, pos_y),
                texto_marcador,
                fontsize=10,
                fontname="helv",
                fill=(0, 0, 0)
            )
            # Segunda inserção ligeiramente deslocada para criar efeito de negrito
            pagina.insert_text(
                (pos_x + 0.4, pos_y),
                texto_marcador,
                fontsize=10,
                fontname="helv",
                fill=(0, 0, 0)
            )
    except Exception as e:
        print(f"Erro ao adicionar marcadores de página: {e}")


def destacar_textos_amarelos(pagina):
    try:
        data = pagina.get_text("dict")
        linhas = []
        for bloco in data.get("blocks", []):
            if "lines" in bloco:
                for linha in bloco["lines"]:
                    texto = " ".join([s["text"]
                                     for s in linha["spans"]]).strip()
                    spans = linha["spans"]
                    if texto:
                        linhas.append({"texto": texto, "spans": spans})

        spans_para_aumentar = []
        spans_com_info = []

        for idx, linha in enumerate(linhas):
            texto = linha["texto"]
            if "ITEM" in texto.upper():
                itens_encontrados = []
                for j in range(idx + 1, len(linhas)):
                    prox_texto = linhas[j]["texto"]
                    if "QTD. TOTAL DE ITENS" in prox_texto.upper() or "CONSUMIDOR" in prox_texto.upper():
                        break
                    if " - " in prox_texto and len(prox_texto) > 3:
                        itens_encontrados.append(linhas[j]["spans"])

                tem_multiplos = len(itens_encontrados) > 1

                for spans_item in itens_encontrados:
                    spans_para_aumentar.extend(spans_item)
                    for span in spans_item:
                        spans_com_info.append({
                            "span": span,
                            "tem_multiplos_itens": tem_multiplos
                        })
            elif "QTD. TOTAL DE ITENS" in texto:
                for j in range(idx + 1, min(idx + 4, len(linhas))):
                    prox_texto = linhas[j]["texto"]
                    if prox_texto.strip().isdigit():
                        spans_para_aumentar.extend(linhas[j]["spans"])
                        for span in linhas[j]["spans"]:
                            spans_com_info.append({
                                "span": span,
                                "tem_multiplos_itens": False
                            })
                        break

        for span_info in spans_com_info:
            span = span_info["span"]
            tem_multiplos_itens = span_info["tem_multiplos_itens"]

            texto_span = span.get("text", "").strip()
            if not texto_span:
                continue
            try:
                tamanho_original = span.get("size", 10)
                x0, y0, x1, y1 = span["bbox"]
                pad = 1
                rect = fitz.Rect(x0 - pad, y0 - pad, x1 + pad, y1 + pad)
                pagina.draw_rect(rect, color=(1, 1, 1), fill=(1, 1, 1))

                if tem_multiplos_itens:
                    fator_aumento = 1.2
                else:
                    fator_aumento = AUMENTO_TEXTO_FACTOR

                if " - " in texto_span:
                    partes = texto_span.split(" - ", 1)
                    sku = partes[0].strip()
                    nome_produto = partes[1].strip() if len(partes) > 1 else ""

                    pos_x_atual = x0

                    if sku:
                        novo_tamanho_sku = max(
                            8, int(tamanho_original * fator_aumento))
                        pos_y = y0 + novo_tamanho_sku * 0.8
                        pagina.insert_text((pos_x_atual, pos_y), sku,
                                           fontsize=novo_tamanho_sku, fontname="helv", fill=(0, 0, 0))

                        largura_sku_estimada = len(
                            sku) * novo_tamanho_sku * 0.6
                        pos_x_atual += largura_sku_estimada

                    if nome_produto:
                        traco = " - "
                        pos_y_traco = y0 + tamanho_original * 0.8
                        pagina.insert_text((pos_x_atual, pos_y_traco), traco,
                                           fontsize=tamanho_original, fontname="helv", fill=(0, 0, 0))

                        largura_traco_estimada = len(
                            traco) * tamanho_original * 0.6
                        pos_x_atual += largura_traco_estimada

                        pos_y_nome = y0 + tamanho_original * 0.8
                        pagina.insert_text((pos_x_atual, pos_y_nome), nome_produto,
                                           fontsize=tamanho_original, fontname="helv", fill=(0, 0, 0))

                        import re
                        padrao_quantidade = r'(\d+,\d+)\s+UN'
                        match_quantidade = re.search(
                            padrao_quantidade, nome_produto)

                        if match_quantidade:
                            quantidade_texto = match_quantidade.group(0)
                            quantidade_numero = match_quantidade.group(1)

                            try:
                                quantidade_float = float(
                                    quantidade_numero.replace(',', '.'))

                                if quantidade_float > 1:
                                    pos_quantidade = nome_produto.find(
                                        quantidade_texto)
                                    if pos_quantidade != -1:
                                        largura_antes_quantidade = len(
                                            nome_produto[:pos_quantidade]) * tamanho_original * 0.5
                                        pos_x_quantidade = pos_x_atual + largura_antes_quantidade

                                        numero_quantidade = quantidade_numero
                                        largura_numero = len(
                                            numero_quantidade) * tamanho_original * 5

                                        # Ajustar espessura da linha baseado na quantidade
                                        # Se quantidade > 2, linha bem mais grossa e visível
                                        if quantidade_float > 2:
                                            line_width = 6.0  # Linha bem mais grossa e visível
                                        else:
                                            line_width = 4.0  # Linha padrão para quantidade > 1

                                        y_highlight = y1 + 3
                                        pagina.draw_line(
                                            (pos_x_quantidade, y_highlight),
                                            (pos_x_quantidade +
                                             largura_numero, y_highlight),
                                            color=(0, 0, 0),
                                            width=line_width
                                        )
                            except ValueError:
                                pass
                else:
                    novo_tamanho = max(
                        8, int(tamanho_original * fator_aumento))
                    pos_y = y0 + novo_tamanho * 0.8
                    pagina.insert_text((x0, pos_y), texto_span,
                                       fontsize=novo_tamanho, fontname="helv", fill=(0, 0, 0))
            except Exception:
                continue
    except Exception as e:
        print("Erro em destacar_textos_amarelos:", e)


def escanear_documento_para_skus(doc):
    occurrences_per_page = {}
    totals = Counter()

    for pidx, page in enumerate(doc):
        data = page.get_text("dict")
        linhas = []
        for bloco in data.get("blocks", []):
            if "lines" in bloco:
                for linha in bloco["lines"]:
                    texto = " ".join([s["text"]
                                     for s in linha["spans"]]).strip()
                    spans = linha["spans"]
                    if texto:
                        linhas.append((texto, spans))

        occs = []
        for i, (texto, spans) in enumerate(linhas):
            if "ITEM" in texto.upper() and i + 1 < len(linhas):
                prox_texto, prox_spans = linhas[i + 1]
                prox_text = prox_texto.strip()
                if "|" in prox_text:
                    prox_text = prox_text.split("|")[0].strip()

                # Tentar capturar SKU usando o padrão
                m = SKU_PATTERN.match(prox_text)
                sku_bruto = None

                if m:
                    sku_bruto = m.group(1).strip()
                else:
                    # Se o padrão não funcionar, tentar capturar até o primeiro " - " ou " -"
                    # Isso ajuda com casos onde há códigos com barra mas o padrão não captura
                    if " - " in prox_text:
                        sku_bruto = prox_text.split(" - ", 1)[0].strip()
                    elif " -" in prox_text:
                        sku_bruto = prox_text.split(" -", 1)[0].strip()

                if sku_bruto:
                    # Normalizar SKU para garantir consistência com códigos que têm barra
                    sku_normalizado = normalizar_sku(sku_bruto)
                    # Extrair apenas a parte numérica do SKU (ignorar descrições)
                    sku = extrair_sku_numerico(sku_normalizado)

                    quantidade = 1
                    for j in range(i + 1, min(i + 6, len(linhas))):
                        t2, _ = linhas[j]
                        if "Quantidade:" in t2:
                            try:
                                quantidade = int(t2.split("Quantidade:")[
                                                 1].strip().split()[0])
                            except Exception:
                                quantidade = 1
                            break
                    x0s = [s["bbox"][0] for s in prox_spans]
                    y0s = [s["bbox"][1] for s in prox_spans]
                    x1s = [s["bbox"][2] for s in prox_spans]
                    y1s = [s["bbox"][3] for s in prox_spans]
                    line_bbox = (min(x0s), min(y0s), max(x1s), max(y1s))
                    occs.append({"sku": sku, "bbox": line_bbox,
                                "qtd": quantidade, "text": prox_text})
                    totals[sku] += 1

        occurrences_per_page[pidx] = occs

    return occurrences_per_page, totals


def ordenar_paginas_por_grupo(doc, occurrences_per_page):
    page_meta = {}
    num_pages = len(doc)
    for pidx in range(num_pages):
        page = doc[pidx]
        text = page.get_text("text") or ""
        m = FRACAO_PATTERN.search(text)
        skus = [occ["sku"] for occ in occurrences_per_page.get(pidx, [])]
        if m:
            num = int(m.group(1))
            total = int(m.group(2))
            sku_chave = skus[0] if skus else f"NOPSKU_{pidx}"
            group_key = (sku_chave, total)
            page_meta[pidx] = {"group": group_key, "num": num}
        else:
            page_meta[pidx] = {"group": None, "num": None}

    groups = defaultdict(list)
    for pidx, meta in page_meta.items():
        if meta["group"] is not None:
            groups[meta["group"]].append((meta["num"], pidx))

    groups_ordered = sorted(
        groups.items(), key=lambda kv: min([p for _, p in kv[1]]))

    ordered = []
    for key, lst in groups_ordered:
        lst_sorted = sorted(lst, key=lambda x: x[0])
        ordered.extend([pidx for _, pidx in lst_sorted])

    for pidx in range(num_pages):
        if pidx not in ordered:
            ordered.append(pidx)

    return ordered


def compactar_pdf_com_recorte(entrada, saida, etiquetas_por_pagina=3, destacar_textos=False, texto_personalizado=""):
    doc = fitz.open(entrada)
    novo = fitz.open()

    largura, altura = fitz.paper_size("a4")
    slot_h = altura / etiquetas_por_pagina
    contador = 0

    occurrences_per_page, totals = escanear_documento_para_skus(doc)
    usados_global = defaultdict(int)

    # Atualizar planilha de SKUs (Bling/DANFE) se disponível
    skus_atualizados = False
    skus_encontrados = {}
    try:
        if SKU_TRACKER_AVAILABLE:
            skus_encontrados = coletar_skus_de_ocorrencias(
                occurrences_per_page)
            if skus_encontrados:
                try:
                    skus_atualizados = atualizar_planilha_skus(
                        skus_encontrados)
                except Exception as e:
                    print(
                        f"⚠️ Erro ao atualizar planilha de SKUs (Bling): {e}")
                    skus_atualizados = False
    except Exception as e:
        print(f"⚠️ Erro ao coletar/atualizar SKUs: {e}")
        skus_atualizados = False

    ordem_paginas = ordenar_paginas_por_grupo(doc, occurrences_per_page)

    for orig_index in ordem_paginas:
        pagina = doc[orig_index]
        if contador % etiquetas_por_pagina == 0:
            pagina_nova = novo.new_page(width=largura, height=altura)
            if texto_personalizado and texto_personalizado.strip():
                adicionar_texto_personalizado_pagina(
                    pagina_nova, texto_personalizado, 0)

        linha = contador % etiquetas_por_pagina
        destino = fitz.Rect(5, linha * slot_h, largura -
                            5, (linha + 1) * slot_h)

        clip = detectar_bbox(pagina)

        tmp = fitz.open()
        tmp_page = tmp.new_page(width=pagina.rect.width,
                                height=pagina.rect.height)
        tmp_page.show_pdf_page(tmp_page.rect, doc, orig_index)

        if destacar_textos:
            try:
                destacar_textos_amarelos(tmp_page)
            except Exception:
                pass

        pos_x_ideal = detectar_posicao_ideal_indice(tmp_page)

        occs = occurrences_per_page.get(orig_index, [])
        for occ in occs:
            sku = occ["sku"]
            x0, y0, x1, y1 = occ["bbox"]
            usados_global[sku] += 1
            idx_atual = usados_global[sku]
            idx_total = totals.get(sku, 1)
            texto_indice = f"[{idx_atual}/{idx_total}]"

            pos_x_fixa = pos_x_ideal
            pos_y_relativa = y0 + INDICE_OFFSET_Y

            try:
                tmp_page.insert_text((pos_x_fixa, pos_y_relativa),
                                     texto_indice,
                                     fontsize=TAMANHO_INDICE,
                                     fontname="helv",
                                     fill=(0, 0, 0))
            except Exception:
                try:
                    tmp_page.insert_text((pos_x_fixa, y0 + 4),
                                         texto_indice,
                                         fontsize=TAMANHO_INDICE,
                                         fontname="helv",
                                         fill=(0, 0, 0))
                except Exception:
                    try:
                        tmp_page.insert_text((x1 + DESLOC_INDICE_X, y0 + DESLOC_INDICE_Y),
                                             texto_indice,
                                             fontsize=TAMANHO_INDICE,
                                             fontname="helv",
                                             fill=(0, 0, 0))
                    except Exception:
                        pass

            qtd_produto = occ.get("qtd", 1)
            if qtd_produto > 1:
                try:
                    largura_pagina = tmp_page.rect.width
                    y_linha = y1 + HIGHLIGHT_OFFSET
                    x_inicio = HIGHLIGHT_MARGIN_LEFT
                    x_fim = largura_pagina - HIGHLIGHT_MARGIN_RIGHT

                    # Ajustar espessura da linha baseado na quantidade
                    # Se quantidade > 2, linha bem mais grossa e visível
                    if qtd_produto > 2:
                        line_width = 5.0  # Linha bem mais grossa e visível
                    else:
                        line_width = 3.0  # Linha média para quantidade > 1

                    tmp_page.draw_line(
                        (x_inicio, y_linha), (x_fim, y_linha), color=(0, 0, 0), width=line_width)
                except Exception:
                    pass

        pagina_nova.show_pdf_page(destino, tmp, 0, clip=clip)
        tmp.close()
        contador += 1

    # Adicionar marcadores de página antes de salvar
    adicionar_marcadores_pagina(novo)

    novo.save(saida)
    novo.close()
    doc.close()
    # Retornar informações sobre atualização de SKUs
    return {"ok": True, "skus_updated": skus_atualizados, "skus": skus_encontrados}

# =========================
# Interface (CustomTkinter) - VERSÃO SEGURA
# =========================


class PDFApp(ctk.CTk):
    def __init__(self):
        # Configurar ambiente consistente ANTES de criar a janela
        configurar_ambiente_consistente()

        super().__init__()
        self.title("DANFE Compactador")

        # Definir tamanho fixo para consistência entre computadores (aumentado para acomodar switch)
        self.geometry("350x370")
        self.resizable(False, False)

        # Forçar tamanho mínimo e máximo para garantir consistência
        self.minsize(350, 375)  # altura x largura
        self.maxsize(350, 375)

        # Tentar carregar ícone, mas não falhar se não existir
        try:
            if os.path.exists('lc.ico'):
                self.iconbitmap('lc.ico')
        except Exception:
            pass

        # Configuração visual profissional com tratamento de erro
        try:
            ctk.set_appearance_mode("light")
            ctk.set_default_color_theme("blue")
        except Exception:
            pass

        # Centralizar janela
        self.center_window()

        self.arquivo_entrada = None
        self.destacar_textos = ctk.BooleanVar(value=True)
        self.texto_personalizado = ctk.StringVar(value="")
        self.modo_processamento = ctk.StringVar(
            value="auto")  # auto, bling, ml

        # Configurar grid responsivo
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # Header com título e ícone
        # self.create_header()

        # Main content area
        self.create_main_content()

        # Footer com informações
        self.create_footer()

        # Verificar atualizações em background (após 3 segundos)
        # REMOVIDO: self.after() que causava problemas
        if UPDATE_AVAILABLE:
            # Executar verificação de atualizações em thread separada
            def check_updates():
                time.sleep(2)  # Aguardar 3 segundos
                try:
                    if self.winfo_exists():
                        check_and_update(parent=self, silent=True)
                except Exception as e:
                    print(f"Erro na verificação de atualizações: {e}")

            threading.Thread(target=check_updates, daemon=True).start()

    # ---------------------------
    # Janela de Resumo de SKUs
    # ---------------------------
    def open_sku_panel(self):
        """Abre uma nova janela que substitui a aplicação principal mostrando o resumo de SKUs"""
        # Esconder janela principal

        try:
            self.withdraw()
            messagebox.showinfo(
                "Atenção", "App ainda em fase de desenvolvimento.\nQualquer erro ou problema, favor avisar.\n"
                "Intuito desta aba é acompanhar os SKUs que estão \nparados sem venda a longo prazo.")
        except Exception:
            pass

        # Criar toplevel seguro (CTkToplevel se disponível)
        try:
            top = ctk.CTkToplevel(self)
        except Exception:
            # Fallback para Toplevel puro
            from tkinter import Toplevel
            top = Toplevel(self)

        top.title("Resumo de SKUs")

        try:
            # Desabilitar redimensionamento ANTES de definir geometria (evita flicker)
            top.resizable(False, False)

            # Centralizar a janela do painel na tela
            w, h = 700, 520
            sw = top.winfo_screenwidth()
            sh = top.winfo_screenheight()
            x = max(0, (sw - w) // 2)
            y = max(0, (sh - h) // 2)
            top.geometry(f"{w}x{h}+{x}+{y}")
            top.update_idletasks()

        except Exception:
            pass

        def on_close():
            try:
                top.destroy()
            except Exception:
                pass
            try:
                self.deiconify()
            except Exception:
                pass

        top.protocol("WM_DELETE_WINDOW", on_close)

        # Header
        header = ctk.CTkLabel(top, text="Resumo da Planilha de SKUs",
                              font=ctk.CTkFont(size=16, weight="bold"))
        header.grid(row=0, column=0, columnspan=3, pady=(12, 8), padx=12)

        # Info / status
        status = get_tracker_status()
        path_label = ctk.CTkLabel(
            top, text=f"Backend: {status.get('backend')} - {status.get('path')}", font=ctk.CTkFont(size=10))
        path_label.grid(row=1, column=0, columnspan=3, sticky="w", padx=12)

        # Controls
        control_frame = ctk.CTkFrame(top, fg_color="transparent")
        control_frame.grid(row=2, column=0, columnspan=3,
                           sticky="ew", padx=12, pady=(8, 8))
        control_frame.grid_columnconfigure((0, 1, 2, 3, 4), weight=1)

        self.sku_period = ctk.StringVar(value="Dia")
        period_menu = ctk.CTkOptionMenu(control_frame, values=[
                                        "Dia", "Semana", "Mês"], variable=self.sku_period, width=120)
        period_menu.grid(row=0, column=0, padx=6)

        btn_top = ctk.CTkButton(control_frame, text="Top 20", command=lambda: self._refresh_sku_list(
            top, mode="top"), width=100)
        btn_top.grid(row=0, column=1, padx=6)

        btn_no_move = ctk.CTkButton(control_frame, text="Sem movimento",
                                    command=lambda: self._refresh_sku_list(top, mode="no_move"), width=120)
        btn_no_move.grid(row=0, column=2, padx=6)

        btn_open = ctk.CTkButton(
            control_frame, text="Abrir planilha", command=self._open_planilha, width=120)
        btn_open.grid(row=0, column=3, padx=6)

        btn_download = ctk.CTkButton(
            control_frame, text="Download bling+", command=self._download_bling_plus, width=140)
        btn_download.grid(row=0, column=4, padx=6)

        # Result area
        result_frame = ctk.CTkScrollableFrame(
            top, width=660, height=340, corner_radius=8)
        result_frame.grid(row=3, column=0, columnspan=3,
                          padx=12, pady=(6, 12), sticky="nsew")
        top.grid_rowconfigure(3, weight=1)
        top.grid_columnconfigure(0, weight=1)

        self.sku_result_frame = result_frame

        # Footer buttons
        footer_frame = ctk.CTkFrame(top, fg_color="transparent")
        footer_frame.grid(row=4, column=0, columnspan=3,
                          sticky="ew", padx=12, pady=(0, 12))
        footer_frame.grid_columnconfigure((0, 1), weight=1)

        btn_refresh = ctk.CTkButton(
            footer_frame, text="Atualizar", command=lambda: self._refresh_sku_list(top, mode="all"))
        btn_refresh.grid(row=0, column=0, sticky="w", padx=6)

        btn_back = ctk.CTkButton(footer_frame, text="Voltar", command=on_close)
        btn_back.grid(row=0, column=1, sticky="e", padx=6)

        # Inicializar lista (mostrar TODOS os itens por padrão)
        self._refresh_sku_list(top, mode="all")

    def _open_planilha(self):
        """Abre a planilha no sistema"""
        status = get_tracker_status()
        path = status.get("path")
        if not path or not os.path.exists(path):
            messagebox.showwarning(
                "Aviso", "Arquivo de controle não encontrado.")
            return
        # Usar função de abrir comum
        try:
            abrir_arquivo_para_impressao(path)
        except Exception as e:
            messagebox.showerror(
                "Erro", f"Não foi possível abrir o arquivo: {e}")

    def _download_bling_plus(self):
        """Abre o link direto do GitHub para o arquivo bling+.zip"""
        url = "https://github.com/Luukz-Y/danfe-updates/raw/refs/heads/main/bling+.zip"
        try:
            import webbrowser
            webbrowser.open(url)
            messagebox.showinfo(
                "Download", "Abrindo link de download automatico bling+\nDownload direto para sua pasta padrão do windows")
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir o link: {e}")

    def _refresh_sku_list(self, parent, mode="all"):
        """Atualiza a lista exibida no painel de SKUs"""
        # Limpar frame
        for w in self.sku_result_frame.winfo_children():
            w.destroy()

        data, path = carregar_planilha_em_dict()
        period = self.sku_period.get()
        key = "qtd_dia" if period == "Dia" else (
            "qtd_semana" if period == "Semana" else "qtd_mes")

        items = [(sku, vals.get(key, 0), vals.get("ultima", ""))
                 for sku, vals in data.items()]

        if mode == "top":
            items = sorted(items, key=lambda x: x[1], reverse=True)[:20]
        elif mode == "no_move":
            items = [it for it in items if it[1] == 0]
            items = sorted(items, key=lambda x: x[0])
        else:
            # Mostrar todos os itens ordenados por quantidade (descendente)
            items = sorted(items, key=lambda x: x[1], reverse=True)

        # Mostrar sumário (centralizado)
        total_skus = len(data)
        total_vendidos = sum(1 for _, v, _ in items if v >
                             0) if mode != "no_move" else len(items)
        summary = ctk.CTkLabel(
            self.sku_result_frame, text=f"Total SKUs: {total_skus}    Itens mostrados: {len(items)}", anchor="center")
        summary.pack(fill="x", pady=(6, 4), padx=8)

        # Mostrar itens
        for sku, qtd, ultima in items:
            # Mostrar apenas o código do SKU (parte antes de " - ")
            sku_display = str(sku)
            if " - " in sku_display:
                sku_display = sku_display.split(" - ", 1)[0].strip()
            # Se ainda houver espaços, mostrar só o primeiro token (ex: "80447 - ..." -> "80447")
            sku_display = sku_display.split(
            )[0] if sku_display.split() else sku_display
            lbl = ctk.CTkLabel(
                self.sku_result_frame, text=f"SKU: {sku_display} — Saida: {qtd} — Última atualização: {ultima}", anchor="w")
            lbl.pack(fill="x", padx=8, pady=2)

        if not items:
            lbl = ctk.CTkLabel(
                self.sku_result_frame, text="Nenhum item encontrado para o filtro.\nContinue usando o App para atualizar os dados da planilha.", anchor="w")
            lbl.pack(fill="x", padx=8, pady=8)

    def center_window(self):
        """Centraliza a janela na tela de forma robusta considerando DPI"""
        try:
            # Usar valores fixos para evitar problemas com update_idletasks
            width, height = 400, 450

            screen_width = self.winfo_screenwidth()
            screen_height = self.winfo_screenheight()

            x = (screen_width - width) // 2
            y = (screen_height - height) // 2

            x = max(0, min(x, screen_width - width))
            y = max(0, min(y, screen_height - height))

            self.geometry(f"{width}x{height}+{x}+{y}")

        except Exception:
            self.geometry("400x450")

    # def create_header(self):
        """Cria o cabeçalho da aplicação"""
        # header_frame = ctk.CTkFrame(self, height=60, corner_radius=0)
        # header_frame.grid(row=0, column=0, sticky="ew", padx=0, pady=0)
        # header_frame.grid_columnconfigure(0, weight=1)

        # title_label = ctk.CTkLabel(
        # header_frame,
        # text="Danfe Simplificado",
        # font=ctk.CTkFont(size=22, weight="bold"),
        # text_color=("#1f538d", "#14375e")
        # )
        # title_label.grid(row=0, column=0, pady=(10, 2))

        # subtitle_label = ctk.CTkLabel(
        # header_frame,
        # text="Bling (DANFE) • ML (SKUs) • Detecção Automática",
        # font=ctk.CTkFont(size=12),
        # text_color=("#666666", "#cccccc")
        # )
        # subtitle_label.grid(row=1, column=0, pady=(0, 8))

    def create_main_content(self):
        """Cria a área principal de conteúdo"""
        main_frame = ctk.CTkFrame(self, corner_radius=12)
        main_frame.grid(row=1, column=0, sticky="nsew", padx=15, pady=8)
        main_frame.grid_columnconfigure(0, weight=1)

        file_section = ctk.CTkFrame(main_frame, fg_color="transparent")
        file_section.grid(row=0, column=0, sticky="ew", padx=15, pady=12)
        file_section.grid_columnconfigure(0, weight=1)

        file_icon = ctk.CTkLabel(
            file_section,
            text="📄",
            font=ctk.CTkFont(size=32)
        )
        file_icon.grid(row=0, column=0, pady=(0, 5))

        section_title = ctk.CTkLabel(
            file_section,
            text="Selecionar Arquivo PDF",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        section_title.grid(row=1, column=0, pady=(0, 8))

        self.select_btn = ctk.CTkButton(
            file_section,
            text="📁 Escolher PDF de DANFE",
            command=self.selecionar_pdf,
            height=35,
            font=ctk.CTkFont(size=14, weight="bold"),
            corner_radius=20
        )
        self.select_btn.grid(row=2, column=0, pady=(0, 8))

        self.lbl_pdf = ctk.CTkLabel(
            file_section,
            text="Nenhum arquivo selecionado",
            font=ctk.CTkFont(size=11),
            text_color=("#666666", "#cccccc"),
            wraplength=500
        )
        self.lbl_pdf.grid(row=3, column=0, pady=(0, 10))

        # Seção de modo de processamento
        # mode_section = ctk.CTkFrame(main_frame, fg_color="transparent")
        # mode_section.grid(row=1, column=0, sticky="ew", padx=15, pady=5)
        # mode_section.grid_columnconfigure(0, weight=1)

        # mode_label = ctk.CTkLabel(
        # mode_section,
        # text="⚙️ Modo de Processamento:",
        # font=ctk.CTkFont(size=12, weight="bold")
        # )
        # mode_label.grid(row=0, column=0, pady=(0, 8))

        # Frame para os radio buttons
        # radio_frame = ctk.CTkFrame(mode_section, fg_color="transparent")
        # radio_frame.grid(row=1, column=0, sticky="ew")

        # Radio buttons para modo
        # auto_radio = ctk.CTkRadioButton(
        # radio_frame,
        # text="🔄 Automático (detecta tipo)",
        # variable=self.modo_processamento,
        # value="auto",
        # font=ctk.CTkFont(size=11),
        # text_color=("#666666", "#cccccc")
        # )
        # auto_radio.grid(row=0, column=0, sticky="w", padx=10, pady=2)

        # bling_radio = ctk.CTkRadioButton(
        # radio_frame,
        # text="📄 Bling (DANFE - 3 por folha)",
        # variable=self.modo_processamento,
        # value="bling",
        # font=ctk.CTkFont(size=11),
        # text_color=("#666666", "#cccccc")
        # )
        # bling_radio.grid(row=1, column=0, sticky="w", padx=10, pady=2)

        # ml_radio = ctk.CTkRadioButton(
        # radio_frame,
        # text="📋 ML (SKUs com índices)",
        # variable=self.modo_processamento,
        # value="ml",
        # font=ctk.CTkFont(size=11),
        # text_color=("#666666", "#cccccc")
        # )
        # ml_radio.grid(row=2, column=0, sticky="w", padx=10, pady=2)

        options_section = ctk.CTkFrame(main_frame, fg_color="transparent")
        options_section.grid(row=2, column=0, sticky="ew", padx=15, pady=5)
        options_section.grid_columnconfigure(0, weight=1)

        texto_label = ctk.CTkLabel(
            options_section,
            text="🏪 Nome da Loja (opcional):",
            font=ctk.CTkFont(size=12, weight="bold")
        )
        texto_label.grid(row=0, column=0, pady=(0, 10))

        self.texto_entry = ctk.CTkEntry(
            options_section,
            textvariable=self.texto_personalizado,
            width=250,
            height=30,
            corner_radius=25,
            font=ctk.CTkFont(size=12)
        )
        self.texto_entry.grid(row=1, column=0)

        process_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        process_frame.grid(row=3, column=0, sticky="ew", padx=15, pady=10)
        process_frame.grid_columnconfigure(0, weight=1)

        self.process_btn = ctk.CTkButton(
            process_frame,
            text="🚀 Processar PDF",
            command=self.processar,
            height=40,
            font=ctk.CTkFont(size=16, weight="bold"),
            corner_radius=25,
            fg_color=("#1f538d", "#14375e"),
            hover_color=("#2a6bb3", "#1a4a7a")
        )
        self.process_btn.grid(row=0, column=0, pady=5)

    def create_footer(self):
        """Cria o rodapé com informações"""
        footer_frame = ctk.CTkFrame(
            self, height=40, corner_radius=0, fg_color="transparent")
        footer_frame.grid(row=2, column=0, sticky="ew", padx=0, pady=0)
        footer_frame.grid_columnconfigure(0, weight=1)

        version_text = f"v{get_version()}" if UPDATE_AVAILABLE else "v1.0.3"
        system_info = ctk.CTkLabel(
            footer_frame,
            text=f"Desenvolvido por Lucas Alexandre • {version_text}",
            font=ctk.CTkFont(size=9),
            text_color=("#999999", "#666666")
        )
        system_info.grid(row=1, column=0, pady=(0, 5))

        self.about = ctk.CTkButton(
            footer_frame,
            text="",
            height=10,
            width=10,
            font=ctk.CTkFont(size=10, weight="bold"),
            corner_radius=30,
            fg_color=("#1499aa", "#1c4b81"),
            hover_color=("#2a6bb3", "#1a4a7a"),
            command=self.open_sku_panel
        )
        self.about.grid(row=1, column=0, pady=0, padx=15, sticky="e")

    def selecionar_pdf(self):
        """Seleciona arquivo PDF com feedback visual melhorado"""
        caminho = filedialog.askopenfilename(
            title="Selecionar PDF de DANFE",
            filetypes=[("Arquivos PDF", "*.pdf"), ("Todos os arquivos", "*.*")]
        )
        if caminho:
            self.arquivo_entrada = caminho
            nome_arquivo = os.path.basename(caminho)
            self.lbl_pdf.configure(
                text=f"✅ {nome_arquivo}",
                text_color=("#2d5a27", "#90ee90")
            )
            self.process_btn.configure(state="normal")

    def processar(self):
        """Processa o PDF com feedback visual"""
        if not self.arquivo_entrada:
            messagebox.showwarning(
                "Atenção", "Selecione um arquivo PDF primeiro.")
            return

        self.process_btn.configure(state="disabled", text="⏳ Processando...")
        self.update()

        diretorio_origem = os.path.dirname(self.arquivo_entrada)
        nome_origem = os.path.splitext(
            os.path.basename(self.arquivo_entrada))[0]

        # Determinar modo de processamento
        modo_selecionado = self.modo_processamento.get()

        try:
            # Se modo automático, detectar tipo de folha
            if modo_selecionado == "auto":
                import fitz
                doc_temp = fitz.open(self.arquivo_entrada)
                tipo_detectado = detectar_tipo_folha(doc_temp)
                doc_temp.close()

                self.lbl_pdf.configure(
                    text=f"🔍 Detectado: {tipo_detectado}",
                    text_color=("#1f538d", "#87ceeb")
                )
                self.update()
                time.sleep(1)

                modo_final = tipo_detectado.lower()
            else:
                modo_final = modo_selecionado

            # Processar baseado no modo
            if modo_final == "ml":
                saida = os.path.join(
                    diretorio_origem, f"{nome_origem}_Indice.pdf")

                self.lbl_pdf.configure(
                    text="🔄 Processando SKUs... Aguarde...",
                    text_color=("#1f538d", "#87ceeb")
                )
                self.update()

                pdf_saida, csv_saida, skus_updated = processar_pdf_ml(
                    self.arquivo_entrada,
                    modo="global"
                )

                if os.path.exists(pdf_saida) and os.path.getsize(pdf_saida) > 0:
                    nome_saida = os.path.basename(pdf_saida)
                    self.lbl_pdf.configure(
                        text=f"✅ {nome_saida}",
                        text_color=("#2d5a27", "#90ee90")
                    )

                    mensagem = f"PDF processado com sucesso!\n\n📁 Arquivo: {nome_saida}\n📂 Local: {diretorio_origem}"
                    if csv_saida:
                        mensagem += f"\n📊 CSV: {os.path.basename(csv_saida)}"
                    if skus_updated:
                        mensagem += "\n\n📊 Planilha de SKUs atualizada com os itens do PDF."

                    if abrir_arquivo_para_impressao(pdf_saida):
                        mensagem += "\n\n🖨️ Arquivo aberto automaticamente!"

                    messagebox.showinfo("🎉 Sucesso!", mensagem)
                else:
                    self.lbl_pdf.configure(
                        text="❌ Erro ao criar arquivo",
                        text_color=("#d32f2f", "#ffcdd2")
                    )
                    messagebox.showerror(
                        "Erro", f"Arquivo não foi criado corretamente:\n{pdf_saida}")

            else:  # Modo Bling (DANFE)
                saida = os.path.join(
                    diretorio_origem, f"{nome_origem}-Atualizado.pdf")

                self.lbl_pdf.configure(
                    text="🔄 Compactando DANFE... Aguarde...",
                    text_color=("#1f538d", "#87ceeb")
                )
                self.update()

                result = compactar_pdf_com_recorte(
                    self.arquivo_entrada,
                    saida,
                    etiquetas_por_pagina=3,
                    destacar_textos=self.destacar_textos.get(),
                    texto_personalizado=self.texto_personalizado.get()
                )

                skus_updated = False
                if isinstance(result, dict):
                    skus_updated = bool(result.get("skus_updated"))

                if os.path.exists(saida) and os.path.getsize(saida) > 0:
                    time.sleep(0.5)

                    nome_saida = os.path.basename(saida)
                    self.lbl_pdf.configure(
                        text=f"✅ {nome_saida}",
                        text_color=("#2d5a27", "#90ee90")
                    )

                    mensagem = (
                        f"PDF compactado e salvo com sucesso!\n\n"
                        f"📁 Arquivo: {nome_saida}\n"
                        f"📂 Local: {diretorio_origem}"
                    )

                    if skus_updated:
                        mensagem += "\n\n📊 Planilha de SKUs atualizada com os itens do PDF."

                    if abrir_arquivo_para_impressao(saida):
                        mensagem += "\n\n🖨️ Arquivo aberto automaticamente para impressão!"
                        messagebox.showinfo("🎉 Sucesso!", mensagem)
                    else:
                        messagebox.showinfo("✅ Sucesso!", mensagem)
                else:
                    self.lbl_pdf.configure(
                        text="❌ Erro ao criar arquivo",
                        text_color=("#d32f2f", "#ffcdd2")
                    )
                    messagebox.showerror("Erro",
                                         f"Arquivo não foi criado corretamente:\n{saida}")

        except Exception as e:
            self.lbl_pdf.configure(
                text="❌ Erro no processamento",
                text_color=("#d32f2f", "#ffcdd2")
            )
            messagebox.showerror("Erro", f"Ocorreu um problema:\n{e}")
        finally:
            self.process_btn.configure(
                state="normal", text="🚀 Processar PDF")


def create_main_app():
    """Cria a aplicação principal"""
    if not verificar_dependencias():
        print(
            "Erro: Dependências não encontradas. Execute 'pip install -r requirements.txt'")
        input("Pressione Enter para sair...")
        exit(1)

    configurar_ambiente_consistente()

    try:
        return PDFApp()
    except Exception as e:
        print(f"Erro ao iniciar aplicação: {e}")
        input("Pressione Enter para sair...")
        exit(1)


if __name__ == "__main__":
    if LOADING_AVAILABLE:
        # show_loading_screen(create_main_app)
        # else:
        app = create_main_app()
        app.mainloop()
