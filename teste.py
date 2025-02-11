import tkinter as tk
from tkinter import Canvas, Scrollbar
import io
from PIL import Image, ImageOps
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib.styles import getSampleStyleSheet

# Variáveis globais para armazenar os textos com quebras HTML para o PDF
info_v1_pdf = ""
info_v2_pdf = ""
info_v3_pdf = ""

# ==================================================
# Funções de desenho das folhas (frontend)
# ==================================================


def desenhar_folha(canvas, dimensao_folha, dimensoes_corte, cortes_cabem,
                   sobra1=None, sobra2=None, cortes_requisitados=None):
    W, H = dimensao_folha
    pW, pH = dimensoes_corte
    canvas.delete("all")
    escala = min(400 / W, 300 / H)
    W_scaled = W * escala
    H_scaled = H * escala
    pW_scaled = pW * escala
    pH_scaled = pH * escala
    offset_x = (canvas.winfo_width() - W_scaled) / 3
    offset_y = (canvas.winfo_height() - H_scaled) / 4
    canvas.create_rectangle(offset_x, offset_y, offset_x + W_scaled, offset_y + H_scaled,
                            outline="black", fill="white")
    n_horiz = W // pW
    n_vert = H // pH
    for coluna in range(n_horiz):
        for linha in range(n_vert):
            index = linha + coluna * n_vert
            if index >= cortes_cabem:
                break
            cor = "#ADD8E6"
            if cortes_requisitados is not None and index >= cortes_requisitados:
                cor = "#dbdbdb"
            pos_x = coluna * pW_scaled
            pos_y = linha * pH_scaled
            canvas.create_rectangle(offset_x + pos_x, offset_y + pos_y,
                                    offset_x + pos_x + pW_scaled, offset_y + pos_y + pH_scaled,
                                    outline="black", fill=cor)
            canvas.create_text(offset_x + pos_x + pW_scaled/2, offset_y + pos_y + pH_scaled/2,
                               text=str(index + 1), font=("Arial", 5), fill="black")

    def add_sobra_info(sobra, position, angle=0):
        if sobra:
            try:
                dim1, dim2 = map(int, sobra.split('x'))
                canvas.create_text(position[0], position[1],
                                   text=f"Sobra: {sobra}", font=("Arial", 8), fill="blue", angle=angle)
            except Exception:
                pass
    if sobra1:
        add_sobra_info(sobra1, (offset_x + W_scaled + 10,
                       offset_y + H_scaled/2), angle=90)
    if sobra2:
        add_sobra_info(sobra2, (offset_x + W_scaled /
                       2, offset_y + H_scaled + 10))
    canvas.create_line(offset_x, offset_y - 10, offset_x, offset_y + H_scaled + 10,
                       fill="red", width=2)
    canvas.create_line(offset_x - 10, offset_y, offset_x + W_scaled + 10, offset_y,
                       fill="red", width=2)
    canvas.create_text(offset_x - 10, offset_y + H_scaled/2,
                       text=f"{H} mm", font=("Arial", 8), fill="red", angle=90)
    canvas.create_text(offset_x + W_scaled/2, offset_y - 10,
                       text=f"{W} mm", font=("Arial", 8), fill="red")


def desenhar_folha_mista(canvas, dimensao_folha, dimensoes_corte, cortes_cabem, arranjo, cortes_requisitados=None):
    W, H = dimensao_folha
    escala = min(400 / W, 300 / H)
    offset_x = (canvas.winfo_width() - W * escala) / 3
    offset_y = (canvas.winfo_height() - H * escala) / 4
    canvas.delete("all")
    canvas.create_rectangle(offset_x, offset_y, offset_x + W * escala, offset_y + H * escala,
                            outline="black", fill="white")
    pW, pH = dimensoes_corte
    if arranjo == "normal":
        base_pW, base_pH = pW, pH
        extra_dims = (pH, pW)
    else:
        base_pW, base_pH = pH, pW
        extra_dims = (pW, pH)
    base_pW_scaled = base_pW * escala
    base_pH_scaled = base_pH * escala
    n_horiz = W // base_pW
    n_vert = H // base_pH
    count = 0
    for coluna in range(n_horiz):
        for linha in range(n_vert):
            index = count
            cor = "#c4ec8c" if (
                cortes_requisitados is None or index < cortes_requisitados) else "#dbdbdb"
            x = coluna * base_pW_scaled
            y = linha * base_pH_scaled
            canvas.create_rectangle(offset_x + x, offset_y + y,
                                    offset_x + x + base_pW_scaled, offset_y + y + base_pH_scaled,
                                    outline="black", fill=cor)
            canvas.create_text(offset_x + x + base_pW_scaled/2, offset_y + y + base_pH_scaled/2,
                               text=str(index + 1), font=("Arial", 8), fill="black")
            count += 1
    extra_cell_width, extra_cell_height = extra_dims
    extra_cell_width_scaled = extra_cell_width * escala
    extra_cell_height_scaled = extra_cell_height * escala
    if arranjo == "normal":
        leftover_right = W - n_horiz * base_pW
        if leftover_right >= pH:
            extra_cols = leftover_right // pH
            extra_rows = H // pW
            for col in range(extra_cols):
                for row in range(extra_rows):
                    index = count
                    cor = "#c4ec8c" if (
                        cortes_requisitados is None or index < cortes_requisitados) else "#dbdbdb"
                    x = n_horiz * base_pW_scaled + col * extra_cell_width_scaled
                    y = row * extra_cell_height_scaled
                    canvas.create_rectangle(offset_x + x, offset_y + y,
                                            offset_x + x + extra_cell_width_scaled, offset_y + y + extra_cell_height_scaled,
                                            outline="black", fill=cor)
                    canvas.create_text(offset_x + x + extra_cell_width_scaled/2, offset_y + y + extra_cell_height_scaled/2,
                                       text=str(index + 1), font=("Arial", 8), fill="black")
                    count += 1
        leftover_bottom = H - n_vert * base_pH
        if leftover_bottom >= pH:
            extra_cols = W // pW
            extra_rows = leftover_bottom // pH
            for col in range(extra_cols):
                for row in range(extra_rows):
                    index = count
                    cor = "#c4ec8c" if (
                        cortes_requisitados is None or index < cortes_requisitados) else "#dbdbdb"
                    x = col * extra_cell_width_scaled
                    y = n_vert * base_pH_scaled + row * extra_cell_height_scaled
                    canvas.create_rectangle(offset_x + x, offset_y + y,
                                            offset_x + x + extra_cell_width_scaled, offset_y + y + extra_cell_height_scaled,
                                            outline="black", fill=cor)
                    canvas.create_text(offset_x + x + extra_cell_width_scaled/2, offset_y + y + extra_cell_height_scaled/2,
                                       text=str(index + 1), font=("Arial", 8), fill="black")
                    count += 1
    else:
        leftover_right = W - n_horiz * base_pW
        if leftover_right >= pH:
            extra_cols = leftover_right // pH
            extra_rows = H // pW
            for col in range(extra_cols):
                for row in range(extra_rows):
                    index = count
                    cor = "#c4ec8c" if (
                        cortes_requisitados is None or index < cortes_requisitados) else "#dbdbdb"
                    x = n_horiz * base_pW_scaled + col * extra_cell_width_scaled
                    y = row * extra_cell_height_scaled
                    canvas.create_rectangle(offset_x + x, offset_y + y,
                                            offset_x + x + extra_cell_width_scaled, offset_y + y + extra_cell_height_scaled,
                                            outline="black", fill=cor)
                    canvas.create_text(offset_x + x + extra_cell_width_scaled/2, offset_y + y + extra_cell_height_scaled/2,
                                       text=str(index + 1), font=("Arial", 8), fill="black")
                    count += 1
        leftover_bottom = H - n_vert * base_pH
        if leftover_bottom >= pH:
            extra_cols = W // pW
            extra_rows = leftover_bottom // pH
            for col in range(extra_cols):
                for row in range(extra_rows):
                    index = count
                    cor = "#c4ec8c" if (
                        cortes_requisitados is None or index < cortes_requisitados) else "#dbdbdb"
                    x = col * extra_cell_width_scaled
                    y = n_vert * base_pH_scaled + row * extra_cell_height_scaled
                    canvas.create_rectangle(offset_x + x, offset_y + y,
                                            offset_x + x + extra_cell_width_scaled, offset_y + y + extra_cell_height_scaled,
                                            outline="black", fill=cor)
                    canvas.create_text(offset_x + x + extra_cell_width_scaled/2, offset_y + y + extra_cell_height_scaled/2,
                                       text=str(index + 1), font=("Arial", 8), fill="black")
                    count += 1
    canvas.create_line(offset_x, offset_y - 10, offset_x, offset_y + H * escala + 10,
                       fill="red", width=2)
    canvas.create_line(offset_x - 10, offset_y, offset_x + W * escala + 10, offset_y,
                       fill="red", width=2)
    canvas.create_text(offset_x - 10, offset_y + H * escala / 2,
                       text=f"{H} mm", font=("Arial", 8), fill="red", angle=90)
    canvas.create_text(offset_x + W * escala / 2, offset_y - 10,
                       text=f"{W} mm", font=("Arial", 8), fill="red")


# ==================================================
# Funções de cálculo (backend)
# ==================================================

def calcular_cortes_cabem(dimensao_folha, dimensoes_corte):
    W, H = dimensao_folha
    pW, pH = dimensoes_corte
    if pW == 0 or pH == 0:
        return 0
    n_horiz = W // pW
    n_vert = H // pH
    return n_horiz * n_vert


def calcular_sobras(dimensao_folha, dimensoes_corte, cortes_cabem, rotacionada=False):
    W, H = dimensao_folha
    pW, pH = dimensoes_corte
    if pW == 0 or pH == 0:
        return None, None
    n_horiz = W // pW
    n_vert = H // pH
    if not rotacionada:
        sobra_lateral = W - n_horiz * pW
        sobra_inferior = H - n_vert * pH
        s1 = f"{sobra_lateral}x{H}" if sobra_lateral >= 100 else None
        s2 = f"{n_horiz * pW}x{sobra_inferior}" if sobra_inferior >= 100 else None
        return s1, s2
    else:
        sobra_lateral = W - n_horiz * pW
        sobra_horizontal = H - n_vert * pH
        s1 = f"{sobra_lateral}x{H}" if sobra_lateral >= 100 else None
        s2 = f"{W}x{sobra_horizontal}" if sobra_horizontal >= 100 else None
        if sobra_lateral >= 100 and sobra_horizontal >= 100:
            area_lateral = sobra_lateral * H
            area_horizontal = W * sobra_horizontal
            if area_horizontal > area_lateral:
                s1 = f"{sobra_lateral}x{n_vert * pH}"
        return s1, s2


def calcular_cortes_misturados(dimensao_folha, dimensoes_corte):
    W, H = dimensao_folha
    pW, pH = dimensoes_corte
    n_horiz_1 = W // pW
    n_vert_1 = H // pH
    base1 = n_horiz_1 * n_vert_1
    sobra_largura_1 = W - n_horiz_1 * pW
    sobra_altura_1 = H - n_vert_1 * pH
    extra1 = 0
    if sobra_largura_1 >= pH:
        extra1 = max(extra1, (sobra_largura_1 // pH) * (H // pW))
    if sobra_altura_1 >= pH:
        extra1 = max(extra1, (sobra_altura_1 // pH) * (W // pW))
    total1 = base1 + extra1
    n_horiz_2 = W // pH
    n_vert_2 = H // pW
    base2 = n_horiz_2 * n_vert_2
    sobra_largura_2 = W - n_horiz_2 * pH
    sobra_altura_2 = H - n_vert_2 * pW
    extra2 = 0
    if sobra_largura_2 >= pH:
        extra2 = max(extra2, (sobra_largura_2 // pH) * (H // pW))
    if sobra_altura_2 >= pH:
        extra2 = max(extra2, (sobra_altura_2 // pH) * (W // pW))
    total2 = base2 + extra2
    if total1 >= total2:
        return total1, "normal"
    else:
        return total2, "rotacionada"


def calcular_sobras_mistura(dimensao_folha, dimensoes_corte):
    W, H = dimensao_folha
    pW, pH = dimensoes_corte
    n_horiz = W // pW
    n_vert = H // pH
    sobra_largura = W - n_horiz * pW
    sobra_inferior = H - n_vert * pH
    s_lateral = f"{sobra_largura}x{H}" if sobra_largura >= 100 else None
    s_inferior = f"{n_horiz * pW}x{sobra_inferior}" if sobra_inferior >= 100 else None
    s_inf_direita = f"{sobra_largura}x{sobra_inferior}" if (
        sobra_largura >= 100 and sobra_inferior >= 100) else None
    return s_lateral, s_inferior, s_inf_direita


def calcular_melhor_orientacao(dimensao_folha, dimensoes_corte):
    W, H = dimensao_folha
    pW, pH = dimensoes_corte
    cortes_normal = calcular_cortes_cabem((W, H), (pW, pH))
    cortes_rotacionada = calcular_cortes_cabem((W, H), (pH, pW))
    misturado, arranjo = calcular_cortes_misturados((W, H), (pW, pH))
    if misturado >= cortes_normal and misturado >= cortes_rotacionada:
        return ((pW, pH), "Mista", arranjo), misturado
    elif cortes_rotacionada >= cortes_normal:
        return ((pH, pW), "Rotacionada", None), cortes_rotacionada
    else:
        return ((pW, pH), "Normal", None), cortes_normal


def calcular_aproveitamento(dimensao_folha, dimensoes_corte, cortes_cabem):
    W, H = dimensao_folha
    pW, pH = dimensoes_corte
    area_total = W * H
    area_corte = pW * pH
    area_cortes = cortes_cabem * area_corte
    aproveitamento_pct = (area_cortes / area_total) * \
        100 if area_total > 0 else 0
    return aproveitamento_pct


# ==================================================
# Função de exportação dos canvas para PDF (via postscript)
# ==================================================
def exportar_canvas_pdf(canvas, nome_arquivo, zoom=4):
    """
    Exporta o conteúdo do canvas para um arquivo PDF utilizando o método postscript.
    O parâmetro zoom aumenta a resolução.
    """
    try:
        ps = canvas.postscript(colormode='color',
                               pagewidth=canvas.winfo_width()*zoom,
                               pageheight=canvas.winfo_height()*zoom)
        img = Image.open(io.BytesIO(ps.encode('utf-8')))
        img = img.convert("RGB")
        img.save(nome_arquivo, "PDF")
        print(f"Canvas exportado como {nome_arquivo}.")
    except Exception as e:
        print("Erro ao exportar canvas:", e)


# ==================================================
# Função de exportação do Relatório de Corte de Papel (PDF com layout único)
# ==================================================
def exportar_relatorio_pdf():
    """
    Gera um PDF com o relatório de corte de papel para as versões V1 (Normal), V2 (Rotacionada) e V3 (Mista)
    em uma única página, utilizando os textos com <br/> para as quebras de linha.
    """
    try:
        root.update()
        zoom = 4

        # Função auxiliar para converter um canvas em RLImage
        def canvas_para_rlimage(canvas):
            ps = canvas.postscript(colormode='color',
                                   pagewidth=canvas.winfo_width()*zoom,
                                   pageheight=canvas.winfo_height()*zoom)
            pil_img = Image.open(io.BytesIO(ps.encode('utf-8')))
            pil_img = pil_img.convert("RGB")
            buf = io.BytesIO()
            pil_img.save(buf, format='PNG')
            buf.seek(0)
            return RLImage(buf)

        # Gera as imagens dos três canvas
        rl_img_v1 = canvas_para_rlimage(canvas_normal)
        rl_img_v2 = canvas_para_rlimage(canvas_rotacionado)
        rl_img_v3 = canvas_para_rlimage(canvas_requisitada)

        # Obtemos os textos formatados para PDF a partir das variáveis globais
        global info_v1_pdf, info_v2_pdf, info_v3_pdf
        info_total = Paragraph(info_v1_pdf + "<br/><br/>" +
                               info_v2_pdf + "<br/><br/>" +
                               info_v3_pdf,
                               getSampleStyleSheet()['Normal'])

        styles = getSampleStyleSheet()
        title = Paragraph("Relatório de Corte de Papel", styles['Title'])
        spacer = Spacer(1, 12)

        page_width, page_height = landscape(A4)
        margin = 36
        available_width = page_width - 2 * margin

        col_esq_width = available_width * 0.5
        col_dir_width = available_width * 0.5

        image_width_esq = col_esq_width
        image_height_esq = 200
        image_width_dir = col_dir_width
        image_height_dir = 200

        def ajustar_imagem(rl_img, largura, altura):
            rl_img.drawWidth = largura
            rl_img.drawHeight = altura
            return rl_img

        rl_img_v1 = ajustar_imagem(
            rl_img_v1, image_width_esq, image_height_esq)
        rl_img_v2 = ajustar_imagem(
            rl_img_v2, image_width_esq, image_height_esq)
        rl_img_v3 = ajustar_imagem(
            rl_img_v3, image_width_dir, image_height_dir)

        data = []
        data.append([title, ""])
        data.append([rl_img_v1, rl_img_v3])
        data.append([rl_img_v2, info_total])
        col_widths = [available_width / 2, available_width / 2]
        table = Table(data, colWidths=col_widths)
        table.setStyle(TableStyle([
            ('SPAN', (0, 0), (1, 0)),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('BOX', (0, 0), (-1, -1), 0.5, 'black'),
            ('INNERGRID', (0, 0), (-1, -1), 0.25, 'black'),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
        ]))

        pdf_filename = "Relatorio_Corte_Papel.pdf"
        doc = SimpleDocTemplate(pdf_filename, pagesize=landscape(A4),
                                rightMargin=margin, leftMargin=margin,
                                topMargin=margin, bottomMargin=margin)
        flowables = [table]
        doc.build(flowables)
        print(f"Relatório exportado como '{pdf_filename}'.")
    except Exception as e:
        print("Erro ao exportar relatório:", e)


# ==================================================
# Função de atualização da visualização dos canvas e labels
# ==================================================
def atualizar_visualizacao(event=None):
    global info_v1_pdf, info_v2_pdf, info_v3_pdf
    try:
        # Leitura dos valores dos inputs
        largura_folha = int(entry_largura_folha.get())
        altura_folha = int(entry_altura_folha.get())
        largura_corte = int(entry_largura_corte.get())
        altura_corte = int(entry_altura_corte.get())
    except ValueError:
        return

    try:
        if entry_cortes_requisitados.get().strip():
            cortes_requisitados = int(entry_cortes_requisitados.get().strip())
        else:
            cortes_requisitados = None
    except ValueError:
        cortes_requisitados = None

    A_total = largura_folha * altura_folha

    # --- V1 – Versão Normal (Horiz.) ---
    cortes_v1 = calcular_cortes_cabem(
        (largura_folha, altura_folha), (largura_corte, altura_corte))
    desenhar_folha(canvas_normal, (largura_folha, altura_folha),
                   (largura_corte, altura_corte), cortes_v1)
    A_used_v1 = cortes_v1 * (largura_corte * altura_corte)
    n_horiz_v1 = largura_folha // largura_corte
    n_vert_v1 = altura_folha // altura_corte
    sobra_v_v1 = largura_folha - n_horiz_v1 * largura_corte
    sobra_h_v1 = altura_folha - n_vert_v1 * altura_corte
    A_vert = sobra_v_v1 * altura_folha
    A_horiz = largura_folha * sobra_h_v1
    A_overlap = sobra_v_v1 * sobra_h_v1
    valid_vert = A_vert - A_overlap
    valid_horiz = A_horiz - A_overlap
    sobra_total = A_vert + A_horiz - A_overlap
    if sobra_total > 0:
        perc_vert_v1 = valid_vert / sobra_total * \
            (100 - (A_used_v1/A_total*100))
        perc_horiz_v1 = valid_horiz / sobra_total * \
            (100 - (A_used_v1/A_total*100))
    else:
        perc_vert_v1 = 0
        perc_horiz_v1 = 0

    # Gera o texto para visualização (com \n) e para PDF (com <br/>)
    info_v1_label = "Versão 1 (Normal)\n"
    info_v1_label += f"Corte p/ folha: {cortes_v1}\n"
    info_v1_label += f"Sobra Vertical: {sobra_v_v1}x{altura_folha}\n"
    info_v1_label += f"Sobra Horizontal: {largura_folha}x{sobra_h_v1}\n"
    info_v1_label += f"% Aproveitamento: {A_used_v1/A_total*100:.1f}%\n"
    # info_v1_label += f"Sobra Vertical: {sobra_v_v1}x{altura_folha} ({perc_vert_v1:.1f}%)\n"
    # info_v1_label += f"Sobra Horizontal: {largura_folha}x{sobra_h_v1} ({perc_horiz_v1:.1f}%)"
    label_resultado_normal.config(text=info_v1_label)
    info_v1_pdf = info_v1_label.replace("\n", "<br/>")

    # --- V2 – Versão Rotacionada (Vert.) ---
    cortes_v2 = calcular_cortes_cabem(
        (largura_folha, altura_folha), (altura_corte, largura_corte))
    desenhar_folha(canvas_rotacionado, (largura_folha, altura_folha),
                   (altura_corte, largura_corte), cortes_v2)
    A_used_v2 = cortes_v2 * (largura_corte * altura_corte)
    n_horiz_v2 = largura_folha // altura_corte
    n_vert_v2 = altura_folha // largura_corte
    sobra_v_v2 = largura_folha - n_horiz_v2 * altura_corte
    sobra_h_v2 = altura_folha - n_vert_v2 * largura_corte
    A_vert_v2 = sobra_v_v2 * altura_folha
    A_horiz_v2 = largura_folha * sobra_h_v2
    A_overlap_v2 = sobra_v_v2 * sobra_h_v2
    valid_vert_v2 = A_vert_v2 - A_overlap_v2
    valid_horiz_v2 = A_horiz_v2 - A_overlap_v2
    sobra_total_v2 = A_vert_v2 + A_horiz_v2 - A_overlap_v2
    if sobra_total_v2 > 0:
        perc_vert_v2 = valid_vert_v2 / sobra_total_v2 * \
            (100 - (A_used_v2/A_total*100))
        perc_horiz_v2 = valid_horiz_v2 / sobra_total_v2 * \
            (100 - (A_used_v2/A_total*100))
    else:
        perc_vert_v2 = 0
        perc_horiz_v2 = 0

    info_v2_label = "Versão 2 (Rotacionada)\n"
    info_v2_label += f"Corte p/ folha: {cortes_v2}\n"
    info_v2_label += f"Sobra Vertical: {sobra_v_v2}x{altura_folha}\n"
    info_v2_label += f"Sobra Horizontal: {largura_folha}x{sobra_h_v2}\n"
    info_v2_label += f"% Aproveitamento: {A_used_v2/A_total*100:.1f}%\n"
    # info_v2_label += f"Sobra Vertical: ({perc_vert_v2:.1f}%)\n"
    # info_v2_label += f"Sobra Horizontal: ({perc_horiz_v2:.1f}%)"
    label_resultado_rotacionado.config(text=info_v2_label)
    info_v2_pdf = info_v2_label.replace("\n", "<br/>")

    # --- V3 – Versão Mista ---
    cortes_misturados, arranjo_mistura = calcular_cortes_misturados(
        (largura_folha, altura_folha), (largura_corte, altura_corte))
    max_cortes = cortes_misturados
    if arranjo_mistura == "normal":
        desenhar_folha_mista(canvas_requisitada, (largura_folha, altura_folha), (largura_corte, altura_corte),
                             max_cortes, "normal", cortes_requisitados if cortes_requisitados is not None else max_cortes)
    else:
        desenhar_folha_mista(canvas_requisitada, (largura_folha, altura_folha), (largura_corte, altura_corte),
                             max_cortes, "rotacionada", cortes_requisitados if cortes_requisitados is not None else max_cortes)
    info_v3_label = f"Versão 3 (Mista, arranjo: {arranjo_mistura})\n"
    info_v3_label += f"Cortes por folha: {max_cortes}\n"
    if cortes_requisitados is None:
        info_v3_label += "Qtd Pedida: -\nFolhas Incompletas: -"
    else:
        if max_cortes > 0:
            folhas_necessarias = (cortes_requisitados +
                                  max_cortes - 1) // max_cortes
            folhas_incompletas = 1 if (
                cortes_requisitados % max_cortes != 0 and cortes_requisitados > 0) else 0
        else:
            folhas_necessarias = 0
            folhas_incompletas = 0
        info_v3_label += f"Qtd Pedida: {cortes_requisitados}\n"
        info_v3_label += f"Folhas necessárias: {folhas_necessarias}\n"
        info_v3_label += f"Folhas Incompletas: {folhas_incompletas}"
    label_resultado_requisitada.config(text=info_v3_label)
    info_v3_pdf = info_v3_label.replace("\n", "<br/>")

    internal_frame.update_idletasks()
    canvas_frame.config(scrollregion=canvas_frame.bbox("all"))
    print("Visualização atualizada.")


# ==================================================
# Configuração da interface gráfica (frontend)
# ==================================================
root = tk.Tk()
root.title("Calculadora de Corte de Papel")
root.bind("<Return>", atualizar_visualizacao)

frame_inputs = tk.Frame(root, padx=10, pady=10)
frame_inputs.pack(side=tk.LEFT)

label_largura_folha = tk.Label(frame_inputs, text="Largura da Folha (mm):")
label_largura_folha.grid(row=0, column=0, sticky="w")
entry_largura_folha = tk.Entry(frame_inputs)
entry_largura_folha.grid(row=0, column=1)
entry_largura_folha.bind("<Return>", atualizar_visualizacao)

label_altura_folha = tk.Label(frame_inputs, text="Altura da Folha (mm):")
label_altura_folha.grid(row=1, column=0, sticky="w")
entry_altura_folha = tk.Entry(frame_inputs)
entry_altura_folha.grid(row=1, column=1)
entry_altura_folha.bind("<Return>", atualizar_visualizacao)

label_largura_corte = tk.Label(frame_inputs, text="Largura do Corte (mm):")
label_largura_corte.grid(row=2, column=0, sticky="w")
entry_largura_corte = tk.Entry(frame_inputs)
entry_largura_corte.grid(row=2, column=1)
entry_largura_corte.bind("<Return>", atualizar_visualizacao)

label_altura_corte = tk.Label(frame_inputs, text="Altura do Corte (mm):")
label_altura_corte.grid(row=3, column=0, sticky="w")
entry_altura_corte = tk.Entry(frame_inputs)
entry_altura_corte.grid(row=3, column=1)
entry_altura_corte.bind("<Return>", atualizar_visualizacao)

label_cortes_requisitados = tk.Label(
    frame_inputs, text="Quantidade (opcional):")
label_cortes_requisitados.grid(row=4, column=0, sticky="w")
entry_cortes_requisitados = tk.Entry(frame_inputs)
entry_cortes_requisitados.grid(row=4, column=1)
entry_cortes_requisitados.bind("<Return>", atualizar_visualizacao)

button_calcular = tk.Button(
    frame_inputs, text="Calcular", command=atualizar_visualizacao)
button_calcular.grid(row=5, column=0, columnspan=2, pady=10)

button_exportar_relatorio = tk.Button(
    frame_inputs, text="Gerar Relatório de Corte de Papel", command=exportar_relatorio_pdf)
button_exportar_relatorio.grid(row=6, column=0, columnspan=2, pady=10)

# ==================================================
# Configuração dos widgets de exibição
# ==================================================
frame_chapas = tk.Frame(root, bg="#ececea", width=700, height=900)
frame_chapas.pack(side=tk.RIGHT, padx=10, pady=10)

canvas_frame = Canvas(frame_chapas, width=580, height=500)
canvas_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

scrollbar = tk.Scrollbar(
    frame_chapas, orient=tk.VERTICAL, command=canvas_frame.yview)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
canvas_frame.config(yscrollcommand=scrollbar.set)

internal_frame = tk.Frame(canvas_frame)
canvas_frame.create_window((0, 0), window=internal_frame, anchor="nw")

# Ordem dos canvas e labels: V3, V2, V1
canvas_requisitada = Canvas(internal_frame, width=550, height=350, bg="#ececea",
                            highlightthickness=1, highlightbackground="#b4b4b4")
canvas_requisitada.pack(pady=5)
label_resultado_requisitada = tk.Label(internal_frame, text="", justify="left")
label_resultado_requisitada.pack(pady=5)

canvas_rotacionado = Canvas(internal_frame, width=550, height=350, bg="#ececea",
                            highlightthickness=1, highlightbackground="#b4b4b4")
canvas_rotacionado.pack(pady=5)
label_resultado_rotacionado = tk.Label(internal_frame, text="", justify="left")
label_resultado_rotacionado.pack(pady=5)

canvas_normal = Canvas(internal_frame, width=550, height=350, bg="#ececea",
                       highlightthickness=1, highlightbackground="#b4b4b4")
canvas_normal.pack(pady=5)
label_resultado_normal = tk.Label(internal_frame, text="", justify="left")
label_resultado_normal.pack(pady=5)

internal_frame.update_idletasks()
canvas_frame.config(scrollregion=canvas_frame.bbox("all"))


def _on_mousewheel(event):
    canvas_frame.yview_scroll(int(-1*(event.delta/120)), "units")


canvas_frame.bind_all("<MouseWheel>", _on_mousewheel)

root.mainloop()
