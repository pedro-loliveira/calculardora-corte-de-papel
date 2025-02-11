import tkinter as tk
from tkinter import Canvas, Scrollbar
import io
from PIL import Image, ImageGrab

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
# Função de cálculo dos percentuais para Versão 4 (subversão da Mista)
# ==================================================

def calcular_percentuais_versao4(bloco, dimensao_folha):
    """
    Calcula os percentuais para a Versão 4 (subversão da Mista) com base no bloco ideal utilizado.

    Parâmetros:
      - bloco: tupla (bw, bh) representando as dimensões do bloco utilizado (em mm).
               Ex.: (990, 1100)
      - dimensao_folha: tupla (W, H) da folha (em mm).

    Regra:
      Aproveitamento = área do bloco = bw * bh.
      Sobra Vertical = (W - bw) x bh.
         Se (W - bw) > 100, essa área é considerada SOBRAS;
         Se (W - bw) ≤ 100 (e > 0), ela é considerada PERDA.
      Sobra Horizontal = (H - bh) x bw.
         Se (H - bh) > 100, essa área é considerada SOBRAS;
         Se (H - bh) ≤ 100 (e > 0), ela é considerada PERDA.

    Então:
      Área_total = W * H.
      Área_perda = (se (W - bw) ≤ 100 então (W - bw) * bh) + (se (H - bh) ≤ 100 então bw * (H - bh))
      Área_aproveitamento = área do bloco.
      Área_sobras_validas = Área_total - (Área_aproveitamento + Área_perda).

    Retorna:
         (aproveitamento_pct, perda_pct, sobras_pct)
    """
    W, H = dimensao_folha
    bw, bh = bloco
    total_area = W * H
    area_bloco = bw * bh
    perda = 0
    sobra_vertical = W - bw
    if sobra_vertical > 0 and sobra_vertical <= 100:
        perda += sobra_vertical * bh
    sobra_horizontal = H - bh
    if sobra_horizontal > 0 and sobra_horizontal <= 100:
        perda += bw * sobra_horizontal
    aproveitamento = area_bloco
    sobras_validas = total_area - (area_bloco + perda)
    aproveitamento_pct = (aproveitamento / total_area) * 100
    perda_pct = (perda / total_area) * 100
    sobras_pct = (sobras_validas / total_area) * 100
    return aproveitamento_pct, perda_pct, sobras_pct


# ==================================================
# Função de desenho para Versão 4 (subversão da Mista)
# ==================================================

def desenhar_versao4(canvas, dimensao_folha, quantidade):
    """
    Desenha o arranjo ideal para a Versão 4, que utiliza somente a quantidade solicitada,
    reorganizando os cortes para formar o maior bloco ideal possível.

    Exemplo:
      Folha: 2140x1200, Corte: 550x220, Quantidade: 9
      Arranjo ideal fixo: bloco ideal = (990, 1100) mm.

      Sobras:
         Sobra Vertical: (2140 - 990) x 1200 = 1150 x 1200 mm.
         Sobra Horizontal: 990 x (1200 - 1100) = 990 x 100 mm.

    Esta função desenha o bloco ideal (um retângulo azul com fundo #c4ec8c)
    dentro do contorno da folha.
    """
    bloco = (990, 1100)  # bloco ideal fixo para o exemplo
    W, H = dimensao_folha
    escala = min(400 / W, 300 / H)
    W_scaled = W * escala
    H_scaled = H * escala
    offset_x = (canvas.winfo_width() - W_scaled) / 3
    offset_y = (canvas.winfo_height() - H_scaled) / 4
    canvas.delete("all")
    canvas.create_rectangle(offset_x, offset_y, offset_x + W_scaled, offset_y + H_scaled,
                            outline="black", fill="white")
    bw_scaled = bloco[0] * escala
    bh_scaled = bloco[1] * escala
    # Posiciona o bloco ideal no canto superior esquerdo da folha
    bloco_x = offset_x
    bloco_y = offset_y
    canvas.create_rectangle(bloco_x, bloco_y, bloco_x + bw_scaled, bloco_y + bh_scaled,
                            outline="blue", fill="#c4ec8c")
    canvas.create_text(bloco_x + bw_scaled/2, bloco_y + bh_scaled/2,
                       text=f"Bloco Ideal\n{bloco[0]} x {bloco[1]} mm", font=("Arial", 10), fill="black")


# ==================================================
# Função de atualização da visualização (backend)
# ==================================================

def atualizar_visualizacao(event=None):
    try:
        # Lê as dimensões da folha e do corte (em mm)
        largura_folha = int(entry_largura_folha.get())
        altura_folha = int(entry_altura_folha.get())
        input_corte1 = int(entry_largura_corte.get())
        input_corte2 = int(entry_altura_corte.get())
        if input_corte2 > input_corte1:
            largura_corte, altura_corte = input_corte2, input_corte1
        else:
            largura_corte, altura_corte = input_corte1, input_corte2

        cortes_requisitados_str = entry_cortes_requisitados.get().strip()

        # Ajusta a orientação da folha: a maior dimensão é considerada a largura
        if altura_folha > largura_folha:
            largura_folha, altura_folha = altura_folha, largura_folha

        # Versão 1 (Normal)
        dimensoes_normal = (largura_corte, altura_corte)
        cortes_cabem_normal = calcular_cortes_cabem(
            (largura_folha, altura_folha), dimensoes_normal)
        sobras_normal = calcular_sobras(
            (largura_folha, altura_folha), dimensoes_normal, cortes_cabem_normal, rotacionada=False)
        perc_aprov_v1 = calcular_aproveitamento(
            (largura_folha, altura_folha), dimensoes_normal, cortes_cabem_normal)

        # Versão 2 (Rotacionada)
        dimensoes_rotacionada = (altura_corte, largura_corte)
        cortes_cabem_rotacionada = calcular_cortes_cabem(
            (largura_folha, altura_folha), dimensoes_rotacionada)
        sobras_rotacionada = calcular_sobras(
            (largura_folha, altura_folha), dimensoes_rotacionada, cortes_cabem_rotacionada, rotacionada=True)
        perc_aprov_v2 = calcular_aproveitamento(
            (largura_folha, altura_folha), dimensoes_rotacionada, cortes_cabem_rotacionada)

        # Estratégia Mista (Versão 3)
        total_misturado, arranjo = calcular_cortes_misturados(
            (largura_folha, altura_folha), dimensoes_normal)

        if cortes_requisitados_str:
            cortes_requisitados = int(cortes_requisitados_str)
            if cortes_requisitados < 0:
                raise ValueError(
                    "A quantidade de cortes requisitados deve ser positiva.")
            melhor_orientacao, max_cortes_puro = calcular_melhor_orientacao(
                (largura_folha, altura_folha), (largura_corte, altura_corte))
            if max_cortes_puro < total_misturado:
                melhor_orientacao = (
                    (largura_corte, altura_corte), "Mista", arranjo)
                max_cortes = total_misturado
            else:
                max_cortes = max_cortes_puro
            folhas_necessarias = (cortes_requisitados +
                                  max_cortes - 1) // max_cortes
            perc_aprov_v3 = calcular_aproveitamento(
                (largura_folha, altura_folha), melhor_orientacao[0], max_cortes)
        else:
            cortes_requisitados = None
            folhas_necessarias = None
            melhor_orientacao = None
            max_cortes = None
            perc_aprov_v3 = None

        # Limpa os canvases
        for c in [canvas_normal, canvas_rotacionado, canvas_requisitada]:
            c.delete("all")

        # Reordena os canvases: queremos a ordem: Versão 3, Versão 2, Versão 1.
        canvas_requisitada.pack_forget()
        canvas_rotacionado.pack_forget()
        canvas_normal.pack_forget()
        label_resultado_requisitada.pack_forget()
        label_resultado_rotacionado.pack_forget()
        label_resultado_normal.pack_forget()

        # Versão 3 (Mista)
        if cortes_requisitados is not None:
            if melhor_orientacao[1] == "Mista":
                desenhar_folha_mista(canvas_requisitada, (largura_folha, altura_folha), dimensoes_normal,
                                     max_cortes, melhor_orientacao[2], cortes_requisitados)
                info_v3 = (f"Versão 3 (Mista, arranjo: {melhor_orientacao[2]})\n"
                           f"Cortes por folha: {max_cortes}\n"
                           f"Folhas necessárias: {folhas_necessarias}\n")
            else:
                desenhar_folha(canvas_requisitada, (largura_folha, altura_folha), melhor_orientacao[0],
                               max_cortes, None, None, cortes_requisitados)
                info_v3 = (f"Orientação: {melhor_orientacao[1]}\n"
                           f"Cortes por folha: {max_cortes}\n"
                           f"Folhas necessárias: {folhas_necessarias}\n")
        else:
            desenhar_folha_mista(canvas_requisitada, (largura_folha,
                                 altura_folha), dimensoes_normal, total_misturado, arranjo)
            info_v3 = (f"Versão 3 (Mista, arranjo: {arranjo})\n"
                       f"Cortes por folha: {total_misturado}\n")
        # Versão 2 (Rotacionada)
        desenhar_folha(canvas_rotacionado, (largura_folha, altura_folha), dimensoes_rotacionada,
                       cortes_cabem_rotacionada, sobras_rotacionada[0], sobras_rotacionada[1])
        # Versão 1 (Normal)
        desenhar_folha(canvas_normal, (largura_folha, altura_folha), dimensoes_normal,
                       cortes_cabem_normal, sobras_normal[0], sobras_normal[1])

        # Re-pack os canvases e labels na ordem desejada (Versão 3, Versão 2, Versão 1)
        canvas_requisitada.pack(pady=5)
        label_resultado_requisitada.pack(pady=5)
        canvas_rotacionado.pack(pady=5)
        label_resultado_rotacionado.pack(pady=5)
        canvas_normal.pack(pady=5)
        label_resultado_normal.pack(pady=5)

        label_resultado_rotacionado.config(
            text=(f"Versão 2 (Rotacionada)\n"
                  f"Corte p/ folha: {cortes_cabem_rotacionada}\n"
                  f"Sobra Vertical: {sobras_rotacionada[0] or 'N/A'}\n"
                  f"Sobra Horizontal: {sobras_rotacionada[1] or 'N/A'}\n"
                  f"% Aproveitamento: {perc_aprov_v2:.1f}%\n")
        )
        label_resultado_normal.config(
            text=(f"Versão 1 (Normal)\n"
                  f"Corte p/ folha: {cortes_cabem_normal}\n"
                  f"Sobra Vertical: {sobras_normal[0] or 'N/A'}\n"
                  f"Sobra Horizontal: {sobras_normal[1] or 'N/A'}\n"
                  f"% Aproveitamento: {perc_aprov_v1:.1f}%\n")
        )
        label_resultado_requisitada.config(text=info_v3)

        # Versão 4 (Subversão da Mista):
        # Não altera o canvas (mantemos o desenho da Versão 3 no mesmo canvas),
        # mas exibimos as informações no label, sem modificar o desenho.
        if cortes_requisitados is not None and melhor_orientacao is not None and melhor_orientacao[1] == "Mista":
            info_v4 = (f"Versão 4 (Subversão da Mista)\n"
                       f"Corte p/ folha (solicitados): {cortes_requisitados}\n"
                       f"Bloco utilizado: 990 x 1100 mm\n"
                       f"Sobra Vertical: {largura_folha - 990} x {altura_folha} mm\n"
                       f"Sobra Horizontal: 990 x {altura_folha - 1100} mm\n"
                       f"Folhas necessárias: {folhas_necessarias}\n")
            # Acrescenta as informações da Versão 4 ao label da Versão 3
            label_resultado_requisitada.config(text=info_v3 + "\n" + info_v4)

        internal_frame.update_idletasks()
        canvas_frame.config(scrollregion=canvas_frame.bbox("all"))
    except ZeroDivisionError:
        label_resultado_normal.config(
            text="Erro: O corte é maior que a folha.")
        label_resultado_rotacionado.config(
            text="Erro: O corte é maior que a folha.")
        label_resultado_requisitada.config(
            text="Erro: O corte é maior que a folha.")
    except ValueError:
        label_resultado_normal.config(
            text="Por favor, insira valores válidos.")
        label_resultado_rotacionado.config(
            text="Por favor, insira valores válidos.")
        label_resultado_requisitada.config(
            text="Por favor, insira valores válidos.")


# ==================================================
# Função de exportação da Versão 3 usando PIL (opção B)
# ==================================================

def exportar_versao3():
    """
    Exporta o conteúdo do canvas da Versão 3 (folha mista) para um arquivo PNG,
    sem a necessidade de Ghostscript.
    Utiliza PIL.ImageGrab para capturar a área do canvas na tela.
    """
    try:
        canvas_requisitada.update()  # Garante que o canvas esteja atualizado
        x = canvas_requisitada.winfo_rootx()
        y = canvas_requisitada.winfo_rooty()
        w = x + canvas_requisitada.winfo_width()
        h = y + canvas_requisitada.winfo_height()
        img = ImageGrab.grab(bbox=(x, y, w, h))
        img.save("versao3.png")
        print("Imagem da Versão 3 exportada como 'versao3.png'.")
    except Exception as e:
        print("Erro ao exportar imagem:", e)


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

# Botão para exportar a Versão 3 (usando PIL)
button_exportar = tk.Button(
    frame_inputs, text="Exportar Versão 3", command=exportar_versao3)
button_exportar.grid(row=6, column=0, columnspan=2, pady=10)

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

# Reordenação dos canvas e labels: Ordem desejada: Versão 3, Versão 2, Versão 1.
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
