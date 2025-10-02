# ============================================================
# layout.py
# ------------------------------------------------------------
# Responsável por montar o layout da Tabela de Folha de Tarefa
# utilizando placeholders substituídos por valores dinâmicos.
# ============================================================

import re
import pandas as pd
from reportlab.platypus import Table, TableStyle, Paragraph
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm


# ============================================================
# Estilos base
# ============================================================

styles = getSampleStyleSheet()

style_text = styles["Normal"]
style_text.fontSize = 9
style_text.leading = 12
style_text.alignment = 0

style_subtitle = ParagraphStyle(
    "Subtitle",
    parent=style_text,
    fontSize=9,
    leading=10,
    textColor=colors.black,
    fontName="Helvetica-Bold",
    alignment=1
)

style_crono = ParagraphStyle(
    "Crono",
    parent=style_text,
    fontSize=9,
    leading=10,
    textColor=colors.black,
    fontName="Helvetica",
    alignment=1
)

# ============================================================
# Estilos especiais
# ============================================================

style_descricao = ParagraphStyle(
    "Descricao",
    parent=style_text,
    fontName="Helvetica-Bold",
    textColor=colors.blue,
    fontSize=9,
    leading=12,
    alignment=0
)

style_passagem = ParagraphStyle(
    "PassagemServico",
    parent=style_text,
    fontName="Helvetica-Bold",
    textColor="#B30000",
    fontSize=8,
    leading=12,
    alignment=0
)


# ============================================================
# Regex para placeholders (ex.: [ID], [Local])
# ============================================================

pattern = re.compile(r"\[([^\]]+)\]")


# ============================================================
# Funções utilitárias
# ============================================================

def is_flowable(obj):
    """Verifica se o objeto é um Flowable (usado no ReportLab)."""
    return hasattr(obj, "wrap") and hasattr(obj, "draw")


def _is_na(value):
    """Retorna True se o valor for None, NaN ou equivalente."""
    return value is None or (isinstance(value, float) and pd.isna(value))


def replace_placeholders(cell, mapping):
    """
    Substitui placeholders no texto usando o dicionário 'mapping'.

    - Se o valor mapeado for Flowable, mantém o Flowable.
    - Se o valor for NaN/None, substitui por string vazia.
    - Substitui placeholders embutidos no texto.
    """
    if is_flowable(cell):
        return cell

    # Trata valores não-string
    if not isinstance(cell, str):
        if _is_na(cell):
            return ""
        return cell

    matches = pattern.findall(cell)
    if not matches:
        return cell

    # Caso a célula seja exatamente um único placeholder
    if len(matches) == 1 and cell.strip() == f"[{matches[0]}]":
        value = mapping.get(matches[0], "")
        if is_flowable(value):
            return value
        if _is_na(value):
            return ""
        if isinstance(value, str) and value.lower() == "nan":
            return ""
        return str(value)

    # Substitui placeholders embutidos
    def _replace(match):
        key = match.group(1)
        value = mapping.get(key, "")
        if _is_na(value):
            return ""
        if isinstance(value, str) and value.lower() == "nan":
            return ""
        return str(value)

    return pattern.sub(_replace, cell)


# ============================================================
# Função principal: Montagem da tabela
# ============================================================

def build_tabela(variables, checkbox_img, tabela_idx):
    """
    Gera uma tabela de Folha de Tarefa a partir das variáveis e índice informado.
    Retorna um objeto Table pronto para renderização no PDF.
    """

    # Valores para comparação de estilos especiais
    descricao_val = variables.get("Descrição", "")
    descricao_val = "" if _is_na(descricao_val) else str(descricao_val)

    passagem_val = variables.get("Passagem de Serviço", "")
    passagem_val = "" if _is_na(passagem_val) else str(passagem_val)

    # Layout base com placeholders
    dados = [
        (
            Paragraph("FRENTE:", style_subtitle), str(tabela_idx),
            Paragraph("Local:", style_subtitle), "[Local]",
            Paragraph("Cronograma", style_subtitle),
            Paragraph("Pendência", style_subtitle), " ",
            Paragraph("Atividade", style_subtitle), " ",
            Paragraph("INFORMAÇÕES DE PT", style_subtitle)
        ),
        ("[Descrição]", "", "", "", "", "Documentação", checkbox_img, "Nova", checkbox_img, "Nº PT:"),
        ("", "", "", "", "[Name]", "Projeto", checkbox_img, "Andamento:", checkbox_img, "Horário Solicitação:"),
        ("", "", "", "", "", "Suprimento", checkbox_img, "Paralisada:", checkbox_img, "Horário Liberação:"),
        ("Observações:", "", "[Passagem de Serviço]", "", "", "Outros", checkbox_img, "Finalizada:", checkbox_img, "Impacto de:"),
        ("Impacto ou Paralisação:", "", "", "", "", "", "", "", "", "ATÉ")
    ]

    # Substitui placeholders
    dados_substituidos = [
        [replace_placeholders(cel, variables) for cel in linha]
        for linha in dados
    ]

    # Aplica estilos aos textos
    dados_formatados = []
    for linha in dados_substituidos:
        nova_linha = []
        for cel in linha:
            if is_flowable(cel):
                nova_linha.append(cel)
                continue

            # Normaliza valor
            texto = "" if _is_na(cel) else str(cel).strip()
            if texto.lower() == "nan":
                texto = ""

            # Aplica estilo conforme conteúdo
            if descricao_val and texto == descricao_val.strip():
                nova_linha.append(Paragraph(texto, style_descricao))
            elif passagem_val and texto == passagem_val.strip():
                nova_linha.append(Paragraph(texto, style_passagem))
            else:
                nova_linha.append(Paragraph(texto, style_text))
        dados_formatados.append(nova_linha)

    # Define largura e altura das colunas
    col_widths = [2*cm, 1*cm, 1.5*cm, 3.5*cm, 4*cm, 3*cm, 1*cm, 3*cm, 1*cm, 6*cm]

    # Cria tabela final
    tabela = Table(
        dados_formatados,
        colWidths=col_widths,
        rowHeights=0.7*cm,
        hAlign="CENTER"
    )

    # Estilo da tabela
    tabela.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.5, colors.black),
        ("BOX", (0, 0), (-1, 0), 0.5, colors.black),
        ("BOX", (4, 0), (4, -2), 0.5, colors.black),
        ("BOX", (5, 0), (6, -2), 0.5, colors.black),
        ("BOX", (7, 0), (8, -2), 0.5, colors.black),
        ("BOX", (9, 0), (9, -2), 0.5, colors.black),
        ("BOX", (0, -1), (-1, -1), 0.5, colors.black),
        ("INNERGRID", (-1, 0), (-1, -2), 0.5, colors.black),
        ("BACKGROUND", (0, 0), (-1, 0), "#D9D9D9"),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("SPAN", (0, 1), (3, 1)), # Descrição
        ("SPAN", (0, -2), (1, -2)), # Observações
        ("SPAN", (2, -2), (3, -2)), #Passagem de Serviço
        ("SPAN", (0, -1), (1, -1)), #Impacto ou Paralisação
        ("SPAN", (0, -1), (4, -1)),
        ("SPAN", (-3, 0), (-2, 0)),
        ("SPAN", (-5, 0), (-4, 0))
    ]))

    return tabela
