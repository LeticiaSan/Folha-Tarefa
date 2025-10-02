# ============================================================
# gerar_capa.py
# ------------------------------------------------------------
# Responsável por montar a capa da Folha de Tarefa.
# Recebe o nome do responsável, data e turno,
# e gera a tabela de colaboradores com seus respectivos campos.
# ============================================================
import os
import sys
import locale
import datetime
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import Image, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# Suporte a caracteres UTF-8 no terminal
sys.stdout.reconfigure(encoding="utf-8")

# Importa o dicionário de equipes externas
from equipes import EQUIPES

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


# ============================================================
# Função principal
# ============================================================

def gerar_capa(responsavel: str, data_escolhida: str, turno: str):
    """
    Monta os elementos da capa da Folha de Tarefa.

    Parâmetros:
        responsavel (str): Nome do responsável pela equipe.
        data_escolhida (str): Data da planilha (formato DD/MM/YYYY).
        turno (str): Turno selecionado (ex.: 'MANHÃ', 'NOITE').

    Retorna:
        list: Lista de elementos Flowable (tabela + espaçamento).
    """

    # ------------------------------------------------------------
    # Mapeia o nome do responsável para a lista de colaboradores
    # (Ignora maiúsculas/minúsculas)
    # ------------------------------------------------------------
    mapa_equipes = {k.lower(): v for k, v in EQUIPES.items()}
    colaboradores = mapa_equipes.get(responsavel.lower(), [])

    # ------------------------------------------------------------
    # Imagem do logotipo (ou texto substituto, se não encontrada)
    # ------------------------------------------------------------
    try:
        # Caminho absoluto da pasta atual (funções/)
        BASE_DIR = os.path.dirname(os.path.abspath(__file__))

        # Caminho para a pasta de imagens (sobe um nível e entra em 'imagens')
        IMG_DIR = os.path.join(BASE_DIR, "..", "imagens")

        # Caminho completo do logo
        logo_path = os.path.join(IMG_DIR, "logo.png")

        logo_img = Image(logo_path, width=3.5 * cm, height=2 * cm)

    except Exception as e:
        print(f"⚠️ Erro ao carregar logo: {e}")
        logo_img = Paragraph("LOGO", getSampleStyleSheet()["Normal"])

    # ------------------------------------------------------------
    # Estilos de texto
    # ------------------------------------------------------------
    styles = getSampleStyleSheet()

    style_text = styles["Normal"]
    style_text.fontSize = 7
    style_text.leading = 9
    style_text.alignment = 1

    style_title = ParagraphStyle(
        "Title",
        parent=style_text,
        fontSize=10,
        leading=12,
        alignment=1,
        textColor=colors.black,
        spaceAfter=4,
        fontName="Helvetica-Bold"
    )

    style_subtitle = ParagraphStyle(
        "Subtitle",
        parent=style_text,
        fontSize=7,
        leading=10,
        textColor=colors.black,
        fontName="Helvetica-Bold"
    )

    style_label = ParagraphStyle(
        "Label",
        parent=style_text,
        fontSize=8,
        leading=10,
        textColor=colors.black,
        fontName="Helvetica-Bold"
    )

    style_value = ParagraphStyle(
        "Value",
        parent=style_text,
        fontSize=7,
        leading=10,
        textColor="#0070C0",
        fontName="Helvetica-Bold",
        alignment=0
    )

    # ------------------------------------------------------------
    # Obtém o dia da semana (usando locale pt_BR)
    # ------------------------------------------------------------
    try:
        locale.setlocale(locale.LC_TIME, "pt_BR.UTF-8")
    except Exception:
        pass  # Ignora erro se o sistema não tiver suporte

    try:
        dia_semana = datetime.datetime.strptime(
            data_escolhida, "%d/%m/%Y"
        ).strftime("%A").capitalize()
    except Exception:
        dia_semana = ""

    # ------------------------------------------------------------
    # Cabeçalho da tabela
    # ------------------------------------------------------------
    dados = [
        (
            "", Paragraph("FOLHA TAREFA", style_title), "", "", "", "", "", "", "", "", "", "", "", "",
            Paragraph("CONTRATO:", style_label), "", Paragraph("5900.0126135.23.3", style_value)
        ),
        (
            logo_img, "", "", "", "", "", "", "", "", "", "", "", "", "",
            Paragraph("DATA:", style_label), "", Paragraph(data_escolhida, style_value)
        ),
        (
            "", "", "", "", "", "", "", "", "", "", "", "", "", "",
            Paragraph("DIA:", style_label), "", Paragraph(dia_semana, style_value)
        ),
        (
            Paragraph("EMPREENDIMENTO:", style_label), "", "REVAMP DA U-272D", "", "", "", "", "", "", "", "", "", "", "",
            Paragraph("TURNO:", style_label), "", Paragraph(turno, style_value)
        ),
        (
            Paragraph(f"(COLABORADOR) EQUIPE - {responsavel}", style_subtitle),
            Paragraph("FUNÇÃO", style_subtitle),
            Paragraph("FRENTE", style_subtitle), "", "", "", "", "", "", "", "", "", "", "",
            Paragraph("CONTROLE<br/>DE HORAS", style_subtitle), "", ""
        ),
        ("", "", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "", "", "")
    ]

    # ------------------------------------------------------------
    # Linhas dos colaboradores (máx. 17 linhas)
    # ------------------------------------------------------------
    linhas = []
    for i in range(17):
        if i < len(colaboradores):
            nome = colaboradores[i]["NOME"]
            funcao = colaboradores[i]["FUNÇÃO"]
        else:
            nome, funcao = "", ""

        linha = [
            Paragraph(nome, style_value),
            Paragraph(funcao, style_value)
        ] + [""] * 13 + [Paragraph("ATÉ", style_subtitle)]

        linhas.append(linha)

    dados.extend(linhas)

    # ------------------------------------------------------------
    # Configuração da tabela
    # ------------------------------------------------------------
    col_widths = [
        6 * cm, 4 * cm, 1 * cm, 1 * cm, 1 * cm, 1 * cm,
        1 * cm, 1 * cm, 1 * cm, 1 * cm, 1 * cm, 1 * cm,
        1 * cm, 1 * cm, 3 * cm, 1 * cm, 3 * cm
    ]

    tabela = Table(dados, colWidths=col_widths, rowHeights=0.8 * cm, hAlign="CENTER")

    # ------------------------------------------------------------
    # Estilo da tabela
    # ------------------------------------------------------------
    tabela.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.5, colors.black),
        ("BOX", (0, 0), (-1, 2), 0.5, colors.black),
        ("BOX", (0, 3), (13, 3), 0.5, colors.black),
        ("BOX", (-3, 3), (-1, 3), 0.5, colors.black),
        ("BOX", (-3, 0), (-1, 2), 0.5, colors.black),
        ("BOX", (-3, 4), (-1, 5), 0.5, colors.black),
        ("BACKGROUND", (0, 4), (-1, 5), "#D9D9D9"),
        ("INNERGRID", (0, 4), (-3, -1), 0.5, colors.black),
        ("INNERGRID", (-1, 6), (-1, -1), 0.5, colors.black),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),

        # Mesclagens de células
        ("SPAN", (1, 0), (13, 2)),
        ("SPAN", (0, 3), (1, 3)),
        ("SPAN", (2, 3), (13, 3)),
        ("SPAN", (2, 4), (13, 4)),
        ("SPAN", (0, 4), (0, 5)),
        ("SPAN", (1, 4), (1, 5)),
        ("SPAN", (-3, 4), (-1, 5)),
    ]))

    # ------------------------------------------------------------
    # Retorno final
    # ------------------------------------------------------------
    elementos = [tabela, Spacer(1, 12)]

    if colaboradores:
        print(f"✅ Capa montada para {responsavel} com {len(colaboradores)} colaboradores")
    else:
        print(f"⚠️ Responsável '{responsavel}' não encontrado no dicionário. Capa em branco montada.")

    return elementos
