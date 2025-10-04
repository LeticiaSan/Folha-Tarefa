import os
import sys
import pandas as pd
from datetime import datetime, time, timedelta
from tkinter import Tk, filedialog, Toplevel, Label, Button, StringVar, OptionMenu
from tkcalendar import DateEntry
from reportlab.platypus import SimpleDocTemplate, Spacer, Image, PageBreak
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import cm

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

IMG_DIR = os.path.join(BASE_DIR, "imagens")
SAIDAS_DIR = os.path.join(BASE_DIR, "folhatarefa")
FUNCOES_DIR = os.path.join(BASE_DIR, "funcoes")

if FUNCOES_DIR not in sys.path:
    sys.path.append(FUNCOES_DIR)

# Importações de módulos internos
from funcoes.layout import build_tabela
from funcoes.gerar_capa import gerar_capa
from funcoes.processar_planilha_monday import processar_excel


# ============================================================
# 📅 Função para selecionar data e turno
# ============================================================
def selecionar_data_turno():
    """Abre janela para o usuário escolher a data e o turno da folha-tarefa."""
    root = Tk()
    root.withdraw()

    escolha = {"data": None, "turno": None}

    def confirmar():
        escolha["data"] = date_entry.get_date().strftime("%d/%m/%Y")
        escolha["turno"] = turno_var.get()
        janela.destroy()

    janela = Toplevel(root)
    janela.title("Selecione a data e o turno da Folha-Tarefa")

    Label(janela, text="Escolha a data:").pack(pady=5)
    date_entry = DateEntry(janela, date_pattern="dd/MM/yyyy", locale="pt_BR")
    date_entry.pack(pady=5)

    Label(janela, text="Selecione o turno:").pack(pady=5)
    turno_var = StringVar(janela)
    turno_var.set("Manhã")  # valor padrão
    OptionMenu(janela, turno_var, "Manhã", "Noite").pack(pady=5)

    Button(janela, text="Confirmar", command=confirmar).pack(pady=10)

    janela.grab_set()
    root.wait_window(janela)

    return escolha


# ============================================================
# 📂 Seleção do arquivo Excel e pré-processamento
# ============================================================
Tk().withdraw()
excel_path = filedialog.askopenfilename(
    title="Selecione o arquivo Excel",
    filetypes=[("Excel files", "*.xlsx *.xls")]
)

if not excel_path:
    print("⚠️ Nenhum arquivo selecionado. Encerrando execução.")
    exit()

# Processa a planilha para garantir formatos padronizados
entrada = excel_path
saida = processar_excel(entrada)

# Carrega o arquivo processado
df = pd.read_excel(saida, dtype=str)


# ============================================================
# 🕓 Conversão de horas
# ============================================================
def converter_hora(valor):
    """Converte texto de hora no formato HH:MM para objeto time."""
    if pd.isna(valor) or str(valor).strip() == "":
        return None
    try:
        horas, minutos = map(int, str(valor).split(":"))
        return time(horas, minutos)
    except Exception:
        return None


for col in ["Hora Início", "Hora Fim"]:
    df[col] = df[col].apply(converter_hora)


# ============================================================
# 🧹 Tratamento de valores nulos
# ============================================================
for col in ["Encarregado Manhã", "Encarregado Noite", "Hora Início", "Hora Fim"]:
    if col in df.columns:
        df[col] = df[col].fillna("")


# ============================================================
# 🗓️ Seleção de data e turno
# ============================================================
escolha = selecionar_data_turno()
data_input = escolha["data"]
turno_escolhido = escolha["turno"].upper()
data_input_dt = datetime.strptime(data_input, "%d/%m/%Y").date()


# ============================================================
# 📆 Conversão de colunas de data
# ============================================================
df["Cronograma - Start"] = pd.to_datetime(
    df["Cronograma - Start"], errors="coerce", dayfirst=True
).dt.date

df["Cronograma - End"] = pd.to_datetime(
    df["Cronograma - End"], errors="coerce", dayfirst=True
).dt.date


# ============================================================
# ⏰ Definição das janelas de turno
# ============================================================
hora_manha_ini = time(8, 30)
hora_manha_fim = time(18, 0)
hora_noite_ini = time(19, 30)
hora_noite_fim = time(5, 0)  # madrugada do dia seguinte


# ============================================================
# 🧭 Função de filtro por turno
# ============================================================
def pertence_turno(row):
    """Retorna True se a atividade estiver dentro da janela do turno selecionado
       ou se o status for 'Atraso' ou 'Em andamento'."""
    
    status = str(row.get("Status", "")).strip().lower()

    # ✅ Regra extra: inclui sempre se o status for "atraso" ou "em andamento"
    if status in ["atraso", "em andamento"]:
        return True

    # Validação de dados obrigatórios para filtragem por horário
    if pd.isna(row["Cronograma - Start"]) or pd.isna(row["Cronograma - End"]):
        return False
    if pd.isna(row["Hora Início"]) or pd.isna(row["Hora Fim"]):
        return False

    # Combina data e hora em objetos datetime
    dt_inicio = datetime.combine(row["Cronograma - Start"], row["Hora Início"])
    dt_fim = datetime.combine(row["Cronograma - End"], row["Hora Fim"])

    # Corrige virada de dia
    if dt_fim <= dt_inicio:
        dt_fim += timedelta(days=1)

    # Define janela conforme turno escolhido
    if turno_escolhido == "MANHÃ":
        janela_ini = datetime.combine(data_input_dt, hora_manha_ini)
        janela_fim = datetime.combine(data_input_dt, hora_manha_fim)
    else:  # NOITE
        janela_ini = datetime.combine(data_input_dt, hora_noite_ini)
        janela_fim = datetime.combine(data_input_dt + timedelta(days=1), hora_noite_fim)

    # Retorna True se houver sobreposição entre atividade e janela
    return (dt_inicio < janela_fim) and (dt_fim > janela_ini)



# ============================================================
# 🔍 Aplicação do filtro de turno
# ============================================================
df_filtrado = df[df.apply(pertence_turno, axis=1)]

atividades_ignoradas = df[~df.index.isin(df_filtrado.index)]
for _, row in atividades_ignoradas.iterrows():
    descricao = row.get("Descrição", "Sem descrição")
    print(f"⚠️ Atividade ignorada (fora do turno): {descricao}")

df = df_filtrado


# ============================================================
# 🧱 Configuração de imagens e pastas de saída
# ============================================================
checkbox_img = Image(os.path.join(IMG_DIR, "square.png"), width=0.35 * cm, height=0.35 * cm)

pasta_nome = f"{data_input_dt.strftime('%d-%m-%Y')}_{turno_escolhido}"
output_dir = os.path.join(SAIDAS_DIR, f"Folhas-Tarefa {pasta_nome}")
os.makedirs(output_dir, exist_ok=True)


# ============================================================
# 👷 Geração das folhas por responsável
# ============================================================
col_responsavel = (
    "Encarregado Manhã" if turno_escolhido == "MANHÃ" else "Encarregado Noite"
)
responsaveis = sorted(set(df[col_responsavel].unique().tolist()))

for responsavel in responsaveis:
    if not responsavel.strip():
        continue

    df_responsavel = df[df[col_responsavel].str.lower() == responsavel.lower()]
    output_pdf = os.path.join(
        output_dir, f"Folha_Tarefa_{responsavel}_{pasta_nome}.pdf"
    )

    doc = SimpleDocTemplate(
        output_pdf,
        pagesize=landscape(A4),
        rightMargin=1 * cm,
        leftMargin=1 * cm,
        topMargin=1 * cm,
        bottomMargin=1 * cm,
    )

    elementos = []
    elementos += gerar_capa(responsavel, data_input, turno_escolhido)

    tables_per_page = 4
    tabela_idx = 1
    count = 0

    # Tabelas com atividades
    for _, row in df_responsavel.iterrows():
        tabela = build_tabela(row.to_dict(), checkbox_img, tabela_idx)
        elementos.append(tabela)
        elementos.append(Spacer(1, 1))
        tabela_idx += 1
        count += 1
        if count == tables_per_page:
            elementos.append(PageBreak())
            count = 0

    # 3 tabelas extras em branco
    for _ in range(3):
        tabela_vazia = {col: "" for col in df.columns}
        tabela = build_tabela(tabela_vazia, checkbox_img, tabela_idx)
        elementos.append(tabela)
        elementos.append(Spacer(1, 1))
        tabela_idx += 1
        count += 1
        if count == tables_per_page:
            elementos.append(PageBreak())
            count = 0

    doc.build(elementos)
    print(f"✅ Folha-tarefa gerada para {responsavel}: {output_pdf}")
