import pandas as pd
import os

def processar_excel(caminho_arquivo: str, caminho_saida: str = None) -> str:
    """
    Processa um arquivo Excel:
      - Remove as duas primeiras linhas
      - Remove linhas sem valor na primeira coluna
      - Formata colunas "Hora Início" e "Hora fim" para HH:MM
      - Formata colunas "Cronograma - Start" e "Cronograma - End" para DD/MM/YYYY
      - Salva em um novo arquivo Excel

    Args:
        caminho_arquivo (str): Caminho completo do arquivo de entrada
        caminho_saida (str, opcional): Caminho do arquivo de saída. 
                                       Se não for informado, salva como 'saida_tratada.xlsx' 
                                       na pasta onde o script está sendo executado.

    Returns:
        str: Caminho do arquivo gerado
    """

    # Ler planilha SEM cabeçalho
    df = pd.read_excel(caminho_arquivo, header=None)

    # Remover as duas primeiras linhas
    df = df.drop(index=[0, 1]).reset_index(drop=True)

    # Definir a primeira linha restante como cabeçalho
    df.columns = df.iloc[0]
    df = df[1:].reset_index(drop=True)

    # Remover linhas sem valor na primeira coluna
    primeira_coluna = df.columns[0]
    df = df[df[primeira_coluna].notna()].reset_index(drop=True)

    # Formatar colunas de hora
    for col in ["Hora Início", "Hora Fim"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce").dt.strftime("%H:%M")

    # Formatar colunas de data
    for col in ["Cronograma - Start", "Cronograma - End"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce").dt.strftime("%d/%m/%Y")

    # Definir caminho de saída → sempre na pasta do script
    if not caminho_saida:
        pasta = os.getcwd()  # pega a pasta de execução do script
        caminho_saida = os.path.join(pasta, "saida_tratada.xlsx")

    # Salvar resultado
    df.to_excel(caminho_saida, index=False)

    return caminho_saida
