import pandas as pd

def abrir_planilha(caminho_arquivo):
    """
    Função para abrir uma planilha Excel usando pandas.
    
    Parâmetros:
    caminho_arquivo (str): O caminho para o arquivo Excel a ser aberto.
    
    Retorna:
    DataFrame: O conteúdo da planilha como um DataFrame pandas.
    """
    try:
        df = pd.read_excel(caminho_arquivo)
        return df
    except FileNotFoundError:
        print(f"O arquivo {caminho_arquivo} não foi encontrado.")
    except Exception as e:
        print(f"Ocorreu um erro ao abrir o arquivo: {e}")

# Exemplo de uso
caminho = r'C:\Users\Vitor\Downloads\Planilha ES (1).xlsx'
df_planilha = abrir_planilha(caminho)

if df_planilha is not None:
    print(df_planilha.head())  # Mostra as primeiras linhas do DataFrame
