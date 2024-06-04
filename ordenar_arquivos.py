from funcoes import FuncoesDataFrame, FuncoesNumericas
import pandas as pd


pastas = pd.read_excel('pastas.xlsx').convert_dtypes()

arquivos = pastas[pastas['Tipo de Item'] == 'Item']

# Agrupado por arquivos
arquivos_ordenados = arquivos.sort_values(
    by='Tamanho do Arquivo', ascending=False)

arquivos_ordenados['Tamanho do Arquivo'] = arquivos_ordenados['Tamanho do Arquivo'].apply(
    FuncoesNumericas.converter_byte)

FuncoesDataFrame.gera_excel(
    arquivos_ordenados, 'Arquivos_SharePoint_Por_Tamanho')
