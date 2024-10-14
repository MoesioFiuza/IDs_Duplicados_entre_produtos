import pandas as pd

novo_excel_file = r'C:\Users\moesios\Desktop\verificações\novo_produto_dez.xlsx'
excel_file = r'C:\Users\moesios\Desktop\verificações\unificado.xlsx'

def forçar_converter_para_string(df):
    for col in df.columns:
        if df[col].dtype != 'object':
            df[col] = df[col].astype(str)
    return df

def adicionar_coluna_duplicada(df, nome_coluna, novo_nome_coluna):
    df[novo_nome_coluna] = df.duplicated(subset=[nome_coluna])
    return df

def verificar_duplicatas_mesma_data(df, nome_coluna, novo_nome_coluna, coluna_data):
    duplicados = df.duplicated(subset=[nome_coluna], keep=False)
    mesma_data = df.groupby(nome_coluna)[coluna_data].transform('nunique') == 1
    df[novo_nome_coluna] = duplicados & mesma_data
    return df

dados_domicilio_df = pd.read_excel(excel_file, sheet_name='DADOS DO DOMICÍLIO', engine='openpyxl')

dados_domicilio_df = forçar_converter_para_string(dados_domicilio_df)
dados_domicilio_df = adicionar_coluna_duplicada(dados_domicilio_df, 'ID_DOMICILIO', 'Duplicatas_ID_DOMICILIO')
dados_domicilio_df = verificar_duplicatas_mesma_data(dados_domicilio_df, 'ID_DOMICILIO', 'Duplicatas_Mesma_Data', 'DATA DA PESQUISA')

arquivos = {
    "P06": r'C:\Users\moesios\Desktop\verificações\P6_BancoDados_Piloto_FORMATADA.XLSX',
    "P10": r'C:\Users\moesios\Desktop\verificações\P10 - ATUALIZADO.XLSX',
    "P07": r'C:\Users\moesios\Desktop\verificações\Produto 7 - Banco de Dados - OD Domiciliar - REV1E.XLSX',
    "P08": r'C:\Users\moesios\Desktop\verificações\Produto 8 - Banco de Dados - OD Domiciliar_rev04.XLSX',
    "P09": r'C:\Users\moesios\Desktop\verificações\Produto 9 - Banco de Dados - OD Domiciliar_rev03042024.XLSX'
}

for key in arquivos.keys():
    dados_domicilio_df[key] = ''

duplicados_por_arquivo = {key: [] for key in arquivos.keys()}
for key, file in arquivos.items():
    df_temp = pd.read_excel(file, engine='openpyxl')
    df_temp = forçar_converter_para_string(df_temp)
    ids_correspondentes = df_temp['ID_DOMICILIO'].isin(dados_domicilio_df['ID_DOMICILIO'])
    ids_correspondentes_unicos = df_temp[ids_correspondentes]['ID_DOMICILIO'].unique()
    dados_domicilio_df.loc[dados_domicilio_df['ID_DOMICILIO'].isin(ids_correspondentes_unicos), key] = dados_domicilio_df['ID_DOMICILIO']
    duplicados_por_arquivo[key] = ids_correspondentes_unicos

dados_domicilio_df['Duplicados_Arquivos'] = ''
for idx, row in dados_domicilio_df.iterrows():
    arquivos_duplicados = []
    for key in arquivos.keys():
        if row['ID_DOMICILIO'] in duplicados_por_arquivo[key]:
            arquivos_duplicados.append(key)
    dados_domicilio_df.at[idx, 'Duplicados_Arquivos'] = '-'.join(arquivos_duplicados)

with pd.ExcelWriter(novo_excel_file, engine='openpyxl') as writer:
    dados_domicilio_df.to_excel(writer, sheet_name='DADOS DO DOMICÍLIO', index=False)