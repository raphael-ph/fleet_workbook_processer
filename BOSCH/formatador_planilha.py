import pandas as pd
import numpy as np

# ================================== FORMATAÇÃO DA PLANILHA ================================

# Planilhas que serão usadads para organizar a planilha formatada
#    * Na planilha FROTAS: substituir pelo nome da planilha de frotas a ser avaliada JÁ FORMATADA
#    * Na planilha PASSAGENS: substituir pelo nome da planilha de passagens JÁ FORMATADA
#    * Na planilha MODELO UNIFICADO: substituir pela planilha com os modelos unificados JÁ FORMATADA
#    * Na planilha MODELO REFERENCIA: substituir pela planilha com os modelos de referencia JÁ FORMATADA

planilha_de_frotas = 'FROTAS.xlsx'
planilha_de_passagens = 'PASSAGENS.xlsx'
planilha_modelo_unificado = 'MODELO_UNIFICADO.xlsx'
planilha_modelo_referencia = 'MODELO_REFERENCIA.xlsx'

# em FROTA_FORMATADA deve ser substituído o nome da planilha que deseja-se como output
planilha_formatada = 'FROTAS_FORMATADA.xlsx'
planilha_oficinas = 'OFICINAS_CONSOLIDADAS.xlsx'

df_frotas = pd.read_excel(planilha_de_frotas)
df_passagens = pd.read_excel(planilha_de_passagens)
df_modelo_unificado = pd.read_excel(planilha_modelo_unificado)
df_modelo_referencia = pd.read_excel(planilha_modelo_referencia)

df_3 = pd.merge(df_frotas, df_passagens[['PLACAS', 'NOME FANTASIA']], on='PLACAS', how='left')

df_3 = pd.merge(df_3, df_modelo_unificado[['PLACAS', 'MODELO UNIFICADO']], on='PLACAS', how='left')

df_3['CONCATENAR'] = df_3['MODELO UNIFICADO'] + " " + df_3['ANO MODELO'].astype(str)

df_3.rename(columns={'CONCATENAR':'MODELO UNIFICADO ANO'}, inplace=True)

df_3 = pd.merge(df_3, df_modelo_referencia[['MODELO UNIFICADO ANO', 'MODELO REFERENCIA']], on='MODELO UNIFICADO ANO', how='left')

df_3.rename(columns={'MODELO UNIFICADO ANO':'CONCATENAR'}, inplace=True)

df_3.rename(columns={'NOME FANTASIA':'OFICINA'}, inplace=True)

df_3 = df_3.drop_duplicates(subset='PLACAS')

df_formatado = df_3.sort_values('CIDADE AONDE ESTÁ O VEÍCULO')


# =============================== fim da formatação ================================

# =============================== organizar oficinas ===============================
'''
Nesta etapa, a ideia é organizar uma rotina que pegue a planilha pronta e preencha as mesmas
regiões com todas as oficinas que atendem
'''

for coluna in df_formatado[['CIDADE AONDE ESTÁ O VEÍCULO']]:
    s_cidade = df_formatado[coluna]

for coluna in df_formatado[['OFICINA']]:
    s_oficina = df_formatado[coluna]
    
df_cidades_oficina = pd.concat([s_cidade, s_oficina], axis=1)

dict_cidades_oficina = {k: g["OFICINA"].tolist() for k,g in df_cidades_oficina.groupby("CIDADE AONDE ESTÁ O VEÍCULO")}

df_cidades_oficina = pd.DataFrame({ key:pd.Series(value) for key, value in dict_cidades_oficina.items() })

df_cidades_oficina = df_cidades_oficina.transpose()

s_cidades_oficina = df_cidades_oficina.apply(lambda x: np.nan if x.isnull().all() else '/'.join(x.dropna()), axis=1)

df_cidades_oficina = s_cidades_oficina.to_frame()

df_cidades_oficina = df_cidades_oficina.replace(np.nan,'')

# ====================================== FINAL PROCESSING ============================

df_formatado.to_excel(planilha_formatada, index=False)
df_cidades_oficina.to_excel(planilha_oficinas)