# Importação de bibliotecas necessárias
import pandas as pd
from datetime import datetime

# Configuração para exibição de todas as colunas e linhas no DataFrame
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)

# Carregamento do arquivo Excel
bucha = pd.read_excel('Padrão.XLSX')

# Conversão da coluna 'Data' para o formato datetime
bucha['Data'] = pd.to_datetime(bucha['Data'], format='%d/%m/%Y')

# Função para processar os dados com base em critérios específicos
def processar_dados(dataframe, tipo_recebimento, observacao, col_renomeadas, arquivo_saida, sort_asc=True, is_group=False):
    """
    Função genérica para processar e exportar dados.

    Parameters:
        dataframe (DataFrame): O DataFrame original.
        tipo_recebimento (str): Tipo de recebimento a ser filtrado ('CRÉDITO' ou 'DÉBITO').
        observacao (str): Observação específica para filtragem ('FLAT', 'BONUS', etc.).
        col_renomeadas (dict): Mapeamento para renomear as colunas.
        arquivo_saida (str): Nome do arquivo Excel de saída.
        sort_asc (bool): Define se a ordenação será ascendente.
        is_group (bool): Define se os dados devem ser agrupados.

    Returns:
        DataFrame: DataFrame processado.
    """
    dados = dataframe[dataframe['Tipo de recebimento'].isin([tipo_recebimento])]
    dados = dados[dados['OBSERVAÇÃO'].isin([observacao])]

    if is_group:
        dados = dados.groupby(['Proposta'])['Valor da Comissao'].sum().reset_index()
        dados = dados.rename(columns=col_renomeadas)
    else:
        dados['Data'] = dados['Data'].dt.strftime('%d/%m/%Y')
        dados = dados.sort_values(by='Data', ascending=sort_asc)
        dados = dados.drop_duplicates(subset="Proposta")
        dados = dados.rename(columns=col_renomeadas)
        dados = dados.drop(['% Comissao', 'OBSERVAÇÃO', 'Valor base para calculo da comissao', 'Tipo de recebimento'], axis=1, errors='ignore')

    dados.to_excel(arquivo_saida, index=False)
    return dados

# Processamento de diferentes categorias
creditos_flat = processar_dados(bucha, 'CRÉDITO', 'FLAT',
                                {'Data': 'Data do repasse FLAT', 'Valor da Comissao': 'COMISSÃO FLAT'},
                                'TesteOrdem.xlsx')

flat_group = processar_dados(bucha, 'CRÉDITO', 'FLAT',
                             {'Valor da Comissao': 'PROVA REAL FLAT'},
                             'testeflat.xlsx', is_group=True)

creditos_bonus = processar_dados(bucha, 'CRÉDITO', 'BONUS',
                                 {'Data': 'DATA REPASSE BONUS', 'Valor da Comissao': 'Valor do BONUS'},
                                 'TesteOrdemBonus.xlsx')

bonus_group = processar_dados(bucha, 'CRÉDITO', 'BONUS',
                              {'Valor da Comissao': 'PROVA REAL BONUS'},
                              'testebonus.xlsx', is_group=True)

debitos_flat = processar_dados(bucha, 'DÉBITO', 'FLAT',
                               {'Data': 'DATA DÉBITO COMISSÃO FLAT', 'Valor da Comissao': 'DÉBITO COMISSÃO FLAT'},
                               'TesteOrdemDebito.xlsx')

# Merge de todos os DataFrames processados
dfs_to_merge = [creditos_flat, flat_group, creditos_bonus, bonus_group, debitos_flat]
merged_df = bucha[['Proposta']].drop_duplicates()

for df in dfs_to_merge:
    merged_df = merged_df.merge(df, on='Proposta', how='left')

# Adicionando soma total de comissão
comissao_total = bucha.groupby(['Proposta'])['Valor da Comissao'].sum().reset_index()
comissao_total = comissao_total.rename(columns={'Valor da Comissao': 'Comissão REAL BANCO'})
merged_df = merged_df.merge(comissao_total, on='Proposta', how='left')

# Preenchendo valores nulos com texto padrão
fill_values = {
    'Valor do BONUS': 'SEM BONUS',
    'DATA REPASSE BONUS': 'SEM BONUS',
    'DATA DÉBITO COMISSÃO FLAT': 'SEM DÉBITO',
    'Comissão REAL BANCO': 0
}
merged_df.fillna(fill_values, inplace=True)

# Reordenando as colunas para a exportação
columns_order = [
    'Proposta', 'Data do repasse FLAT', 'COMISSÃO FLAT', 'DATA REPASSE BONUS',
    'Valor do BONUS', 'DATA DÉBITO COMISSÃO FLAT', 'DÉBITO COMISSÃO FLAT',
    'Comissão REAL BANCO', 'PROVA REAL BONUS', 'PROVA REAL FLAT'
]
merged_df = merged_df[columns_order]

# Exportação final para Excel
merged_df.to_excel('procv.xlsx', index=False)
