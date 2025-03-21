import pandas as pd
from datetime import datetime, timedelta

def filtrar_dados(df, periodo):
    hoje = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)  
    
    df['DATA'] = pd.to_datetime(df['DATA']).dt.floor('D')  

    if periodo == "Semana Atual":
        data_inicio = hoje - timedelta(days=hoje.weekday())
        data_fim = data_inicio + timedelta(days=6)
    elif periodo == "Semana Passada":
        data_inicio = hoje - timedelta(days=hoje.weekday() + 7)
        data_fim = data_inicio + timedelta(days=6)
    elif periodo == "Mês Atual":
        data_inicio = hoje.replace(day=1)
        data_fim = hoje
    elif periodo == "Mês Passado":
        primeiro_dia_mes_passado = hoje.replace(day=1) - timedelta(days=1)
        data_inicio = primeiro_dia_mes_passado.replace(day=1)
        data_fim = primeiro_dia_mes_passado
    elif periodo == "Últimos 3 Meses":
        data_inicio = hoje.replace(day=1) - timedelta(days=89)
        data_fim = hoje
    elif periodo == "Últimos 6 Meses":
        data_inicio = hoje.replace(day=1) - timedelta(days=179)
        data_fim = hoje
    elif periodo == "Ano Atual":
        data_inicio = hoje.replace(month=1, day=1)
        data_fim = hoje
    elif periodo == "Ano Passado":
        data_inicio = (hoje.replace(month=1, day=1) - timedelta(days=1)).replace(day=1, month=1)
        data_fim = data_inicio.replace(month=12, day=31)
    else:
        return df, None, None
    df_filtrado = df[(df['DATA'] >= data_inicio) & (df['DATA'] <= data_fim)]
    
    return df_filtrado, data_inicio, data_fim

def calcular_metricas(df_filtrado):
    vendas_total = df_filtrado['VALOR R$'].sum()
    receita_media = df_filtrado['VALOR R$'].mean() if not df_filtrado.empty else 0
    quantidade_transacoes = len(df_filtrado)
    return vendas_total, receita_media, quantidade_transacoes
 