import pandas as pd
from unidecode import unidecode
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta

def padronizar_nome_usuario(nome):
    if pd.isnull(nome):
        return "NÃO INFORMADO"
    nome = str(nome).upper()
    nome = unidecode(nome)
    nome = nome.strip()
    nome = ' '.join(nome.split())  # <-- Isso resolve espaços múltiplos internos!
    return nome

def padronizar_e_limpar(df):
    for col in ["GRUPO DESPESAS", "TIPO DESPESAS", "USUÁRIO", "DESCRIÇÃO DESPESA"]:
        df[col] = df[col].fillna("Não informado")
    
    df["USUÁRIO"] = df["USUÁRIO"].apply(padronizar_nome_usuario)

    df["VALOR R$"] = pd.to_numeric(df["VALOR R$"], errors="coerce")
    df = df[df["VALOR R$"] > 0]
    df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce")
    df = df.dropna(subset=["DATA", "GRUPO DESPESAS", "TIPO DESPESAS", "USUÁRIO", "VALOR R$"])
    return df

def filtrar_dados(df, periodo):
    hoje = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)
    df = df.copy()
    data_inicio = None
    data_fim = None

    if periodo == "Semana Atual":
        data_inicio = hoje - timedelta(days=hoje.weekday())
        data_fim = hoje
    elif periodo == "Semana Passada":
        data_inicio = hoje - timedelta(days=hoje.weekday() + 7)
        data_fim = data_inicio + timedelta(days=6)
    elif periodo == "Mês Atual":
        data_inicio = hoje.replace(day=1)
        data_fim = hoje
    elif periodo == "Mês Passado":
        primeiro_dia_mes_passado = (hoje.replace(day=1) - timedelta(days=1)).replace(day=1)
        ultimo_dia_mes_passado = hoje.replace(day=1) - timedelta(days=1)
        data_inicio = primeiro_dia_mes_passado
        data_fim = ultimo_dia_mes_passado
    elif periodo == "Últimos 3 Meses":
        data_inicio = (hoje - relativedelta(months=3)).replace(day=1)
        data_fim = hoje
    elif periodo == "Últimos 6 Meses":
        data_inicio = (hoje - relativedelta(months=6)).replace(day=1)
        data_fim = hoje
    elif periodo == "Ano Atual":
        data_inicio = hoje.replace(month=1, day=1)
        data_fim = hoje
    elif periodo == "Ano Passado":
        data_inicio = hoje.replace(year=hoje.year - 1, month=1, day=1)
        data_fim = hoje.replace(year=hoje.year - 1, month=12, day=31)
    elif periodo == "Tempo Todo":
        return df, None, None
    else:
        return df, None, None
    
    if data_inicio and data_fim:
        df_filtrado = df[(df["DATA"] >= data_inicio) & (df["DATA"] <= data_fim)]
        return df_filtrado, data_inicio, data_fim
    else:
        return df, None, None

def calcular_metricas(df_filtrado):
    total = df_filtrado["VALOR R$"].sum()
    media = df_filtrado["VALOR R$"].mean() if not df_filtrado.empty else 0
    quantidade = len(df_filtrado)
    return total, media, quantidade
