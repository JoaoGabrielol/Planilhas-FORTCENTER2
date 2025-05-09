import streamlit as st
import pandas as pd
import environ
from dados_despesas import autenticar_msal, processar_arquivos
from utils_despesas import padronizar_e_limpar, filtrar_dados, calcular_metricas
import plotly.express as px
import re
from datetime import datetime

env = environ.Env()
environ.Env().read_env()

drive_id = env("drive_id")

arquivos = [
    {"nome": "Recebimentos_Caixa.xlsx", "caminho": "/Recebimentos%20Caixa%20(1).xlsx", "aba": "LANÇAMENTO DESPESAS", "linhas_pular": 3},
    {"nome": "PLANILHA_DE_CUSTO.xlsx", "caminho": "/PLANILHA%20DE%20CUSTO%202025.xlsx", "aba": "LANÇAMENTO DESPESAS", "linhas_pular": 3},
]

token = autenticar_msal()
headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

df = processar_arquivos(arquivos, drive_id, headers)
df = padronizar_e_limpar(df)

st.set_page_config(layout="wide", page_title="Análise Financeira de Despesas", page_icon="💸")

st.sidebar.header("Filtros")
opcoes_periodo = [
    "Semana Atual", "Semana Passada", "Mês Atual", "Mês Passado",
    "Últimos 3 Meses", "Últimos 6 Meses",
    "Ano Atual", "Ano Passado", "Tempo Todo"
]
periodo_selecionado = st.sidebar.selectbox("Selecione o período:", opcoes_periodo)

# Filtro de período
df_filtrado, inicio_periodo, fim_periodo = filtrar_dados(df, periodo_selecionado)

# Filtros dinâmicos (multi-select) para grupo, tipo, usuário
grupo_options = sorted(df_filtrado["GRUPO DESPESAS"].unique())
tipo_options = sorted(df_filtrado["TIPO DESPESAS"].unique())
usuario_options = sorted(df_filtrado["USUÁRIO"].unique())

grupo_selecionado = st.sidebar.multiselect("Selecione o(s) Grupo(s) de Despesas:", grupo_options)
tipo_selecionado = st.sidebar.multiselect("Selecione o(s) Tipo(s) de Despesa:", tipo_options)
usuario_selecionado = st.sidebar.multiselect("Selecione o(s) Usuário(s):", usuario_options)

min_valor = float(df_filtrado["VALOR R$"].min())
max_valor = float(df_filtrado["VALOR R$"].max())
intervalo_valor = st.sidebar.slider(
    "Selecione o intervalo de valores (R$):",
    min_value=min_valor, max_value=max_valor,
    value=(min_valor, max_valor), step=1.0
)

# aplica os filtros ao DataFrame
df_exibicao = df_filtrado.copy()
if grupo_selecionado:
    df_exibicao = df_exibicao[df_exibicao["GRUPO DESPESAS"].isin(grupo_selecionado)]
if tipo_selecionado:
    df_exibicao = df_exibicao[df_exibicao["TIPO DESPESAS"].isin(tipo_selecionado)]
if usuario_selecionado:
    df_exibicao = df_exibicao[df_exibicao["USUÁRIO"].isin(usuario_selecionado)]
df_exibicao = df_exibicao[(df_exibicao["VALOR R$"] >= intervalo_valor[0]) & (df_exibicao["VALOR R$"] <= intervalo_valor[1])]

df_exibicao_formatado = df_exibicao.copy()

df_exibicao_formatado["DATA"] = df_exibicao_formatado["DATA"].dt.strftime("%d/%m/%Y")

def limpar_descricao(valor):
    # Se for pandas Timestamp ou datetime.datetime (do Python)
    if isinstance(valor, pd.Timestamp) or isinstance(valor, datetime):
        return valor.strftime("%d/%m/%Y")
    # Se for string que parece data, tenta converter
    if isinstance(valor, str):
        try:
            data = pd.to_datetime(valor, format='%Y-%m-%d %H:%M:%S', errors='raise')
            return data.strftime("%d/%m/%Y")
        except:
            return valor
    return str(valor)

df_exibicao_formatado.loc[:, "DESCRIÇÃO DESPESA"] = df_exibicao_formatado["DESCRIÇÃO DESPESA"].apply(limpar_descricao)

df_exibicao_formatado["VALOR R$"] = df_exibicao_formatado["VALOR R$"].apply(
    lambda x: f"R$ {x:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")
)

df_exibicao_formatado = df_exibicao_formatado.rename(columns={
    "DATA": "Data",
    "GRUPO DESPESAS": "Grupo Despesa",
    "TIPO DESPESAS": "Tipo Despesa",
    "USUÁRIO": "Usuário",
    "DESCRIÇÃO DESPESA": "Descrição Despesa",
    "VALOR R$": "Valor R$"
})
df_exibicao_formatado.reset_index(drop=True, inplace=True)
df_exibicao_formatado.index += 1

# --- INTERFACE PRINCIPAL ---
st.markdown(
    "<h1 style='text-align:center; font-size:36px;'>📊 Análise Financeira de Despesas</h1>",
    unsafe_allow_html=True
)

if inicio_periodo and fim_periodo:
    st.write(f"**Período selecionado:** {inicio_periodo.strftime('%d/%m/%Y')} a {fim_periodo.strftime('%d/%m/%Y')}")

# tabela de dados filtrados
st.markdown(
    """
    <style>
        .css-1d391kg { font-size: 15px; }
        tbody tr td { text-align: left; font-size: 14px; }
    </style>
    """, unsafe_allow_html=True
)

st.dataframe(
    df_exibicao_formatado,
    use_container_width=True
)

st.markdown("## Métricas de Despesas")
total_despesas, media_despesas, quantidade_transacoes = calcular_metricas(df_exibicao)
col1, col2, col3 = st.columns(3)
col1.metric("💸 Total de Despesas", f"R$ {total_despesas:,.2f}")
col2.metric("🧾 Transações", quantidade_transacoes)
col3.metric("📈 Média por Transação", f"R$ {media_despesas:,.2f}")

st.markdown("## Despesas por Grupo")
grupo_despesas = df_filtrado.groupby("GRUPO DESPESAS")["VALOR R$"].sum().reset_index().sort_values(by="VALOR R$", ascending=False)
col4, col5 = st.columns(2)
with col4:
    st.dataframe(grupo_despesas)
with col5:
    fig_grupo = px.pie(grupo_despesas, names="GRUPO DESPESAS", values="VALOR R$", title="Despesas por Grupo")
    st.plotly_chart(fig_grupo, use_container_width=True)

st.markdown("## Despesas por Tipo")
tipo_despesas = df_filtrado.groupby("TIPO DESPESAS")["VALOR R$"].sum().reset_index().sort_values(by="VALOR R$", ascending=False)
col6, col7 = st.columns(2)
with col6:
    st.dataframe(tipo_despesas)
with col7:
    fig_tipo = px.bar(tipo_despesas, x="TIPO DESPESAS", y="VALOR R$", title="Despesas por Tipo", color="TIPO DESPESAS")
    st.plotly_chart(fig_tipo, use_container_width=True)

st.markdown("## Despesas por Usuário")
usuario_despesas = df_filtrado.copy()
usuario_despesas= usuario_despesas[usuario_despesas['USUÁRIO'] != "CORPORATIVO"]
usuario_despesas = usuario_despesas.groupby('USUÁRIO')['VALOR R$'].sum().reset_index() 
usuario_despesas = usuario_despesas.sort_values(by="VALOR R$", ascending=False)
col8, col9 = st.columns(2)
with col8:
    st.dataframe(usuario_despesas)
with col9:
    fig_usuario = px.bar(usuario_despesas, x="USUÁRIO", y="VALOR R$", title="Despesas por Usuário", color="USUÁRIO")
    st.plotly_chart(fig_usuario, use_container_width=True)

if st.checkbox("Mostrar detalhes por Grupo de Despesa para cada Usuário"):
    detalhado = df_exibicao.groupby(['USUÁRIO', 'GRUPO DESPESAS'])['VALOR R$'].sum().reset_index()
    # Tabela dinâmica
    tabela_pivot = detalhado.pivot(index='USUÁRIO', columns='GRUPO DESPESAS', values='VALOR R$').fillna(0)
    st.dataframe(tabela_pivot, use_container_width=True)
    
    # Gráfico de barras empilhadas (stacked bar)
    fig_stack = px.bar(
        detalhado,
        x='USUÁRIO',
        y='VALOR R$',
        color='GRUPO DESPESAS',
        title='Soma das Despesas por Grupo para cada Usuário',
        text_auto=True
    )
    st.plotly_chart(fig_stack, use_container_width=True)
