import json
import requests
from msal import ConfidentialClientApplication
import streamlit as st
import pandas as pd
from unidecode import unidecode
import plotly.express as px
from datetime import datetime, timedelta
from io import BytesIO
import environ
from grafico_tendencia import criar_grafico_tendencia

env = environ.Env()
environ.Env().read_env()

client_id = env("id_do_cliente")
client_secret = env("segredo")
tenant_id = env("tenant_id")
msal_authority = f"https://login.microsoftonline.com/{tenant_id}"
msal_scope = ["https://graph.microsoft.com/.default"]

msal_app = ConfidentialClientApplication(
    client_id=client_id,
    client_credential=client_secret,
    authority=msal_authority,
)

result = msal_app.acquire_token_silent(scopes=msal_scope, account=None)
if not result:
    result = msal_app.acquire_token_for_client(scopes=msal_scope)
if "access_token" in result:
    access_token = result["access_token"]
else:
    raise Exception("No Access Token found")

headers = {
    "Authorization": f"Bearer {access_token}",
    "Content-Type": "application/json",
}

drive_id = env("drive_id")
file_paths = {
    "P._conta_atualizado.xlsx": "/P.conta%202024%20atualizado%20(3).xlsx",
    "Recebimentos_Caixa.xlsx": "/Recebimentos%20Caixa%20(1).xlsx"
}

def download_file(file_name, file_path):
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:{file_path}:/content"
    response = requests.get(url=url, headers=headers)
    if response.status_code == 200:
        with open(file_name, "wb") as f:
            f.write(response.content)
        print(f"{file_name} baixado com sucesso!")
    else:
        print(f"Erro ao acessar {file_name}: {response.status_code}, {response.text}")

for file_name, path in file_paths.items():
    download_file(file_name, path)

planilha_1 = pd.read_excel(
    'Recebimentos_Caixa.xlsx', sheet_name='ENTRADAS', skiprows=4, header=0
)
planilha_2 = pd.read_excel(
    'P._conta_atualizado.xlsx', sheet_name='Prestação', skiprows=5, header=0
)

def padronizar_colunas(df):
    mapeamento = {
        'DATA': 'data',
        'TÉCNICO': 'tecnico',
        'N° OS': 'no os',
        'OPERAÇÃO': 'operacao',
        'TIPO PAG.': 'tipo de pagamento',
        'PEÇAS': 'pecas',
        'M.O': 'mao de obra',
        'VALOR R$': 'valor r$',
        'OBSERVAÇÃO': 'observacao',
        'OUTROS': 'outros',
        'TOTAL C/TX': 'total com taxa',
    }
    df = df.rename(columns=mapeamento)
    return df

planilha_1 = padronizar_colunas(planilha_1)
planilha_2 = padronizar_colunas(planilha_2)

colunas_planilha_1 = ['data', 'tecnico', 'no os', 'operacao', 'tipo de pagamento', 'pecas', 'mao de obra', 'valor r$', 'observacao']
colunas_planilha_2 = ['data', 'tecnico', 'no os', 'operacao', 'tipo de pagamento', 'pecas', 'mao de obra', 'outros', 'valor r$', 'total com taxa']

planilha_1 = planilha_1[[col for col in colunas_planilha_1 if col in planilha_1.columns]]
planilha_2 = planilha_2[[col for col in colunas_planilha_2 if col in planilha_2.columns]]

df = pd.concat([planilha_1, planilha_2], axis=0, ignore_index=True, join="outer")

colunas_finais = [
    'data', 'tecnico', 'no os', 'operacao', 'tipo de pagamento', 'pecas',
    'mao de obra', 'valor r$','total com taxa', 'observacao', 'outros',
]
df = df.reindex(columns=colunas_finais)

colunas_relevantes = [
    'data', 'tecnico', 'no os', 'operacao', 'tipo de pagamento', 'pecas',
    'mao de obra', 'observacao', 'outros', 'total com taxa', 'valor r$'
]

df = df.dropna(subset=colunas_relevantes, how='all')

df = df[~((df['valor r$'].isna()) | (df['valor r$'] == "Não Informado"))]

df[['mao de obra', 'pecas', 'valor r$']] = df[['mao de obra', 'pecas', 'valor r$']].fillna(0)

df[['tecnico', 'observacao']] = df[['tecnico', 'observacao']].fillna("Não Informado")

opcoes_periodo = [
    "Semana Atual",
    "Semana Passada",
    "Mês Atual",
    "Mês Passado",
    "Últimos 3 Meses",
    "Ano Atual",
    "Ano Passado",
    "Tempo Todo",
]

st.sidebar.title("Filtro de Período")
periodo_selecionado = st.sidebar.selectbox("Selecione o Período", opcoes_periodo)

df['data'] = pd.to_datetime(df['data'], format='%d/%m/%Y', errors='coerce').dt.normalize()

hoje = pd.Timestamp.today().normalize()

st.sidebar.subheader("Período Selecionado")

if periodo_selecionado == "Semana Atual":
    inicio_periodo = hoje - pd.Timedelta(days=hoje.weekday())
    fim_periodo = hoje
    df_filtrado = df[df['data'] >= inicio_periodo]
elif periodo_selecionado == "Semana Passada":
    inicio_semana_atual = hoje - pd.Timedelta(days=hoje.weekday())
    inicio_periodo = inicio_semana_atual - pd.Timedelta(weeks=1)
    fim_periodo = inicio_semana_atual - pd.Timedelta(days=1)
    df_filtrado = df[(df['data'] >= inicio_periodo) & (df['data'] <= fim_periodo)]
elif periodo_selecionado == "Mês Atual":
    inicio_periodo = hoje.replace(day=1)
    fim_periodo = hoje
    df_filtrado = df[df['data'].dt.month == hoje.month]
elif periodo_selecionado == "Mês Passado":
    inicio_periodo = (hoje.replace(day=1) - pd.Timedelta(days=1)).replace(day=1)
    fim_periodo = inicio_periodo + pd.DateOffset(months=1) - pd.Timedelta(days=1)
    df_filtrado = df[(df['data'] >= inicio_periodo) & (df['data'] <= fim_periodo)]
elif periodo_selecionado == "Últimos 3 Meses":
    inicio_periodo = hoje - pd.DateOffset(months=3)
    fim_periodo = hoje
    df_filtrado = df[df['data'] >= inicio_periodo]
elif periodo_selecionado == "Ano Atual":
    inicio_periodo = hoje.replace(month=1, day=1)
    fim_periodo = hoje
    df_filtrado = df[df['data'].dt.year == hoje.year]
elif periodo_selecionado == "Ano Passado":
    inicio_periodo = hoje.replace(year=hoje.year - 1, month=1, day=1)
    fim_periodo = inicio_periodo.replace(year=inicio_periodo.year, month=12, day=31)
    df_filtrado = df[(df['data'] >= inicio_periodo) & (df['data'] <= fim_periodo)]
else:  
    inicio_periodo = df['data'].min()
    fim_periodo = df['data'].max()
    df_filtrado = df

st.sidebar.write(f"**Início do Período:** {inicio_periodo.strftime('%d/%m/%Y')}")
st.sidebar.write(f"**Fim do Período:** {fim_periodo.strftime('%d/%m/%Y')}")

ticket_medio = df_filtrado.groupby('tecnico')['valor r$'].mean().reset_index()
ticket_medio.rename(columns={'valor r$': 'ticket médio'}, inplace=True)
ticket_medio = ticket_medio.sort_values(by='ticket médio', ascending=False)

receita_total = df_filtrado.groupby('tecnico')['valor r$'].sum().reset_index()
receita_total.rename(columns={'valor r$': 'receita total'}, inplace=True)
receita_total = receita_total.sort_values(by='receita total', ascending=False)
    
colunas_numericas = ['mao de obra', 'pecas', 'valor r$']

for coluna in colunas_numericas:
    df_filtrado[coluna] = pd.to_numeric(df_filtrado[coluna], errors='coerce').fillna(0)

receita_mao_de_obra = df_filtrado.groupby('tecnico')['mao de obra'].sum().reset_index()
receita_mao_de_obra.rename(columns={'mao de obra': 'receita mão de obra'}, inplace=True)
receita_mao_de_obra = receita_mao_de_obra.sort_values(by='receita mão de obra', ascending=False)

receita_pecas = df_filtrado.groupby('tecnico')['pecas'].sum().reset_index()
receita_pecas.rename(columns={'pecas': 'receita peças'}, inplace=True)
receita_pecas = receita_pecas.sort_values(by='receita peças', ascending=False)

st.title("Análise de Dados por Técnico")

df['mes_ano'] = df['data'].dt.to_period('M').astype(str) 

df_mensal = df.groupby(['mes_ano', 'tecnico'])['valor r$'].sum().reset_index()

tecnicos_unicos = df['tecnico'].unique()

num_tecnicos_por_pagina = 3
pagina_selecionada = 1

if "mostrar_configuracoes" not in st.session_state:
    st.session_state["mostrar_configuracoes"] = False

if st.sidebar.button("Configurações Dinâmicas"):
    st.session_state["mostrar_configuracoes"] = not st.session_state["mostrar_configuracoes"]

if st.session_state["mostrar_configuracoes"]:
    st.sidebar.subheader("Configuração de Filtros Dinâmicos")

    num_tecnicos_por_pagina = st.sidebar.slider(
        "Número de Técnicos por Página",
        min_value=1,
        max_value=10,  
        value=num_tecnicos_por_pagina 
    )

    total_paginas = (len(tecnicos_unicos) // num_tecnicos_por_pagina) + (1 if len(tecnicos_unicos) % num_tecnicos_por_pagina > 0 else 0)

    pagina_selecionada = st.sidebar.slider(
        "Selecione a Página",
        min_value=1,
        max_value=total_paginas,
        value=pagina_selecionada 
    )

    st.sidebar.write(f"Página Selecionada: {pagina_selecionada}")
    st.sidebar.write(f"Número de Técnicos por Página: {num_tecnicos_por_pagina}")
else:
    total_paginas = (len(tecnicos_unicos) // num_tecnicos_por_pagina) + (1 if len(tecnicos_unicos) % num_tecnicos_por_pagina > 0 else 0)

inicio = (pagina_selecionada - 1) * num_tecnicos_por_pagina
fim = inicio + num_tecnicos_por_pagina
tecnicos_filtrados = tecnicos_unicos[inicio:fim]

df_pagina = df_mensal[df_mensal['tecnico'].isin(tecnicos_filtrados)]

st.write(f"Exibindo técnicos: {', '.join(tecnicos_filtrados)}")

inicio = (pagina_selecionada - 1) * num_tecnicos_por_pagina
fim = inicio + num_tecnicos_por_pagina
tecnicos_filtrados = tecnicos_unicos[inicio:fim]

df['data'] = pd.to_datetime(df['data'], format='%d/%m/%Y', errors='coerce')
df = df[df['data'].notna() & (df['data'] >= '2000-01-01') & (df['data'] <= '2100-12-31')]
df['mes_ano'] = df['data'].dt.strftime('%m/%Y') 

colunas_numericas = ['mao de obra', 'pecas', 'valor r$']

for coluna in colunas_numericas:
    df[coluna] = pd.to_numeric(df[coluna], errors='coerce').fillna(0)

ticket_medio_tendencia = df.groupby(['mes_ano', 'tecnico'])['valor r$'].mean().reset_index()
ticket_medio_tendencia.rename(columns={'valor r$': 'ticket médio'}, inplace=True)
ticket_medio_tendencia = ticket_medio_tendencia[ticket_medio_tendencia['tecnico'].isin(tecnicos_filtrados)]

receita_total_tendencia = df.groupby(['mes_ano', 'tecnico'])['valor r$'].mean().reset_index()
receita_total_tendencia.rename(columns={'valor r$': 'receita total'}, inplace=True)
receita_total_tendencia = receita_total_tendencia[receita_total_tendencia['tecnico'].isin(tecnicos_filtrados)]

receita_mao_de_obra_tendencia = df.groupby(['mes_ano', 'tecnico'])['mao de obra'].mean().reset_index()
receita_mao_de_obra_tendencia.rename(columns={'mao de obra': 'receita mão de obra'}, inplace=True)
receita_mao_de_obra_tendencia = receita_mao_de_obra_tendencia[receita_mao_de_obra_tendencia['tecnico'].isin(tecnicos_filtrados)]

receita_pecas_tendencia = df.groupby(['mes_ano', 'tecnico'])['pecas'].mean().reset_index()
receita_pecas_tendencia.rename(columns={'pecas': 'receita peças'}, inplace=True)
receita_pecas_tendencia = receita_pecas_tendencia[receita_pecas_tendencia['tecnico'].isin(tecnicos_filtrados)]

print(df['mes_ano'].unique())

st.subheader("Ticket Médio por Técnico")
col1, col2 = st.columns(2)
with col1:
    st.dataframe(ticket_medio)
with col2:
    fig_ticket_medio = px.bar(
        ticket_medio,
        x='tecnico',
        y='ticket médio',
        title="Ticket Médio por Técnico",
        color='tecnico',
        labels={'tecnico': 'Técnico', 'ticket médio': 'Ticket Médio (R$)'}
    )
    st.plotly_chart(fig_ticket_medio, use_container_width=True)

with col1:
    with st.expander("Tendência de Ticket Médio"):
        if st.checkbox("Mostrar gráfico de tendência - Ticket Médio", key="ticket_medio_tendencia"):
            fig_ticket_tendencia = criar_grafico_tendencia(
                ticket_medio_tendencia,
                'mes_ano',
                'ticket médio',
                'tecnico',
                "Tendência de Ticket Médio por Técnico",
                {'mes_ano': 'Mês/Ano', 'ticket médio': 'Ticket Médio (R$)'}
            )
            st.plotly_chart(fig_ticket_tendencia, use_container_width=True)

st.subheader("Receita Total por Técnico")
col3, col4 = st.columns(2)
with col3:
    st.dataframe(receita_total)
with col4:
    fig_receita_total = px.pie(
        receita_total,
        names='tecnico',
        values='receita total',
        title="Distribuição da Receita Total por Técnico",
    )
    st.plotly_chart(fig_receita_total, use_container_width=True)

with col3:
    with st.expander("Tendência de Receita Total"):
        if st.checkbox("Mostrar gráfico de tendência - Receita Total", key="receita_total_tendencia"):
            fig_receita_total_tendencia = criar_grafico_tendencia(
                receita_total_tendencia,
                'mes_ano',
                'receita total',
                'tecnico',
                "Tendência de Receita Total por Técnico",
                {'mes_ano': 'Mês/Ano', 'receita total': 'Receita Total (R$)'}
            )
            st.plotly_chart(fig_receita_total_tendencia, use_container_width=True)

st.subheader("Receita Mão de obra por Técnico")
col5, col6 = st.columns(2)
with col5:
    st.dataframe(receita_mao_de_obra)
with col6:
    fig_receita_mao_de_obra = px.pie(
        receita_mao_de_obra,
        names='tecnico',
        values='receita mão de obra',
        title="Distribuição da Receita Mão de Obra por Técnico",
    )
    st.plotly_chart(fig_receita_mao_de_obra, use_container_width=True)

with col5:
    with st.expander("Tendência de Receita de mão de obra"):
        if st.checkbox("Mostrar gráfico de tendência - Receita Mão de Obra", key="receita_mao_de_obra_tendencia"):
            fig_receita_mao_de_obra_tendencia = criar_grafico_tendencia(
                receita_mao_de_obra_tendencia,
                'mes_ano',
                'receita mão de obra',
                'tecnico',
                "Tendência de Receita de mão de obra por Técnico",
                {'mes_ano': 'Mês/Ano', 'receita mão de obra': 'Receita Mão de Obra (R$)'}
            )
            st.plotly_chart(fig_receita_mao_de_obra_tendencia, use_container_width=True)

st.subheader("Receita de Peças por Técnico")
col7, col8 = st.columns(2)
with col7:
    st.dataframe(receita_pecas)
with col8:
    fig_receita_pecas = px.pie(
        receita_pecas,
        names='tecnico',
        values='receita peças',
        title="Distribuição da Receita de Peças por Técnico",
    )
    st.plotly_chart(fig_receita_pecas, use_container_width=True)

with col7:
    with st.expander("Tendência de Receita de Peças"):
        if st.checkbox("Mostrar gráfico de tendência - Peças", key="receita_pecas_tendencia"):
            fig_receita_pecas_tendencia = criar_grafico_tendencia(
                receita_pecas_tendencia,
                'mes_ano',
                'receita peças',
                'tecnico',
                "Tendência de Receita de Peças por Técnico",
                {'mes_ano': 'Mês/Ano', 'receita peças': 'Receita de Peças (R$)'}
            )
            st.plotly_chart(fig_receita_pecas_tendencia, use_container_width=True)
