# grafico_tendencia.py

import plotly.express as px

def criar_grafico_tendencia(df, x, y, grupo, titulo, labels):
    """
    Cria um gráfico de tendência utilizando Plotly.
    """
    fig = px.line(
        df,
        x=x,
        y=y,
        color=grupo,
        title=titulo,
        labels=labels,
        category_orders={x: sorted(df[x].unique())}  # Ordenar eixo X
    )
    fig.update_layout(xaxis_title="Mês/Ano", yaxis_title=y)
    return fig
