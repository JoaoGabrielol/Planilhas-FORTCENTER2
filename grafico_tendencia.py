import plotly_express as px

def criar_grafico_tendencia(df, x, y, grupo, titulo, labels):
    df = df.sort_values(by=x)
    fig = px.line(
        df,
        x=x,
        y=y,
        color=grupo,
        title=titulo,
        labels=labels,
        markers=True
    )
   
    # Configura o eixo X pra ser categórico
    fig.update_layout(
        xaxis_title="Mês/Ano",
        yaxis_title=y,
        xaxis=dict(type="category")
    )
    return fig
