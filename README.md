# 📊 Ferramenta de Análise Financeira de Receitas
  Este projeto fornece uma interface interativa para análise detalhada de receitas financeiras, auxiliando na gestão e visualização do desempenho financeiro da sua organização.

---

## 🚀 Principais Funcões

- **Filtros Dinâmicos:** Explore receitas por períodos personalizados, como semana atual, mês passado ou ano corrente.
- **Visualizações Gráficas:** Crie gráficos interativos para análise de receitas por técnico, categorias e tendências.
- **Tendências de Receita:** Identifique padrões de crescimento ou declínio ao longo do tempo.
- **Integração Automática:** Conexão com Microsoft Graph para download de planilhas armazenadas no OneDrive, utilizando autenticação segura OAuth2.
- **Tratamento Avançado de Dados:** Padronização de colunas, preenchimento de valores ausentes e limpeza de dados.

---

## 📋 Pré-requisitos

Certifique-se de ter as seguintes dependências instaladas no ambiente:

- **Python 3.8+**
- **Bibliotecas Necessárias**:
  - `pandas`
  - `streamlit`
  - `requests`
  - `msal`
  - `plotly`
  - `unidecode`
  - `environ`
  - `openpyxl`

Instale todas as dependências executando:
```bash
pip install pandas streamlit requests msal plotly unidecode environ openpyxl

---

## ⚙️ Configuração

### 1️⃣ Configurar Variáveis de Ambiente
Crie um arquivo `.env` no diretório do projeto com as seguintes informações:

```plaintext
id_do_cliente=<SEU_CLIENT_ID>
segredo=<SEU_CLIENT_SECRET>
tenant_id=<SEU_TENANT_ID>
drive_id=<SEU_DRIVE_ID>



