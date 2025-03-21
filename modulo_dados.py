import json
import requests
from msal import ConfidentialClientApplication
import environ
import pandas as pd

def autenticar_msal():
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
        return result["access_token"]
    else:
        raise Exception("Nenhum token de acesso encontrado")

def baixar_arquivo(drive_id, file_name, file_path, headers):
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:{file_path}:/content"
    response = requests.get(url=url, headers=headers)
    if response.status_code == 200:
        with open(file_name, "wb") as f:
            f.write(response.content)
        print(f"{file_name} baixado com sucesso!")
    else:
        print(f"Erro ao acessar {file_name}: {response.status_code}, {response.text}")

def carregar_planilha(file_name, sheet_name, skiprows):
    if "Venda_Balcao" in file_name:
        df = pd.read_excel(file_name, header=0, usecols=['Dt. Neg.', 'Vlr. Nota'])
        df.rename(columns={'Dt. Neg.': 'DATA', 'Vlr. Nota': 'VALOR R$'}, inplace=True)
    else:
        df = pd.read_excel(file_name, sheet_name=sheet_name, skiprows=skiprows, header=0, 
                           usecols=['DATA', 'TÉCNICO', 'N° OS', 'OPERAÇÃO', 'PEÇAS',
                                    'M.O', 'VALOR R$', 'OBSERVAÇÃO'])
    return df

def processar_arquivos(arquivos, drive_id, headers):
    dfs = []
    for arquivo in arquivos:
        baixar_arquivo(drive_id, arquivo["nome"], arquivo["caminho"], headers)
        df = carregar_planilha(arquivo["nome"], arquivo["aba"], arquivo["linhas_pular"])
        dfs.append(df)
    return pd.concat(dfs, ignore_index=True)
