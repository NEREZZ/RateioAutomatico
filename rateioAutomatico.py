import sys
import os
import time
import pickle
import time
import random
from googleapiclient.errors import HttpError
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# Arquivo de credenciais baixado
CLIENT_SECRET_FILE = "client_secret_958231369671-mau6ihucotut1j7qn4cs2icnn39uug5o.apps.googleusercontent.com.json"  # Coloque o caminho do arquivo JSON de credenciais
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]  # Escopo de acesso


# Função para autenticação
def authenticate_google():
    creds = None
    # O arquivo token.pickle armazena o token de acesso do usuário
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)

    # Se não houver credenciais válidas, peça ao usuário para fazer login
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        # Salve as credenciais no arquivo token.pickle para reutilização
        with open("token.pickle", "wb") as token:
            pickle.dump(creds, token)
    return build("sheets", "v4", credentials=creds)


# Planilha que será enviada as colunas
novoRateio = "15Wv1JlR1anmYcT6-5RRPapPWApAEZAyKGgC2g80DFt8"
paginaNova = "Página1!A:E"

medicao = "1Oy220vBs3M7Y04CGHfT26IYQsQtHMnUT0RaIplN9SsM"
paginaMedicao = "medicao!K2:K"

rateioFinanceiro = "1ohl2iwOGbf1Y_cxrQsAhFDGeOPKU52sT6FqJeC35Sb4"
paginafinanceiro = "Rateio Chammas!B:F"
inventarioAtivos = "1P2-6GZli4hHx8KoViE13xj2Li81dzDr4dS9ObhD_tMY"
paginaAtivos = "Notebook!A:A"
service = authenticate_google()
sheet = service.spreadsheets()


def leituraDeDados():
    retorno = (
        sheet.values()
        .get(spreadsheetId=rateioFinanceiro, range=paginafinanceiro)
        .execute()
    )
    linhas = retorno.get("values", [])
    print("Dados lidos")
    # print(linhas)
    escreveDados(linhas)


# Escreva na outra planilha
def escreveDados(linhas):
    body = {"values": linhas}
    escrita = (
        sheet.values()
        .update(
            spreadsheetId=novoRateio,
            range=paginaNova,
            valueInputOption="RAW",
            body=body,
        )
        .execute()
    )
    print("Dados atualizados")


def realiza_calculo():

    # print(linhasMedicao)
    verificaDado()


def verificaDado():
    retornoMedicao = (
        sheet.values().get(spreadsheetId=medicao, range="Medicao!A2:A").execute()
    )
    linhasMedicao = retornoMedicao.get("values", [])
    retornoNovoRateio = (
        sheet.values().get(spreadsheetId=novoRateio, range="Página1!D2:D").execute()
    )
    linhasNovoRateio = retornoNovoRateio.get("values", [])
    retornoInventarioAtivos = (
        sheet.values().get(spreadsheetId=inventarioAtivos, range=paginaAtivos).execute()
    )
    linhasMatriculasAtivos = retornoInventarioAtivos.get("values", [])

    custoMedicao = (
        sheet.values().get(spreadsheetId=medicao, range="Medicao!K2:K").execute()
    )
    custo = custoMedicao.get("values", [])
    porcentagem = (
        sheet.values().get(spreadsheetId=novoRateio, range="Página1!E2:E").execute()
    )
    custoPorcentagem = porcentagem.get("values", [])
    perifericoMedicao = (
        sheet.values().get(spreadsheetId=medicao, range="Medicao!E2:E").execute()
    )
    periferico = perifericoMedicao.get("values", [])

    resultado = []

    # Converte listas de matrículas em conjuntos para busca rápida
    matriculasAtivos = {item[0] for item in linhasMatriculasAtivos if item}
    medicaoLista = {item[0] for item in linhasMedicao if item}

    # Substitui vírgula por ponto em custo e custoPorcentagem
    for i in range(len(custo)):
        if custo[i]:
            custo[i][0] = custo[i][0].replace(",", ".")

    for i in range(len(custoPorcentagem)):
        if custoPorcentagem[i]:
            custoPorcentagem[i][0] = custoPorcentagem[i][0].replace(",", ".")

    resultado = []

    # Converte listas de matrículas em conjuntos para busca rápida
    matriculasAtivos = {item[0] for item in linhasMatriculasAtivos if item}
    medicaoLista = list({item[0] for item in linhasMedicao if item})

    # Substitui vírgula por ponto em custo e custoPorcentagem
    for i in range(len(custo)):
        if custo[i]:
            custo[i][0] = custo[i][0].replace(",", ".")

    for i in range(len(custoPorcentagem)):
        if custoPorcentagem[i]:
            custoPorcentagem[i][0] = custoPorcentagem[i][0].replace(",", ".")

    j = 0
    dispositivo = []
    for i in range(len(linhasNovoRateio)):
        if len(linhasNovoRateio[i]) > 0:
            matricula = linhasNovoRateio[i][0]
        else:
            matricula = ""
        notebooks = list(
            {"HP", "Dell", "Lenovo", "Asus", "Acer", "MacBook", "Positivo", "Samsung"}
        )
        if matricula in matriculasAtivos and matricula in medicaoLista:

            try:
                # Encontrar a linha correspondente da matricula vinculada ao valor da medicao
                linha_medicao = next(
                    (
                        index
                        for index, item in enumerate(linhasMedicao)
                        if item[0] == matricula
                    ),
                    None,
                )

                if linha_medicao is not None and linha_medicao < len(custo):
                    valor_custo = custo[linha_medicao][0]
                    valor_porcentagem = custoPorcentagem[i][0]

                    multiplicador = float(valor_custo) * float(valor_porcentagem)
                    resultado.append([multiplicador])

                    if (
                        i + 1 < len(linhasNovoRateio)
                        and linhasNovoRateio[i + 1] != linhasNovoRateio[i]
                    ):

                        j += 1

                else:
                    resultado.append(["Erro: valor ausente"])
            except (ValueError, IndexError):
                resultado.append(["Erro ao calcular"])
        else:
            resultado.append(["Usuário sem periférico"])
            # verifica o tipo de dispositivo e adiciona a lista de dispositivos
        if any(marca.lower() in periferico[j][0].lower() for marca in notebooks):
            dispositivo.append(["Notebook"])
        else:
            dispositivo.append(["Tablet"])

    paginaNova = f"Página1!G2:G{len(resultado) + 1}"
    body = {"values": resultado}
    escrita = (
        sheet.values()
        .update(
            spreadsheetId=novoRateio,
            range=paginaNova,
            valueInputOption="RAW",
            body=body,
        )
        .execute()
    )
    paginaNova = f"Página1!F2:F{len(dispositivo) + 1}"
    body = {"values": dispositivo}
    escrita = (
        sheet.values()
        .update(
            spreadsheetId=novoRateio,
            range=paginaNova,
            valueInputOption="RAW",
            body=body,
        )
        .execute()
    )

    print("Dados atualizados")


#leituraDeDados()
realiza_calculo()
