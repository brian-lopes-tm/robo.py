# ############################################################################################################################
#                                                                                                                            #
#                                                                                                                            #   
#                               ROBÔ PARA MONITORAR OP. ASSISTIDA v1.4.0 powered by Leandro Freitas                          #
#                                                                                                                            # 
#                                                                                                                            #
# ############################################################################################################################

########################################################### CHANGE LOG #######################################################
# 
# 1.0.0 criação do programa
# 1.1.0 incluido o link da planilha gerada 
# 1.2.0 Incluido os totais de pedidos e faturamentos
# 1.3.0 incluido o farol para pedidos pendentes
# 1.4.0 incluido o farol e a informação de pedidos prestes a expirar 


########################################################### FIM CHANGE LOG #####################################################

import requests #biblioteca de chamada de URL´s
import json #biblioteca para pegar "payloads"
import html
import openpyxl
import pyodbc #biblioteca do banco SQLSERVER
from datetime import datetime
import locale
import calendar
import pandas as pd

# GERCRED248 - Unirmarka
# GERCRED270 - MNH
# GERCRED292 - InBahia
# GERCRED283 - Eldorado
# GERCRED290 - Dismax
# GERCRED294 - Paraty

#dismax = 15/11
#eldorado = 01/11
#MNH = 16/10
#Paraty = 01/11
#unimarka = 17/11

# webhook_url = "https://chat.googleapis.com/v1/spaces/AAAAcmax5yQ/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=RXH_MMXk6hD7HaFgIuSS0YJ9Nk6g816WtDQn1U3rSrI"

    # {
    #     "code": "GERCRED292", 
    #     webhook_url: "https://chat.googleapis.com/v1/spaces/AAAAcmax5yQ/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=RXH_MMXk6hD7HaFgIuSS0YJ9Nk6g816WtDQn1U3rSrI",
    #     name: "InBahia",
    #     sheets: "https://docs.google.com/spreadsheets/u/0/create?usp=sheets_web",
    #     operation_start_date: 
    # },

sellers = [
    {
        "code": "GERCRED248", 
        "webhook_url": "https://chat.googleapis.com/v1/spaces/AAAAcmax5yQ/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=RXH_MMXk6hD7HaFgIuSS0YJ9Nk6g816WtDQn1U3rSrI",
        "name": "Unimarka",
        "sheets": "https://docs.google.com/spreadsheets/d/18j7Vfx5gPdsnuEQ9uE3RuGc6xumC1Bc5IerUPuatzmg/edit#gid=0",
        "operation_start_date": "2023-11-17 00:00"
    },
    {
        "code": "GERCRED270", 
        "webhook_url": "https://chat.googleapis.com/v1/spaces/AAAAcmax5yQ/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=RXH_MMXk6hD7HaFgIuSS0YJ9Nk6g816WtDQn1U3rSrI",
        "name": "MNH",
        "sheets": "https://docs.google.com/spreadsheets/d/1hijwrCulCTj0DZSbTFdbg08MHZYb0XCUUTlcQjgJ23w/edit#gid=0",
        "operation_start_date": "2023-10-16 00:00"
    },
    {
        "code": "GERCRED283", 
        "webhook_url": "https://chat.googleapis.com/v1/spaces/AAAAcmax5yQ/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=RXH_MMXk6hD7HaFgIuSS0YJ9Nk6g816WtDQn1U3rSrI",
        "name": "Eldorado",
        "sheets": "https://docs.google.com/spreadsheets/d/11SuuPGLPVvojW6xmc4Qicmw5dLcrp_5BAWPEnrQXEPA/edit#gid=0",
        "operation_start_date": "2023-11-01 00:00"
    },
    {
        "code": "GERCRED290", 
        "webhook_url": "https://chat.googleapis.com/v1/spaces/AAAAcmax5yQ/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=RXH_MMXk6hD7HaFgIuSS0YJ9Nk6g816WtDQn1U3rSrI",
        "name": "Dismax",
        "sheets": "https://docs.google.com/spreadsheets/d/1k9XonBWI1HLSrVSlgORbTERDV4SWCFSmhluwDUvQh6o/edit#gid=0",
        "operation_start_date": "2023-11-15 00:00"
    },
    {
        "code": "GERCRED294", 
        "webhook_url": "https://chat.googleapis.com/v1/spaces/AAAAcmax5yQ/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=RXH_MMXk6hD7HaFgIuSS0YJ9Nk6g816WtDQn1U3rSrI",
        "name": "Paraty",
        "sheets": "https://docs.google.com/spreadsheets/d/18P-29NrRUKFafxzoJ_FyaeMr_eE8TiLHsgoCRCHHJ-A/edit#gid=0",
        "operation_start_date": "2023-11-01 00:00"
    }
]

#produção
server = "10.210.35.22,1433"
database = "TM_GATEWAY"
username = "svc-monitora-sustentacao"
password = "eSH7osQ3y6"


## Criando a string de conexão com o banco de dados SQL
conn_str = (
    f"DRIVER=ODBC Driver 17 for SQL Server;"
    f"SERVER={server};"
    f"DATABASE={database};"
    f"UID={username};"
    f"PWD={password}"
)

cnxn = pyodbc.connect(conn_str) #Este connect serve para o SQLSERVER
cnxn2 = pyodbc.connect(conn_str) #Este connect serve para o SQLSERVER
cnxn3 = pyodbc.connect(conn_str) #Este connect serve para o SQLSERVER
cnxn4 = pyodbc.connect(conn_str) #Este connect serve para o SQLSERVER


for sellerInfos in sellers:
    webhook_url = sellerInfos["webhook_url"]
    seller = sellerInfos["name"] # Nome que será exibido no relatorio via GCHAT
    NomePasta = sellerInfos["name"] # Nome da pasta criada no drive de relatorios
    AgreementName = sellerInfos["code"] 
    DataInicio = sellerInfos["operation_start_date"]  # data de inicio da operação assistida em produção
    planilha = sellerInfos["sheets"]

    ######################################################################## ///////////////////////////########################################################################

    cursor = cnxn.cursor() #Este cursor serve para o SQLSERVER

    # QUERY QUE REALIZA A CONULTA ENVIADA NO GCHAT
    print("Executando o select.....")
    # QUERY 1, responsavel por coletar a quantidade de pedidos pendentes
    result1 = """
    DECLARE @AgreementName VARCHAR(12) = '{}'
    DECLARE @DataInicio VARCHAR(16) = '{}'
    select 
    count(1) as 'Quantidade',
    case 
    when SUM(Amount) is null then 'Sem pedidos pendentes'
    else FORMAT(SUM(Amount) ,'c', 'pt-br') 
    end as "valor"
    from VW_Authorization va1 with (nolock) where AuthorizationId in 
    (
    select va.AuthorizationId from VW_Authorization va 
    where 1=1
    AND va.AgreementName ='GERCRED40'
    and va.SellerId = @AgreementName
    EXCEPT 
    select AuthorizationId from Invoices i with (nolock)
    where 1=1
    AND i.AgreementName ='GERCRED40'
    and i.ProcessResult  = 1
    and i.SellerName = @AgreementName
    )
    and va1.AuthorizationResult  = 'Approved'
    and va1.CreatedAt > @DataInicio
    """.format(AgreementName, DataInicio)
    cursor.execute(result1)


    cursor2 = cnxn3.cursor() #Este cursor serve para o SQLSERVER
    cursor3 = cnxn4.cursor() #Este cursor serve para o SQLSERVER

    # QUERY QUE REALIZA A CONULTA ENVIADA NO GCHAT

    # QUERY 2, responsavel por coletar a quantidade total de pedidos e notas fiscais

    result2 = """
    DECLARE @AgreementName VARCHAR(12) = '{}'
    DECLARE @DataInicio VARCHAR(16) = '{}'
        SELECT COUNT(va.AuthorizationId), 
        case 
    when SUM(Amount) is null then 'Sem pedidos'
    else FORMAT(SUM(Amount) ,'c', 'pt-br')
    end as 'Valor'
        FROM VW_Authorization va WITH (NOLOCK)
        WHERE CreditReasonId = 1
        AND va.AgreementName ='gercred40'
        AND va.SellerId = @AgreementName
        AND va.CreatedAt > @DataInicio 
        union all 
    SELECT COUNT(va.Id), 
    case 
    when SUM(Amount) is null then 'Sem Faturamentos '
    else FORMAT(SUM(Amount) ,'c', 'pt-br')   
    end as 'Valor'
        FROM Invoices VA WITH (NOLOCK)
        WHERE ProcessResult = 1
        AND va.sellerName = @AgreementName
        AND va.AgreementName ='gercred40'
        AND va.CreatedAt > @DataInicio
    """.format(AgreementName, DataInicio)
    cursor2.execute(result2)


    # QUERY 3, responsavel por coletar os dados para a lista detalhada
    sql_query = """

    DECLARE @AgreementName VARCHAR(12) = '{}'
    DECLARE @DataInicio VARCHAR(16) = '{}'
    SELECT 'Pedidos Não Recebidos' as 'Tipo',
        document as 'CNPJ Cliente',
        '-' as 'CNPJ Emissor',
        Format(amount, 'c', 'pt-br'),
        referencecode as 'Numero do Pedido',
        '-'                             AS 'Numero Nota fiscal',
        createdat as 'Data de criação',
        CONVERT(VARCHAR, expiresat, 20) AS 'DATA EXPIRAÇÃO'
    FROM   vw_authorization va1
    WHERE  authorizationid IN (SELECT va.authorizationid
                            FROM   vw_authorization va
                            WHERE  sellerid = @AgreementName
                            EXCEPT
                            SELECT authorizationid
                            FROM   invoices i WITH (nolock)
                            WHERE  i.sellername = @AgreementName
                                    AND i.processresult = 1)
        AND va1.authorizationresult = 'Approved'
        AND va1.agreementname = 'GERCRED40'
        AND va1.createdat > @DataInicio
    UNION ALL
    SELECT 'Pedidos'                       AS 'Tipo',
        document                        AS 'CNPJ Cliente',
        '-',
        Format(amount, 'c', 'pt-br')    AS 'Valor',
        referencecode                   AS 'Numero do Pedido',
        '-'                             AS 'Numero Nota fiscal',
        createdat                       AS 'Data de criação',
        CONVERT(VARCHAR, expiresat, 20) AS 'Data Expiração'
    FROM   vw_authorization va WITH (nolock)
    WHERE  creditreasonid = 1
        AND va.agreementname = 'gercred40'
        AND va.sellerid = @AgreementName
        AND va.createdat > @DataInicio
    UNION ALL
    SELECT 'Operações',
        i.receiverdocument             AS 'CNPJ Cliente',
        CONVERT(VARCHAR, i.IssuerDocument, 20)			  as 'CNPJ Emissor',
        Format(i.amount, 'c', 'pt-br') AS 'Valor',
        va2.referencecode,
        i.number                       AS 'Numero da Nota',
        i.createdat                    AS 'Data de criação',
        '-'
    FROM   invoices i WITH (nolock),
        vw_authorization va2
    WHERE  i.authorizationid = va2.authorizationid
        AND i.processresult = 1
        AND i.sellername = @AgreementName
        AND i.createdat > @DataInicio
        AND i.agreementname = 'gercred40' 
    """.format(AgreementName, DataInicio)

    # QUERY 1, responsavel por coletar a quantidade de pedidos pendentes
    resultPend1 = """
    DECLARE @AgreementName VARCHAR(12) = '{}'
    select 
    count(1) as 'Quantidade'
    from VW_Authorization va1 with (nolock) where AuthorizationId in 
    (
    select va.AuthorizationId from VW_Authorization va 
    where 1=1
    AND va.AgreementName ='GERCRED40'
    and va.SellerId = @AgreementName
    EXCEPT 
    select AuthorizationId from Invoices i with (nolock)
    where 1=1
    AND i.AgreementName ='GERCRED40'
    and i.ProcessResult  = 1
    and i.SellerName = @AgreementName
    )
    and va1.AuthorizationResult  = 'Approved'
    and ExpiresAt >= CONVERT (date, GETDATE())
    and ExpiresAt <= DATEADD(DAY, +3, GETDATE())
    """.format(AgreementName)
    cursor3.execute(resultPend1)

    # Executando a consulta SQL e armazenando os resultados em um DataFrame do pandas
    df = pd.read_sql(sql_query, cnxn2)


    # Formatando a data atual
    data_atual = date.today()
    data_formatada = data_atual.strftime("%d/%m/%Y")


    # Escrevendo o DataFrame em um arquivo Excel

    caminho = f"/home/brnlp/Desktop/tradeMaster/devOps/devops-scripts/scripts/relatorios/{NomePasta}/Relatorio_Pedidos.xlsx"
    df.to_excel(caminho, index=False)
    print("Arquivo salvo")

    # Formatar o resultado da consulta

    formatted_result = ""
    for row in cursor:
        formatted_result += f"<b>Pedidos pendentes de faturamento na Trademaster :</b> {str(row[0])}<br><b>Valor total de pedidos em aberto:</b> {str(row[1])}<br><br>"


    formatted_result3 = ""
    for row in cursor3:
        formatted_result3 += f"<b>Pedidos pendentes que vão expirar nos proximos 3 dias :</b> {str(row[0])}"

    # Definir as cores do farol
    cor_verde = "🟢"
    cor_amarelo = "🟡"
    cor_vermelho = "🔴"

    # Definir os limites para as cores do farol
    limite_verde = 0
    limite_amarelo = 3
    # qualquer coisa diferente é vermelho

    # Obter a quantidade da variável 'formatted_result'
    # quantidade = int(formatted_result.split(":")[1].split("<br>")[0])
    quantidade = int(formatted_result3.split(":")[1].split("<br>")[0].replace('</b> ', ''))

    # Escolher a cor do farol com base na quantidade
    if quantidade <= limite_verde:
        cor_farol = cor_verde
    elif quantidade <= limite_amarelo:
        cor_farol = cor_amarelo
    else:
        cor_farol = cor_vermelho

    # Formatar o resultado da consulta

    # formatted_result2 = ""
    # for row in cursor2:
    #     formatted_result += f"<b>Total de pedidos Recebidos :</b> {str(row[0][0])}<br><b>Valor total de pedidos Recebidos:</b> {str(row[0][1])}<br><br><b>Total de pedidos Faturados :</b> {str(row[1][0])}<br><b>Valor total de pedidos Faturados:</b> {str(row[1][1])}<br><br>"

    formatted_result2 = ""
    rows = list(cursor2)

    # Primeira linha da matriz
    formatted_result2 += f"<b>Total de pedidos Recebidos :</b> {str(rows[0][0])}<br><b>Valor total de pedidos Recebidos:</b> {str(rows[0][1])}<br><br>"

    # Segunda linha da matriz
    formatted_result2 += f"<b>Total de pedidos Faturados :</b> {str(rows[1][0])}<br><b>Valor total de pedidos Faturados:</b> {str(rows[1][1])}<br><br>"

    # Mensagem a ser enviada, neste caso utilizando HTML
    message = {
        "cards": [
            {
            "sections": [
                                {
                        "widgets": [
                            {
                                "image": {
                                    "imageUrl": "https://images2.imgbox.com/62/3b/QJezuElt_o.png"
                                }
                            },
                            {
                                "textParagraph": {
                                    "text": f'<b><font color="#00FA9A"><h3>Relatório de Pedidos - {seller}</h3></font>',
                                }
                            },
                        ]
                    },
                    {
                        "widgets": [
                            {
                                "textParagraph": {
                                    "text": f"Informação atualizada dos pedidos até o dia <b>{data_formatada}:"
                                }
                            }
                        ]
                    },
                    {
                        "widgets": [
                            {
                                "textParagraph": {
                                    "text": f"{formatted_result}" ,
                                }
                            }
                        ]
                    },
                                    {
                        "widgets": [
                            {
                                "textParagraph": {
                                    "text": f"{cor_farol} -  {formatted_result3}",
                                }
                            }
                        ]
                    },
                    {
                        "widgets": [
                            {
                                "textParagraph": {
                                    "text": formatted_result2,
                                }
                            }
                        ]
                    },
                    {
                        "widgets": [

                        ]
                    },
                    {
                        "widgets": [
                            {
                                "textParagraph": {
                    "text": f'Caso queira ter acesso ao relatorio detalhado dos pedidos, <a href="{planilha}">clique aqui</a>'
                }
                            }
                        ]

                    }
                ]
            }
        ]
    }

    # Enviar a mensagem para o espaço no Google Chat
    response = requests.post(webhook_url, json=message)

    if response.status_code == 200:
        print("Mensagem enviada com sucesso para o Google Chat.")
    else:
        print(f"Erro ao enviar mensagem para o Google Chat. Código de status: {response.status_code}")

    # Enviar mensagem de finalização para o Google Chat
    final_message = {

    }

    response = requests.post(webhook_url, json=final_message)

cursor3.close()
cursor2.close()
cursor.close()

# Fechando a conexão com o banco de dados
cnxn.close()
cnxn2.close()
cnxn3.close()
cnxn4.close()
