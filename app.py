"""
PRECISO AUTOMATIZAR MINHAS MENSAGENS P/ MEUS CLIENTES
"""
import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
# abrir o whatsapp
# try: 
#     web = webbrowser.open('https://web.whatsapp.com/')
#     print(web)
#     sleep(30)
# except Exception as erro:
#     print('erro: ', erro)

# # Ler planilha e guardar informações sobre nome, telefone e data de nascimento
workbook = openpyxl.load_workbook('Testepy.xlsx')
pagina_clientes = workbook['Página1']

for linha in pagina_clientes.iter_rows(min_row=2):
    #nome, telefone, vencimento
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value
    # print(nome)
    # print(type(telefone))
    print (type(vencimento))

    mensagem = f'Olá {nome} seu boleto vence no dia {vencimento} favor pagar no pix'
    print(mensagem)

    # criar links personalizados do whatsapp e enviar mensagem para cada cliente da planilha
    link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
    print(link_mensagem_whatsapp)
    webbrowser.open(link_mensagem_whatsapp)
    sleep(30)
