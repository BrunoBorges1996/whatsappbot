import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui

#ler planilha e guardar informações, como nome, telefone e data de vencimento.
workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = workbook['Plan1']

for linha in pagina_clientes.iter_rows():
    #nome, telefone, data e vencimento
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value

    mensagem = f'Olá, {nome}! Você está recebendo uma mensagem escrita por um bot automatizado em Python, criado para mandar mensagens pelo WhatsApp. Ele procura seu contato numa lista do Excel, e manda um texto pré-programado.'
#criar links personalizados do whatsapp e enviar mensagens para cada cliente.
#https://web.whatsapp.com/send?phone=$text
#com base nos dados da planilha
    link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'

    webbrowser.open(link_mensagem_whatsapp)
    sleep(10)

    try:
        botao = pyautogui.locateCenterOnScreen('botao.png')
        pyautogui.click(botao[0], botao[1])
        sleep(5)
        pyautogui.hotkey('ctrl', 'w')
        sleep(5)
    except:
        print(f'Não foi possível enviar mensagem para {nome}.')
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome}, {telefone}')
