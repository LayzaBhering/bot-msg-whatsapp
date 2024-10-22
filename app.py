import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os

webbrowser.open('https://web.whatsapp.com/')

sleep(30)

import openpyxl.workbook

sheet = openpyxl.load_workbook('Data_Base_Py_Business.xlsx')
pageCustomers = sheet['Página1']

for linha in pageCustomers.iter_rows(min_row=2):
    nome = linha[0].value
    sobrenome = linha[1].value
    telefone = linha[2].value
    vencimento = linha[3].value

    message = f"Olá, {nome}! O seu boleto vence no dia {vencimento.strftime('%d/%m/%Y')}. Favor pagar pelo link: ..."

    try:
        linkStoreWhatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(message)}'
        sleep(20)
        webbrowser.open(linkStoreWhatsapp)
        sleep(20)
        seta = pyautogui.locateCenterOnScreen('bot-simbol.png')
        pyautogui.click(seta[0], seta[1])
        sleep(20)
        pyautogui.hotkey('ctrl', 'w')
        sleep(5)
    except:
        print(f'Não foi possível enviar uma mensagem para {nome}')
        with open('erros.csv', 'a', newline='',encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}{os.linesep}')
    
