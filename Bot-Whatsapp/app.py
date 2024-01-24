import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui

webbrowser.open('https://web.whatsapp.com/')
sleep(30)

workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clients = workbook['Planilha1']

for linha in pagina_clients.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    
    mensagem = f'Ola {nome} to testando um sistema de bots, esse é seu numero => {telefone} ????. fique com um link de agrado :) https://doxbin.com/upload/lulapresidiario'

    try:
        link_msg_whats = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(limk_msg_whats)
        sleep(10)
        seta = pyautogui.locateCenterOnScreen('enviar.png')
        sleep(5)
        pyautogui.click(seta[0], seta[1])
        sleep(5)
        pyautogui.hotkey('ctrl', 'w')
        sleep(5)
    except:
        print(f'Não foi possivel enviar mensagem para {nome}, {telefone}')
        with open('erros.csv', 'a',newLine='', encoding='utf-8' ) as arquivo:
            arquivo.write(f' {nome}, {telefone}')