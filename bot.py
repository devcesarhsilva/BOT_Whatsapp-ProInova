import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui

# Inicio

#webbrowser.open('https://web.whatsapp.com/')
sleep(5)

workbook = openpyxl.load_workbook('lista.xlsx')
pagina_clientes = workbook['Sheet1']

for linha in pagina_clientes.iter_rows(min_row=2):
    #dados da planilha
    nome = linha[0].value
    telefone = linha[1].value
    interesse = linha[2].value
    preco = linha[3].value

filename_path = './imagens/gol 1.jpg'

mensagem = f'Olá {nome} estou entrando em contato referente ao anúncio no facebook, lembra? O seu interesse era o(a) {interesse} no valor de {preco}, ainda é de seu interesse?'

link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
webbrowser.open(link_mensagem_whatsapp)
sleep(10)

try:
    seta = pyautogui.locateCenterOnScreen('botao.png')
    sleep(3)
    pyautogui.click(seta[0],seta[1])
    sleep(3)
    pyautogui.hotkey('ctrl', 'w')
    sleep(3)
except:
    print(f'Não foi possível entrar em contato com {nome}')
    with open('erros.csv', 'a', newline='',encoding='utf-8') as arquivo:
        arquivo.write(f'{nome},{telefone}')

"""
 (\,,,/)
 (=";"=) 
Smelly Cat
  """