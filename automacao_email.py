#Bibliotecas
import os

import pyautogui #automacao mouse e teclado
import time #controlar o tempo
import pyperclip #copiar e colar
import pandas as pd #ler planilhas

#verifica se o arquivo existe e exclui
if os.path.exists('Vendas - Dez.xlsx'):
    os.remove("Vendas - Dez.xlsx")

pyautogui.alert("Automação iniciada, pressione ok e nao mexa no computador")

link="https://drive.google.com/drive/folders/1wRTFw0sUVBjRr4hW5U9LF7DjLixRyxym"

#pausar a cada comando
pyautogui.PAUSE=1.8

#abrir o navegador
pyautogui.press('win')
pyautogui.write("firefox")
pyautogui.click(506, 250)


time.sleep(2)

#clicar na barra de navegação
pyautogui.click(x=1132, y=71)

#copiar link
pyperclip.copy(link)

#colar no navegador
pyautogui.hotkey('ctrl', 'v')
#pesquisa
pyautogui.press('enter')
time.sleep(3)
#clicar no arquivo
pyautogui.click(x=493, y=593)

#clicar nos tres pontos
pyautogui.click(x=1620, y=230)

#clicar em download
pyautogui.click(x=1306, y=790)

time.sleep(2)

#clicar na opcao salvar
pyautogui.click(x=812, y=839)

#clicar em ok
pyautogui.click(x=1215, y=959)

#clicar no caminho
pyautogui.click(x=492, y=85)

#"r" ignora comando pyton (\n)
caminho = r'C:\Users\Pichau\Documents\Projetos\Automacao_Email'
pyperclip.copy(caminho)
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')

pyautogui.click(x=687, y=657)

time.sleep(2)

#ler planilhas
tabelas = pd.read_excel('Vendas - Dez.xlsx')
faturamento = tabelas['Valor Final'].sum()
qtd = tabelas['Quantidade'].sum()

#abrir outra aba no navegador
pyautogui.hotkey('ctrl', 't')

link = 'https://mail.google.com/mail/u/0/#inbox'

pyperclip.copy(link)

#colar link no navegador

pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')

time.sleep(5)

pyautogui.click(x=74, y=250)

time.sleep(3)

pyautogui.write('alissonf.45+diretoria@gmail.com')
#email q sera enviado
corpo =f'''
Prezados, Bom Dia

O faturamento de ontem foi de: R${faturamento:,.2f}
a quantidade de produtos foi de: {qtd:,}

obg
Alisson
'''


#seleciona email
pyautogui.press('tab')

#pula pro campo
pyautogui.press('tab')

pyperclip.copy('Relatório Vendas')

#colar assunto
pyautogui.hotkey('ctrl', 'v')

pyautogui.press('tab')

#copiar email
pyperclip.copy(corpo)

#colar email
pyautogui.hotkey('ctrl', 'v')

#enviar email
pyautogui.hotkey('ctrl', 'enter')

pyautogui.alert("Fim da Automação")