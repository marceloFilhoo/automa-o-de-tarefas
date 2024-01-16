import pandas as pd
import pyautogui
import time
import pyperclip




tabela = pd.read_excel("produtos_ficticios.xlsx")
print(tabela)
pyautogui.PAUSE = 0.3
pyautogui.press('win')
pyautogui.write('chrome')
pyautogui.press("enter")
pyautogui.write("https://cadastro-produtos-devaprender.netlify.app")
pyautogui.press("enter")


for linha in tabela.index:
    time.sleep(2) 
    pyautogui.click(x=426, y=337)
    #nome produto
    texto_acentuacao = tabela.loc[linha, "Nome do produto"]
    pyperclip.copy(texto_acentuacao)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('tab')
    #descrição
    texto_acentuacao = tabela.loc[linha, "Descrição"]
    pyperclip.copy(texto_acentuacao)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('tab')
    #categoria
    texto_acentuacao = tabela.loc[linha, "Categoria"]
    pyperclip.copy(texto_acentuacao)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('tab')
    #codigo do produto
    pyautogui.write(tabela.loc[linha, 'Código do produto'])
    pyautogui.press('tab')
    #peso
    pyautogui.write(str(tabela.loc[linha, 'Peso (em kg)']))
    pyautogui.press('tab')
    #dimensao
    pyautogui.write(str(tabela.loc[linha, 'Dimensões (L x A x P)']))
    pyautogui.press('tab')
    #enviar
    pyautogui.press('enter')
    #preço
    time.sleep(2) 
    pyautogui.press('tab')
    pyautogui.write(str(tabela.loc[linha, 'Preço']))
    pyautogui.press('tab')
    #quantidade estoque
    pyautogui.write(str(tabela.loc[linha, 'Quantidade em estoque']))
    pyautogui.press('tab')
    #data de validade
    texto_acentuacao = tabela.loc[linha, "Data de validade"]
    pyperclip.copy(texto_acentuacao)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('tab')
    #cor
    pyautogui.write(str(tabela.loc[linha, 'Cor']))
    pyautogui.press('tab')
    #tamanho
    pyautogui.write(str(tabela.loc[linha, 'Tamanho']))
    pyautogui.press('tab')
    #Material
    pyautogui.write(str(tabela.loc[linha, 'Material']))
    pyautogui.press('tab')
    #enviar
    pyautogui.press('enter')
    #fabricante
    time.sleep(2 ) 
    pyautogui.press('tab')
    pyautogui.write(str(tabela.loc[linha, 'Fabricante']))
    pyautogui.press('tab')
    #pais de origem
    texto_acentuacao = tabela.loc[linha, "País de origem"]
    pyperclip.copy(texto_acentuacao)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('tab')
    #obs
    texto_acentuacao = tabela.loc[linha, "Observações"]
    pyperclip.copy(texto_acentuacao)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('tab')
    #codigo de barra
    pyautogui.write(str(tabela.loc[linha, 'Código de barras']))
    pyautogui.press('tab')
    #loc armazem
    texto_acentuacao = tabela.loc[linha, "Localização no armazém"]
    pyperclip.copy(texto_acentuacao)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('tab')
    #enviar
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.click(x=1139, y=184)
    time.sleep(1)
    pyautogui.click(x=947, y=617)

    
