# https://cadastro-produtos-devaprender.netlify.app/

# python
# from mouseinfo import mouseInfo
# mouseInfo()

import openpyxl
import pyperclip
import pyautogui
from time import sleep

workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']

for linha in sheet_produtos.iter_rows(min_row=2):

  nome_produto = linha[0].value
  pyperclip.copy(nome_produto)
  pyautogui.click(1568,362,duration=1)
  pyautogui.hotkey('ctrl','v')

  descricao = linha[1].value
  pyperclip.copy(descricao)
  pyautogui.click(1552,462,duration=1)
  pyautogui.hotkey('ctrl','v')

  categoria = linha[2].value
  pyperclip.copy(categoria)
  pyautogui.click(1544,585,duration=1)
  pyautogui.hotkey('ctrl','v')

  codigo_produto = linha[3].value
  pyperclip.copy(codigo_produto)
  pyautogui.click(1531,667,duration=1)
  pyautogui.hotkey('ctrl','v')

  peso = linha[4].value
  pyperclip.copy(peso)
  pyautogui.click(1526,753,duration=1)
  pyautogui.hotkey('ctrl','v')

  dimensoes = linha[5].value
  pyperclip.copy(dimensoes)
  pyautogui.click(1521,839,duration=1)
  pyautogui.hotkey('ctrl','v')

  pyautogui.click(1414,902,duration=1)
  sleep(2)

  preco = linha[6].value
  pyperclip.copy(preco)
  pyautogui.click(1472,388,duration=1)
  pyautogui.hotkey('ctrl','v')

  estoque = linha[7].value
  pyperclip.copy(estoque)
  pyautogui.click(1463,476,duration=1)
  pyautogui.hotkey('ctrl','v')

  validade = linha[8].value
  pyperclip.copy(validade)
  pyautogui.click(1462,554,duration=1)
  pyautogui.hotkey('ctrl','v')

  cor = linha[9].value
  pyperclip.copy(cor)
  pyautogui.click(1461,652,duration=1)
  pyautogui.hotkey('ctrl','v')

  pyautogui.click(1479,731,duration=1)
  tamanho = linha[10].value
  
  if tamanho == 'Pequeno':
    pyautogui.click(1463,763,duration=1)
  elif tamanho == 'Médio':
    pyautogui.click(1463,785,duration=1)
  else:
    pyautogui.click(1463,808,duration=1)

  material = linha[11].value
  pyperclip.copy(material)
  pyautogui.click(1445,817,duration=1)
  pyautogui.hotkey('ctrl','v')

  pyautogui.click(1410,876,duration=1)
  sleep(2)

  fabricante = linha[12].value
  pyperclip.copy(fabricante)
  pyautogui.click(1504,407,duration=1)
  pyautogui.hotkey('ctrl','v')

  origem = linha[13].value
  pyperclip.copy(origem)
  pyautogui.click(1498,494,duration=1)
  pyautogui.hotkey('ctrl','v')

  observacoes = linha[14].value
  pyperclip.copy(observacoes)
  pyautogui.click(1492,587,duration=1)
  pyautogui.hotkey('ctrl','v')

  codigo_barras = linha[15].value
  pyperclip.copy(codigo_barras)
  pyautogui.click(1476,711,duration=1)
  pyautogui.hotkey('ctrl','v')

  localizacao = linha[16].value
  pyperclip.copy(localizacao)
  pyautogui.click(1445,801,duration=1)
  pyautogui.hotkey('ctrl','v')

  pyautogui.click(1410,860,duration=1)

  pyautogui.click(1803,180,duration=1)
  sleep(1)

  pyautogui.click(1640,631,duration=1)