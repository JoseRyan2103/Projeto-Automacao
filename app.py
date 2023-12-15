import openpyxl
import pyperclip
import pyautogui
from time import sleep

workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']

for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1285,295, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(1290,398, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(1294,566, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(1297,676, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    peso= linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(1286,781, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(1298,890, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pyautogui.click(1336,972, duration=1)
    sleep(3)

    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(1288,328, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    quant_em_estoque = linha[7].value
    pyperclip.copy(quant_em_estoque)
    pyautogui.click(1291,433, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(1299,544, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(1296,640, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    tamanho = linha[10].value
    pyautogui.click(1312,755, duration=1)
    if tamanho == 'Pqueno':
        pyautogui.click(1389,793, duration=1)
    elif tamanho == 'Médio':
        pyautogui.click(1372,822, duration=1)
    else:
        pyautogui.click(1375,848, duration=1)

    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(1302,863, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pyautogui.click(1326,942, duration=1)
    sleep(3)

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(1291,350, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(1292,458, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    obsevacoes = linha[14].value
    pyperclip.copy(obsevacoes)
    pyautogui.click(1294,562, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(1297,729, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(1297,836, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pyautogui.click(1320,907, duration=1)
    pyautogui.click(1803,239, duration=1)
    pyautogui.click(1802,239, duration=1)
    pyautogui.click(1578,631, duration=1)