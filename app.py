import openpyxl
import pyperclip
import pyautogui
from time import sleep
# Entrar na planilha
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']
# Copiar informações de um campo e colar no seu campo correspondente
for linha in sheet_produtos.iter_rows(min_row=2):
    # PRIMEIRA PÁGINA
    nome_produto = linha[0].valueExtra
    pyperclip.copy(nome_produto)
    pyautogui.click(1110,338, duration=0.7)
    pyautogui.hotkey('ctrl','v')
  
    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(1116,425, duration=0.7)
    pyautogui.hotkey('ctrl','v')

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(1114,556, duration=0.7)
    pyautogui.hotkey('ctrl','v')

    cod_produto = linha[3].value
    pyperclip.copy(cod_produto)
    pyautogui.click(1114,643, duration=0.7)
    pyautogui.hotkey('ctrl','v')
    
    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(1110,728, duration=0.7)
    pyautogui.hotkey('ctrl','v')

    dimensao = linha[5].value
    pyperclip.copy(dimensao)
    pyautogui.click(1113,809, duration=0.7)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(1126,883, duration=0.7)
    sleep(2)

    # SEGUNDA PÁGINA
    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(1114,372, duration=0.7)
    pyautogui.hotkey('ctrl','v')

    quantidade_estoque = linha[7].value
    pyperclip.copy(quantidade_estoque)
    pyautogui.click(1109,461, duration=0.7)
    pyautogui.hotkey('ctrl','v')

    data_val = linha[8].value
    pyperclip.copy(data_val)
    pyautogui.click(1108,548, duration=0.7)
    pyautogui.hotkey('ctrl','v')

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(1109,632, duration=0.7)
    pyautogui.hotkey('ctrl','v')

    tamanho = linha[10].value
    pyautogui.click(1119,720, duration=0.7)
    if tamanho == 'Pequeno':
        pyautogui.click(1131,753, duration=0.7)
    elif tamanho == "Médio":
        pyautogui.click(1119,774, duration=0.7)
    else:
        pyautogui.click(1119,796, duration=0.7)

    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(1109,808, duration=0.7)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(1122,866, duration=0.7)
    sleep(2)

    # TERCEIRA PÁGINA
    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(1110,394, duration=0.7)
    pyautogui.hotkey('ctrl','v')

    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(1110,482, duration=0.7)
    pyautogui.hotkey('ctrl','v')

    observacao = linha[14].value
    pyperclip.copy(observacao)
    pyautogui.click(1114,572, duration=0.7)
    pyautogui.hotkey('ctrl','v')

    cod_barra = linha[15].value
    pyperclip.copy(cod_barra)
    pyautogui.click(1114,702, duration=0.7)
    pyautogui.hotkey('ctrl','v')

    localizacao = linha[16].value
    pyperclip.copy(localizacao)
    pyautogui.click(1114,785, duration=0.7)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(1129,850, duration=0.7)
    sleep(2)

    pyautogui.click(1609,188, duration=0.7)
    sleep(2)

    pyautogui.click(1430,609, duration=0.7)
    sleep(2)
# Repetir esses passos para outros campos até preencher campos daquela página
# Clicar em próxima 
# Repetir os mesmos passos e ir para a próxima página
# Repetir os mesmos passos e finalizar o cadastro daquele produto e clicar em concluir
# Clicar em OK, para finalizar o processo
# Clicar no OK mais uma vez na mensagem de confirmação de salvamento no banco de dados
# Clicar em "Adicionar mais um e repetir o processo até finalizar a planilha"