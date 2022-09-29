# fakturama -> link na descrição -> ERP de exemplo que vamos usar
# pyautogui -> é a ferramenta que a gente vai usar pra automatizar ERP

# Escrever a lógica que a gente vai construir em português o passo a passo

import pyautogui
import pyperclip
import subprocess
import time
import pandas as pd

pyautogui.FAILSAFE = True

def encontrar_imagem(imagem):
    while not pyautogui.locateOnScreen(imagem, grayscale=True, confidence=0.9):  # reconhecimento da imagem
        time.sleep(1)
    encontrou = pyautogui.locateOnScreen(imagem, grayscale=True, confidence=0.9)
    return encontrou


def escrever_texto(texto):
    pyperclip.copy(texto)
    pyautogui.hotkey('ctrl', 'v')

def direita(posicao_imagem):
    # posicao_imagem = (x, y, largura, altura)
    return posicao_imagem[0] + posicao_imagem[2], posicao_imagem[1] + posicao_imagem[3]/2

# Abrir o ERP (Fakturama)
subprocess.Popen([r"C:\Users\PC\Desktop\Fakturama\Fakturama2\Fakturama.exe"])
time.sleep(10)

# encontrou = None
# encontrou = x, y, largura, altura

encontrou = encontrar_imagem('fakturama.png')

# Fakturama está aberto

# Como fazer isto para vários produtos
tabela_produtos = pd.read_excel('Produtos.xlsx')

for linha in tabela_produtos.index:
    nome = tabela_produtos.loc[linha, 'Nome']
    id = tabela_produtos.loc[linha, 'ID']
    categoria = tabela_produtos.loc[linha, 'Categoria']
    gtin = tabela_produtos.loc[linha, 'GTIN']
    supplier = tabela_produtos.loc[linha, 'Supplier']
    descricao = tabela_produtos.loc[linha, 'Descrição']
    imagem = tabela_produtos.loc[linha, 'Imagem']
    preco = tabela_produtos.loc[linha, 'Preço']
    custo = tabela_produtos.loc[linha, 'Custo']
    estoque = tabela_produtos.loc[linha, 'Estoque']

    # Clicar no Menu New
    encontrou = encontrar_imagem('menu_new.png')
    pyautogui.click(pyautogui.center(encontrou))

    # Clicar no New Product
    encontrou = encontrar_imagem('new_product.png')
    pyautogui.click(pyautogui.center(encontrou))

    # Preencher todos os campos
    encontrou = encontrar_imagem('item_number.png')
    pyautogui.click(direita(encontrou))
    escrever_texto(str(id))

    encontrou = encontrar_imagem('name_produto.png')
    pyautogui.click(direita(encontrou))
    escrever_texto(str(nome))

    encontrou = encontrar_imagem('category_produto.png')
    pyautogui.click(direita(encontrou))
    escrever_texto(str(categoria))

    encontrou = encontrar_imagem('GTIN_produto.png')
    pyautogui.click(direita(encontrou))
    escrever_texto(str(gtin))

    encontrou = encontrar_imagem('supplier_produto.png')
    pyautogui.click(direita(encontrou))
    escrever_texto(str(supplier))

    encontrou = encontrar_imagem('description_produto.png')
    pyautogui.click(direita(encontrou))
    escrever_texto(str(descricao))

    encontrou = encontrar_imagem('preco_produto.png')
    pyautogui.click(direita(encontrou))
    preco_texto = f'{preco:.2f}'.replace('.',',')
    escrever_texto(str(preco_texto))

    encontrou = encontrar_imagem('custo_produto.png')
    pyautogui.click(direita(encontrou))
    custo_texto = f'{custo:.2f}'.replace('.',',')
    escrever_texto(str(custo_texto))

    encontrou = encontrar_imagem('stock_produto.png')
    pyautogui.click(direita(encontrou))
    estoque_texto = f'{estoque:.2f}'.replace('.',',')
    escrever_texto(str(estoque_texto))

    # Selecionar a imagem
    encontrou = encontrar_imagem('select_picture.png')
    pyautogui.click(pyautogui.center(encontrou))

    encontrou = encontrar_imagem('nome_arquivo.png')
    escrever_texto(rf'C:\Users\PC\Desktop\ImagensProdutos\{str(imagem)}')
    pyautogui.press('enter')

    # Clicar em salvar
    encontrou = encontrar_imagem('botao_save.png')
    pyautogui.click(pyautogui.center(encontrou))


    # fakturama está aberto
    # como fazer isso para vários produtos


