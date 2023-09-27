import pandas as pd
from datetime import datetime
import os


def dados():

    # Obter a data atual
    data = datetime.now().strftime("%d/%m/%Y, %H:%M:%S - ")

    # Carregar a planilha
    df = pd.read_excel('Chats Box - Export.xlsx')

    colunas = ['Nome da Fila']

    palavras = ['Tim', 'Ativos', 'Parcelas', 'Pefisa', 'Cac']

    contagem_palavras = {palavra: 0 for palavra in palavras}

    # Iterar sobre as linhas do DataFrame
    for index, row in df.iterrows():
        for coluna in colunas:
            texto = str(row[coluna])
            for palavra in palavras:
                if palavra.lower() in texto.lower():
                    contagem_palavras[palavra] += 1

    # Criar uma string com os resultados
    resultado = ""
    for palavra, contagem in contagem_palavras.items():
        resultado += f'{palavra}: {contagem}  '

    # Exibir a string com os resultados
    print(resultado)

    # Nome do arquivo com base na data
    nome_arquivo = "Quantidade_de_chats.txt"

    # Verifique se o arquivo já existe, se existir, apague-o
    if os.path.exists(nome_arquivo):
        os.remove(nome_arquivo)

    # Crie um novo arquivo e escreva nele
    with open("Quantidade_de_chats.txt", "a") as f:
        f.write('\n')
        f.write(str(data))
        f.write(f' Quantidade de chats para as respectivas carteiras e de {resultado}\n')

    print("Saldo gravado no arquivo txt.")

