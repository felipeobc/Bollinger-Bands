from openpyxl import Workbook

#acao = input("QUal o codigo da Ação que você quer processar: ").upper()

acao = "BIDI4"

with open(f'./dados/{acao}.txt', 'r') as arquivo_cotacao:
    linhas = arquivo_cotacao.readline()
    linhas = [linha.replace("\n", "").strip(";") for linha in linhas]


workbook = Workbook()
planilha_ativa = workbook.active
planilha_ativa.title = "Dados"

workbook.save("./saida/Planilha.xlsx")