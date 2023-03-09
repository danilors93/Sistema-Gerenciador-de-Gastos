import openpyxl
import os

# Criando ou abrindo o arquivo excel
if os.path.exists("controle_gastos.xlsx"):
    workbook = openpyxl.load_workbook("controle_gastos.xlsx")
else:
    workbook = openpyxl.Workbook()
    workbook.active.title = "Gastos"
    workbook.save("controle_gastos.xlsx")

# Selecionando a planilha
sheet = workbook.active

# Pegando o mês atual
mes_atual = input("Digite o mês atual: ")

# Pegando o valor recebido
valor_recebido = float(input("Digite o valor recebido este mês: "))

# Pegando os gastos
gastos = []
while True:
    descricao_gasto = input("Digite a descrição do gasto (ou 'fim' para encerrar): ")
    if descricao_gasto == "fim":
        break
    valor_gasto = float(input("Digite o valor do gasto: "))
    gastos.append((descricao_gasto, valor_gasto))

# Adicionando o cabeçalho na planilha
sheet.append(["Mês", "Valor Recebido", "Descrição do Gasto", "Valor do Gasto"])

# Adicionando os valores na planilha
for gasto in gastos:
    sheet.append([mes_atual, valor_recebido, gasto[0], gasto[1]])

# Calculando o saldo restante
saldo_restante = valor_recebido - sum([gasto[1] for gasto in gastos])

# Adicionando o saldo restante na planilha
sheet.append(["Saldo Restante", saldo_restante])

# Salvando as alterações no arquivo excel
workbook.save("controle_gastos.xlsx")

# Imprimindo resumo dos gastos e saldo restante
print("\nResumo dos gastos:")
for gasto in gastos:
    print(f"{gasto[0]}: R${gasto[1]:.2f}")
print(f"Saldo restante: R${saldo_restante:.2f}")

# Abrindo o arquivo excel para visualização
abrir_excel = input("Deseja abrir o arquivo excel para visualização? (s/n): ")
if abrir_excel.lower() == "s":
    os.startfile("controle_gastos.xlsx")
