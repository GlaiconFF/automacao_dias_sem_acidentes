import re
from docx import Document
from datetime import datetime
import subprocess
from docx2pdf import convert

data_inicio = datetime(2025, 7, 24)

data_atual = datetime.now()
data_atual_str = data_atual.strftime('%d/%m/%Y')

diferenca_dias = (data_atual - data_inicio).days

numero_base_etapa = 75
numero_base_sem_acidentes = 195
numero_atualizado_etapa = numero_base_etapa + diferenca_dias
numero_atualizado_sem_acidentes = numero_base_sem_acidentes + diferenca_dias

print(f"Data atual: {data_atual_str}")
print(f"Número etapa atualizado: {numero_atualizado_etapa}")
print(f"Número sem acidentes atualizado: {numero_atualizado_sem_acidentes}")

doc = Document('dias_sem_acidentes.docx')

regex_data = r'\b\d{2}/\d{2}/\d{4}\b'
regex_numeros = r'\b\d{3,4}\b'

data_substituida = False
contagem_numeros_substituidos = 0

for par in doc.paragraphs:
    for run in par.runs:

        if not data_substituida:
            datas_encontradas = re.findall(regex_data, run.text)
            if datas_encontradas:
                antiga = datas_encontradas[0]
                run.text = run.text.replace(antiga, data_atual_str, 1)
                print(f"Substituindo data {antiga} por {data_atual_str}")
                data_substituida = True

        numeros = re.findall(regex_numeros, run.text)
        for numero in numeros:
            if len(numero) == 3:
                if contagem_numeros_substituidos == 1:
                    run.text = run.text.replace(numero, str(numero_atualizado_etapa), 1)
                    print(f"Substituindo número de etapa {numero} por {numero_atualizado_etapa}")
                    contagem_numeros_substituidos += 1
                    break
                elif contagem_numeros_substituidos == 0:
                    run.text = run.text.replace(numero, str(numero_atualizado_sem_acidentes), 1)
                    print(f"Substituindo número de dias sem acidentes {numero} por {numero_atualizado_sem_acidentes}")
                    contagem_numeros_substituidos += 1
                    break

doc.save('dias_sem_acidentes.docx')
print(f"\nDocumento editado salvo como 'dias_sem_acidentes.docx'.")

sumatra = r"C:\Users\glaicon.felipe\AppData\Local\SumatraPDF\SumatraPDF.exe"
printer = "NPI988D9B (HP LaserJet MFP E42540)"
arquivo = "dias_sem_acidentes.docx"

convert(arquivo)
pdf = "dias_sem_acidentes.pdf"

settings = "1x,simplex"
cmd = [
    sumatra,
    "-print-to", printer,
    "-print-settings", settings,
    pdf
]

print(f"Printer: {printer}\nPDF: {pdf}\nSettings: {settings}\nExecutando: {cmd}\n")
subprocess.run(cmd, check=True)
print("Impressão enviada:", printer)
