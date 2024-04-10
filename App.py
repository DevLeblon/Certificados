#Erick Luan Ferreira de Azevedo
#Pair-programing coop

#Pegar dados da planilha e colocar no .jpg
#curso, participante, tipo de participação, data inicio, data final, carga horaria, data emissão

import openpyxl
from PIL import Image, ImageDraw, ImageFont
from datetime import datetime

# Abrir a planilha
workbook_aluno = openpyxl.load_workbook('planilha_alunos 1.xlsx')
sheet_alunos = workbook_aluno['Sheet1']

# Definir as fontes
fonte_name = ImageFont.truetype("./tahoma.ttf", 13)
fonte_geral = ImageFont.truetype("./tahoma.ttf", 13)
fonte_data = ImageFont.truetype("./tahoma.ttf", 55)

for indice, linha in enumerate(sheet_alunos.iter_rows(min_row=2, max_row=4)):
    # Extrair os dados da linha
    nome_curso = linha[0].value
    nome_participante = linha[1].value
    registro_geral = linha[2].value
    data_emissao = linha[3].value


    # Converter as datas para objetos datetime
if isinstance(data_emissao, str):
    data_emissao = datetime.strptime(data_emissao, "%d/%m/%Y")

    # Abrir a imagem do certificado
    image = Image.open("./certificado_padrao.jpg")
    desenhar = ImageDraw.Draw(image)

    # Desenhar os textos na imagem
    desenhar.text((142, 178), nome_participante, fill='black', font=fonte_name)
    desenhar.text((268, 178), registro_geral, fill='black', font=fonte_geral)
    desenhar.text((1435, 1065), data_emissao, fill='black', font=fonte_geral)

    desenhar.text((2220, 1930), data_emissao.strftime("%d/%m/%Y"), fill='blue', font=fonte_data)

    # Salvar a imagem do certificado
    image.save(f'./{indice} {nome_participante} certificado.png')