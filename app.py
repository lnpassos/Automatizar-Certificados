
# EXPLICAÇÃO PROJETO  & MAPA MENTAL

"""

Quero a possibilidade de criar um programa utilizando Python para automatizar enviando os dados
de uma planilha para preencher campos mutáveis em um certificado padrão.

Exemplo: Nome do Curso, Nome do Participante, Tipo de Participação, Data de Inicio, Data de Finalização
Carga Horária, Data da Emissão do Certificado e as assinaturas do Gestor Geral.


"""


"""
# Pegar os dados da planilha:
Tipo nome do curso, nome participante, tipo de participação, data do inicio, data do final, carga horária, data da emissão do certificado


# Transferir os dados da planilha para a imagem do certificado

"""
import openpyxl
from PIL import Image, ImageDraw, ImageFont

# Pegar os dados da planilha:

# Abrindo Planilha
workbook_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')
sheet_alunos = workbook_alunos['Sheet1']

for indice, linha in enumerate(sheet_alunos.iter_rows(min_row=2, max_row=10)): # OBS: Tirar o max_row para baixar todos os certificados
    # Armazenando as informações da planilha
    nome_curso = linha[0].value
    nome_participante = linha[1].value
    tipo_participacao = linha[2].value
    data_inicio = linha[3].value
    data_final = linha[4].value
    carga_horaria = linha[5].value
    data_emissao_certificado = linha[6].value

    # Transferir os dados da planilha para a imagem do certificado
    fonte_nome = ImageFont.truetype('./fonts/arialbd.ttf', 90)
    fonte_geral = ImageFont.truetype('./fonts/arial.ttf', 75)

    image = Image.open('./certificado_padrao.jpg')
    desenhar = ImageDraw.Draw(image)

    # OBS: Aqui pode variar pra cada pessoa, pode ser necessário ajustar as coordenadas!
    desenhar.text((1020, 835), nome_participante, fill='black', font=fonte_nome)
    desenhar.text((1060, 965), nome_curso, fill='black', font=fonte_geral)
    desenhar.text((1435, 1082), tipo_participacao, fill='black', font=fonte_geral)
    desenhar.text((1488, 1200), str(carga_horaria), fill='black', font=fonte_geral)
    desenhar.text((710, 1770), data_inicio, fill='blue', font=fonte_geral)
    desenhar.text((710, 1920), data_final, fill='blue', font=fonte_geral)
    desenhar.text((2190, 1915), data_emissao_certificado, fill='blue', font=fonte_geral)
    ####################################################################################

    image.save(f'./{indice} {nome_participante} certificado.png')
    

