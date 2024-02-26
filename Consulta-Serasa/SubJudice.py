import os
from docx import Document

# Abrir o documento do Word existente

diretorysave = r'REGISTRO SUSPENSO\Enviar'
def createRegistro(cpf, name, card, value):
    print('GERANDO REGISTRO SUSPENSO...')
    doc = Document('modelregistro.docx')
    for paragraph in doc.paragraphs:
        if 'CPFCLIENTE' in paragraph.text:
            paragraph.text = paragraph.text.replace('CPFCLIENTE', cpf)
        elif 'NAMECLIENTE' in paragraph.text:
            paragraph.text = paragraph.text.replace('NAMECLIENTE', name)
        elif 'CARDCLIENTE' in paragraph.text:
            paragraph.text = paragraph.text.replace('CARDCLIENTE', card)
        elif 'VALUEPROTESTO' in paragraph.text:
            paragraph.text = paragraph.text.replace('VALUEPROTESTO', f'R$ {value}')
    files = os.listdir(diretorysave)
    filename = f'{name} - {card}.docx'
    if filename not in files:
        doc.save(f'{diretorysave}\\{filename}')
        print('REGISTRO SUSPENSO GERADO...')

