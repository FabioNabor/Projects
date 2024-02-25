import os
from docx import Document

# Abrir o documento do Word existente
doc = Document('modelregistro.docx')
diretorysave = r'REGISTRO SUSPENSO\Enviar'
def createRegistro(cpf, name, card, value):
    for paragraph in doc.paragraphs:
        paragraph.text = paragraph.text.replace('CPFCLIENTE', cpf)
        paragraph.text = paragraph.text.replace('NAMECLIENTE', name)
        paragraph.text = paragraph.text.replace('CARDCLIENTE', card)
        paragraph.text = paragraph.text.replace('VALUEPROTESTO', value)
    try:
        os.makedirs(diretorysave)
    except:
        pass
    files = os.listdir(diretorysave)
    filename = f'{name} - {card}.docx'
    if filename not in files:
        doc.save(f'{diretorysave}\\{filename}')

