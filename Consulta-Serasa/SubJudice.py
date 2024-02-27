import os
from docx import Document
import MainApp as mp
from tkinter import messagebox

# Abrir o documento do Word existente

diretorysave = r'Registro Suspenso\Enviar'
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

def getRegistro(diretory):
    doc = Document(diretory)
    getreturn = {'CPF':None,
                 'CONTRATO':None,
                 'VALOR':None}
    for paragraph in doc.paragraphs:
        if ':' in paragraph.text and len(paragraph.text) < 50:
            vectorp = str(paragraph.text).split(':')
            find = vectorp[0]
            inf = vectorp[1]
            if 'CPF' in find:
                getreturn['CPF'] = inf
            elif 'Contrato' in find:
                getreturn['CONTRATO'] = inf
            elif 'Valor' in find:
                getreturn['VALOR'] = inf
    return getreturn

if __name__ == "__main__":
    try:
        mp.downRegister()
        messagebox.showinfo('DownOrgãos', f'Registros Validados com Sucesso!')
    except Exception as e:
        messagebox.showerror('DownOrgãos', f'{e}')



