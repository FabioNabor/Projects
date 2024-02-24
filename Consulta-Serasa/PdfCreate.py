from docx import Document

# Abrir o documento do Word existente
doc = Document('modelo registro suspenso.docx')

# Percorrer os parágrafos do documento
for paragraph in doc.paragraphs:
    if '023.785.054-00' in paragraph.text:
        paragraph.text = paragraph.text.replace('023.785.054-00', '166.349.476-27')

# Salvar as alterações no documento
doc.save('documento_editado.docx')