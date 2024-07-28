from docx import Document

document = Document()

#print(dir(document))
#print(help(document.add_paragraph))
document.add_paragraph('Salve Maria Santíssima')


document.save('DevCatolico.docx')