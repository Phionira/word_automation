import os, sys
from docxtpl import DocxTemplate

# Change path to current working directory
os.chdir(sys.path[0])

doc = DocxTemplate('Template.docx')
context = { 'fecha' : '20 de octubre de 2023', 'escuela' : "Mexico nuevo"}

doc.render(context)
doc.save("oficio_1_rendered.docx")