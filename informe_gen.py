from docx import Document
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from data import *

# Crear el documento
doc = Document()

img_path = "utils/oefa-logo-header.png"

section = doc.sections[0]
section.left_margin = Cm(Margenes_Informe.izquierdo)    # Margen izquierdo
section.right_margin = Cm(Margenes_Informe.derecho)   # Margen derecho
section.top_margin = Cm(Margenes_Informe.superior)     # Margen superior
section.bottom_margin = Cm(Margenes_Informe.inferior)  # Margen inferior

# Crear el encabezado
header = section.header

# Configurar la tabla en el encabezado
tabla = header.add_table(rows=1, cols=3, width=Cm(16.7))  # Ancho total de la tabla
tabla.style = 'Table Grid'
tabla.alignment = WD_TABLE_ALIGNMENT.CENTER

# Primera celda (imagen)
celda_1 = tabla.cell(0, 0)
celda_1.height = Cm(tabla_cabecera.celda_alto)
celda_1.width = Cm(tabla_cabecera.celda1_ancho)
celda_1.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
parrafo_imagen = celda_1.paragraphs[0]
parrafo_imagen.alignment = WD_ALIGN_PARAGRAPH.CENTER
parrafo_imagen.add_run().add_picture(img_path, width=Cm(3.8), height=Cm(0.80))

# Segunda celda (texto centrado)
celda_2 = tabla.cell(0, 1)
celda_2.height = Cm(tabla_cabecera.celda_alto)
celda_2.width = Cm(tabla_cabecera.celda2_ancho)
celda_2.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
parrafo_2 = celda_2.paragraphs[0]
parrafo_2.alignment = WD_ALIGN_PARAGRAPH.CENTER
parrafo_2.add_run(tabla_cabecera.texto_mapro)

# Tercera celda (texto pequeño, alineado a la izquierda)
celda_3 = tabla.cell(0, 2)
celda_3.height = Cm(tabla_cabecera.celda_alto)
celda_3.width = Cm(tabla_cabecera.celda3_ancho)
parrafo_3 = celda_3.paragraphs[0]
parrafo_3.alignment = WD_ALIGN_PARAGRAPH.LEFT
run = parrafo_3.add_run(tabla_cabecera.texto_version)
run.font.size = Pt(9)

# Guardar el documento
doc.save("documento_con_tabla.docx")