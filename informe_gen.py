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


# Configurar el estilo global del documento
estilo = doc.styles['Normal']
fuente = estilo.font
fuente.name = 'Arial'  # Tipo de letra
fuente.size = Pt(10)   # Tamaño de letra

titulo_parrafo = doc.add_paragraph()
titulo = titulo_parrafo.add_run('INFORME DE ACTIVIDADES')
titulo_parrafo.alignment = WD_ALIGN_PARAGRAPH.CENTER
titulo.font.bold = True

enca1 = doc.add_paragraph()
enca1_valor = enca1.add_run('\tA')
enca1_valor.font.bold = True
enca1_valor = enca1.add_run('\t\t\t:\t')
enca1_valor = enca1.add_run(receptor.director_tercero)

enca2 = doc.add_paragraph()
enca2_valor = enca2.add_run('\tDE')
enca2_valor.font.bold = True
enca2_valor = enca2.add_run('\t\t\t:\t')
enca2_valor = enca2.add_run(emisor.nombre)

enca3 = doc.add_paragraph()
enca3_valor = enca3.add_run('\tNÚMERO DE')
enca3_valor.font.bold = True
enca3_valor = enca3.add_run('\t\t:\t')
enca3_valor = enca3.add_run(Contrato_Adenda()[1])
enca3_valor = enca3.add_run('\n\tCONTRATO')
enca3_valor.font.bold = True

aucnoc = list(fechas_entregables.keys())

enca4 = doc.add_paragraph()
enca4_valor = enca4.add_run('\tPERIODO DE')
enca4_valor.font.bold = True
enca4_valor = enca4.add_run('\t\t:\t')
enca4_valor = enca4.add_run(f'Del {datetime.strftime(datetime.strptime(emisor.inicio_contrato,"%d/%m/%y"),"%d de %B")} al {datetime.strftime(datetime.strptime(aucnoc[-1], "%d/%m/%y"),"%d de %B de %Y")}')
enca4_valor = enca4.add_run('\n\tCONTRATACIÓN')
enca4_valor.font.bold = True

enca5 = doc.add_paragraph()
enca5_valor = enca5.add_run('\tPERIODO EN EL')
enca5_valor.font.bold = True
enca5_valor = enca5.add_run('\n\tQUE SE OTROGA')
enca5_valor.font.bold = True
enca5_valor = enca5.add_run('\t:\t')
enca5_valor = enca5.add_run(f'Del {datetime.strftime(datetime.strptime(fecha_posterior,"%d/%m/%y"),"%d de %B")} al {datetime.strftime(datetime.strptime(entregable_actual,"%d/%m/%y"),"%d de %B de %Y")}')
enca5_valor.font.bold = True
enca5_valor = enca5.add_run('\n\tLA CONFORMIDAD')
enca5_valor.font.bold = True

# Guardar el documento
doc.save("documento_con_tabla.docx")