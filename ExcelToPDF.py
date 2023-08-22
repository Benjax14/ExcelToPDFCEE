import pandas as pd
import locale
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, PageBreak
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from datetime import datetime

# Configura el locale para español
locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')

# Lee el archivo Excel
excel_file = "ramos.xlsx"
df = pd.read_excel(excel_file)

# Obtiene la fecha actual
actualDate = datetime.now()

# Formatea la fecha actual en el formato deseado (por ejemplo, "día de mes de año")
fixedDate = actualDate.strftime("%d de %B de %Y")

# Separa los datos por "ramos"
ramos = df['Seleccione ramo/s'].str.split(', ').explode().unique()

# Crea el archivo PDF
pdf_file = "resultado.pdf"
doc = SimpleDocTemplate(pdf_file, pagesize=letter)
elements = []

cover_text = """
<size=30><b>Resultados de Ramos no Impartidos</b></size>
"""

cover_date = f"""
<size=20>Fecha: {fixedDate}</size>
"""

# Agrega una portada al principio del documento
cover_title_paragraph = Paragraph(cover_text, getSampleStyleSheet()['Title'])
cover_date_paragraph = Paragraph(cover_date, getSampleStyleSheet()['Title'])
elements.append(cover_title_paragraph)
elements.append(cover_date_paragraph)
elements.append(PageBreak())  # Agregar un salto de página después de la portada

# Crea una tabla separada para cada "ramo"
for ramo in ramos:
    ramo_data = df[df['Seleccione ramo/s'].str.contains(ramo, na=False)][['Nombre Completo', 'Rut', 'Correo', 'Generación']]
    table_data = [ramo_data.columns.tolist()] + ramo_data.values.tolist()
    
    # Agrega el título del ramo antes de la tabla
    styles = getSampleStyleSheet()
    title_style = styles['Title']
    title = Paragraph(f"<b>{ramo}</b>", title_style)
    elements.append(title)

    table = Table(table_data)
    table.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                               ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                               ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                               ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                               ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                               ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                               ('GRID', (0, 0), (-1, -1), 1, colors.black)]))
    
    elements.append(table)

# Agrega las tablas al documento y cierra el PDF
doc.build(elements)

print(f"Archivo PDF '{pdf_file}' creado con éxito.")
