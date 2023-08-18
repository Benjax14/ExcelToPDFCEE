import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

# Lee el archivo Excel
excel_file = "ramos.xlsx"
df = pd.read_excel(excel_file)

# Separa los datos por "ramos"
ramos = df['Seleccione ramo/s'].str.split(', ').explode().unique()

# Crea el archivo PDF
pdf_file = "resultado.pdf"
doc = SimpleDocTemplate(pdf_file, pagesize=letter)
elements = []

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