import openpyxl
from tabulate import tabulate
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

wb = openpyxl.load_workbook("ventas_tecnologia_v2.xlsx", data_only=True)
sheet = wb.active


data = []

for row in sheet.iter_rows(min_row=2, values_only=True):
    data.append(list(row))

headers = ["ID", "Producto", "Categoría", "Precio ($)", "Ventas (unidad)"]
aligments = ("center", "left", "center", "right", "right")


total_unidades = 0
sumatoria_precio = 0

for r in data:
    total_unidades = total_unidades + r[4]
    sumatoria_precio = sumatoria_precio + r[3]

precio_promedio = sumatoria_precio / len(data) if len(data) > 0 else 0


print(f"\n=== REPORTE DE VENTAS EN CONSOLA ({len(data)} PRODUCTOS) ===")
print(tabulate(data, headers=headers, tablefmt="fancy_grid", colalign=aligments))
print(f"\nTotal de unidades vendidas: {total_unidades:,}")
print(f"Precio promedio: ${precio_promedio:,.2f}")


doc = SimpleDocTemplate("reporte_ejecutivo_ventas.pdf", pagesize=letter)
styles = getSampleStyleSheet()

title_style = ParagraphStyle(
    "TitleStyle", parent=styles["Title"],
    fontName="Helvetica-Bold", fontSize=18,
    textColor=colors.HexColor("#1A365D"), spaceAfter=10)

body_style = ParagraphStyle("BodyStyle", parent=styles["Normal"], fontSize=9)

header_style = ParagraphStyle(
    "HeaderStyle", parent=styles["Normal"],
    fontName="Helvetica-Bold", fontSize=10,
    textColor=colors.whitesmoke, alignment=1)

story = [
    Paragraph("Informe Ejecutivo de Ventas Tecno", title_style),
    Paragraph(f"<b>Total de Productos Procesados:</b> {len(data)}", body_style),
    Spacer(1, 12),
]

headers_paragraphs = [Paragraph(f"<b>{h}</b>", header_style) for h in headers]
data_paragraphs = [
    [Paragraph(str(r[0]), ParagraphStyle("C", parent=body_style, alignment=1)),
     Paragraph(str(r[1]), ParagraphStyle("L", parent=body_style, alignment=0)),
     Paragraph(str(r[2]), ParagraphStyle("C", parent=body_style, alignment=1)),
     Paragraph(f"${r[3]:,.2f}", ParagraphStyle("R", parent=body_style, alignment=2)),
     Paragraph(f"{r[4]:,}", ParagraphStyle("R", parent=body_style, alignment=2))]
    for r in data
]

# --- Fila de resumen (última fila de la tabla) ---
summary_row = [
    Paragraph("", body_style),
    Paragraph("<b>RESUMEN</b>", body_style),
    Paragraph("", body_style),
    Paragraph(f"<b>${precio_promedio:,.2f}</b>", ParagraphStyle("R", parent=body_style, alignment=2)),
    Paragraph(f"<b>{total_unidades:,}</b>", ParagraphStyle("R", parent=body_style, alignment=2)),
]

cols_widths = [40, 160, 100, 80, 80]
pdf_table = Table([headers_paragraphs] + data_paragraphs + [summary_row],
                  colWidths=cols_widths, repeatRows=1)

pdf_table.setStyle(TableStyle([
    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1A365D")),
    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E0")),
    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F7FAFC")]),
]))

story.append(pdf_table)
doc.build(story)

print("\n[EXITO] Archivo \"reporte_ejecutivo_ventas.pdf\" generado correctamente.")
