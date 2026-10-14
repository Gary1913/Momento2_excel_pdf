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



def generar_pdf():
    
    total_unidades = 0
    sumatoria_precio = 0

    for r in data:
        total_unidades = total_unidades + r[4]
        sumatoria_precio = sumatoria_precio + r[3]

    precio_promedio = sumatoria_precio / len(data) if len(data) > 0 else 0

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



while True:
    print("\n========== MENÚ VENTAS TECNO ==========")
    print("1) Ver catálogo completo")
    print("2) Filtrar por precio (> $100)")
    print("3) Buscar por ID")
    print("4) Exportar informe PDF")
    print("5) Salir")

    opcion = input("Elige una opción (1-5): ")

   
    if opcion == "1":
        print(f"\n=== CATÁLOGO COMPLETO ({len(data)} PRODUCTOS) ===")
        print(tabulate(data, headers=headers, tablefmt="fancy_grid", colalign=aligments))

    
    elif opcion == "2":
        filtrados = []
        for r in data:
            if r[3] > 100:
                filtrados.append(r)

        print(f"\n=== PRODUCTOS CON PRECIO MAYOR A $100 ({len(filtrados)} PRODUCTOS) ===")
        print(tabulate(filtrados, headers=headers, tablefmt="fancy_grid", colalign=aligments))

    
    elif opcion == "3":
        texto_id = input("Ingrese el ID del producto: ")

       
        if texto_id.isdigit():
            id_buscado = int(texto_id)

            encontrados = []
            for r in data:
                if r[0] == id_buscado:
                    encontrados.append(r)

            if len(encontrados) > 0:
                print(f"\n=== PRODUCTO CON ID {id_buscado} ===")
                print(tabulate(encontrados, headers=headers, tablefmt="fancy_grid", colalign=aligments))
            else:
                print(f"\n[ADVERTENCIA] No existe ningún producto con el ID {id_buscado}.")
        else:
            print("\n[ADVERTENCIA] El ID debe ser un número.")

   
    elif opcion == "4":
        generar_pdf()

   
    elif opcion == "5":
        print("\n¡Hasta luego!")
        break

    
    else:
        print("\n[ADVERTENCIA] Opción no válida. Escribe un número del 1 al 5.")
