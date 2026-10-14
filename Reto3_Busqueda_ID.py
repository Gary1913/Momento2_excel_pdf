
import openpyxl
from tabulate import tabulate

wb = openpyxl.load_workbook("ventas_tecnologia_v2.xlsx", data_only=True)
sheet = wb.active


id_buscado = int(input("Ingrese el ID del producto: "))


data = []

for row in sheet.iter_rows(min_row=2, values_only=True):
    if row[0] == id_buscado:
        data.append(list(row))

headers = ["ID", "Producto", "Categoría", "Precio ($)", "Ventas (unidad)"]
aligments = ("center", "left", "center", "right", "right")


if len(data) > 0:
    print(f"\n=== PRODUCTO CON ID {id_buscado} ===")
    print(tabulate(data, headers=headers, tablefmt="fancy_grid", colalign=aligments))
else:
    print(f"\n[ADVERTENCIA] No existe ningún producto con el ID {id_buscado}.")
