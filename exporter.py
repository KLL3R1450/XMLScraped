from openpyxl import Workbook
import os

def export_to_excel(datos_filas, headers, ruta_salida):
    wb = Workbook()
    ws = wb.active
    ws.title = "Datos XML"

    # Append header row
    ws.append(headers)

    # Append data rows
    for fila in datos_filas:
        # Convert any string representation of numbers to actual floats where possible,
        # but parser already did this for the values we care about. We just need to make sure we append them correctly.
        processed_row = []
        for val in fila:
            if isinstance(val, (int, float)):
                processed_row.append(val)
            else:
                # If it's a string, try converting to float if it looks like a number, otherwise write as string
                val_str = str(val)
                try:
                    # Remove commas/formatting if any, try float parsing
                    if val_str.replace('.', '', 1).replace('-', '', 1).isdigit():
                        processed_row.append(float(val_str))
                    else:
                        processed_row.append(val_str)
                except ValueError:
                    processed_row.append(val_str)
        ws.append(processed_row)

    wb.save(ruta_salida)
    wb.close()
