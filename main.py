import zipfile
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, Toplevel, Label, ttk

from parser import extraer_datos_xml, extraer_datos_nomina, extraer_datos_deducciones
from exporter import export_to_excel

def pantalla_progreso(root):
    ventana = Toplevel(root)
    ventana.title("Procesando archivos")
    ventana.geometry("600x200")
    label = Label(ventana, text="Procesando archivos, por favor espera...", font=("Arial", 12))
    label.pack(pady=40)
    ventana.grab_set()
    return ventana

def procesar_zip_y_guardar_excel(root, opcion):
    def tarea():        
        ruta_zip = filedialog.askopenfilename(title="Selecciona archivo ZIP", filetypes=[("Archivos ZIP", "*.zip")])
        if not ruta_zip:
            return
            
        campos_nomina = ["EstadoSAT", "FechaEmision", "FechaTimbrado", "Serie", "Folio", "UUID",
                        "RFC Receptor", "NombreReceptor", "RFC Emisor", "NombreEmisor",
                        "RegistroPatronal", "TipoNomina", "FechaPago", "FechaInicialPago",
                        "FechaFinalPago", "NumDiasPagados", "TotalPercepciones", "TotalDeducciones",
                        "TotalOtrosPagos", "SubTotal", "Descuento", "Total", "MetodoPago", "Regimen",
                        "ArchivoXML", "Conceptos", "TipoComprobante", "Version", "Moneda",
                        "ReceptorCurp", "NumSeguridadSocial", "FechaInicioRelLaboral",
                        "TotalSueldosPer", "TotalGravadoPercepcion", "TotalExentoPercepcion",
                        "TotalOtrasDeducciones", "TotalImpuestosRetenidosDed",
                        "P01_SueldoSalarioExe", "P01_SueldoSalarioGra",
                        "P02_AguinaldoExe", "P02_AguinaldoGra",
                        "P05_FondodeAhorroExe", "P05_FondodeAhorroGra",
                        "P21_PrimVacacionalExe", "P21_PrimVacacionalGra",
                        "P29_ValesDespensaExe", "P29_ValesDespensaGra",
                        "P38_OtrosingresosGra", "D01_SeguroSocial", "D02_ISR",
                        "D04_OtrasDeducciones", "D10_CreditoVivienda"]
        
        campos_deducciones = ["TIPO", "Estado SAT", "Version", "Tipo", "Fecha Emision", "Fecha Timbrado",
                              "Serie", "Folio", "UUID", "UUID Relacion", "RFC Emisor", "Nombre Emisor",
                              "LugarDeExpedicion", "RFC Receptor", "Nombre Receptor", "UsoCFDI",
                              "SubTotal", "Descuento", "IVA 16%", "Total", "Moneda", "FormaDePago",
                              "Metodo de Pago", "Conceptos", "OBSERVACIONES"]

        ventana_cargando = pantalla_progreso(root)

        headers = []
        if opcion == "ambas":
            headers = ["Folio Fiscal", "Folio", "Fecha", "Concepto", "Subtotal", "IVA", "Retención IVA", "Retención ISR", "Total", "Estado del CFDI", "Tipo CFDI", "Metodo Pago", "Forma Pago", "Uso CFDI", "Moneda", "RFC_Emisor"]
        elif opcion == "ret_iva":
            headers = ["Folio Fiscal", "Folio", "Fecha", "Concepto", "Subtotal", "IVA", "Retención IVA", "Total", "Estado del CFDI", "Tipo CFDI", "Metodo Pago", "Forma Pago", "Uso CFDI", "Moneda", "RFC_Emisor"]
        elif opcion == "ret_isr":
            headers = ["Folio Fiscal", "Folio", "Fecha", "Concepto", "Subtotal", "IVA", "Retención ISR", "Total", "Estado del CFDI", "Tipo CFDI", "Metodo Pago", "Forma Pago", "Uso CFDI", "Moneda", "RFC_Emisor"]
        elif opcion == "ninguna":
            headers = ["Folio Fiscal", "Folio", "Fecha", "Concepto", "Subtotal", "IVA", "Total", "Estado del CFDI", "Tipo CFDI", "Metodo Pago", "Forma Pago", "Uso CFDI", "Moneda", "RFC_Emisor"]
        elif opcion == "nomina":
            headers = campos_nomina
        elif opcion == "deducciones":
            headers = campos_deducciones

        datos_filas = []

        with zipfile.ZipFile(ruta_zip, 'r') as zip_ref:
            archivos_xml = [f for f in zip_ref.namelist() if f.lower().endswith('.xml')]

            for nombre_archivo in archivos_xml:
                with zip_ref.open(nombre_archivo) as archivo:
                    archivo.name = nombre_archivo
                    
                    if opcion == "nomina":
                        datos = extraer_datos_nomina(archivo)
                        if datos:
                            fila = [datos.get(campo, "") for campo in campos_nomina]
                            datos_filas.append(fila)
                    elif opcion == "deducciones":
                        datos = extraer_datos_deducciones(archivo)
                        if datos:
                            fila = [datos.get(campo, "") for campo in campos_deducciones]
                            datos_filas.append(fila)
                    else:    
                        datos = extraer_datos_xml(archivo, opcion)
                        if datos:
                            datos_filas.append(datos)

        ventana_cargando.destroy()

        ruta_salida = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if ruta_salida:
            export_to_excel(datos_filas, headers, ruta_salida)
            messagebox.showinfo("Éxito", f"Datos extraídos y guardados en Excel correctamente en la ruta {ruta_salida}")
        # Restores the main window so the user can perform another action
        root.deiconify()
        
    threading.Thread(target=tarea).start()

def mostrar_ventana_principal():
    ventana = tk.Tk()
    ventana.title("Tipo de CFDI")
    ventana.geometry("320x350")

    tipo = tk.StringVar(value="ingresos")
    opcion = tk.StringVar(value="ninguna")

    ttk.Label(ventana, text="Selecciona el tipo de CFDI:").pack(pady=10)

    def mostrar_opciones():
        check_frame.pack(pady=5)

    def ocultar_opciones():
        check_frame.pack_forget()

    ttk.Radiobutton(ventana, text="Ingresos", variable=tipo, value="ingresos", command=mostrar_opciones).pack()
    ttk.Radiobutton(ventana, text="Gastos", variable=tipo, value="gastos", command=ocultar_opciones).pack()
    
    check_frame = ttk.Frame(ventana)
    ttk.Radiobutton(check_frame, text="Retención IVA", variable=opcion, value="ret_iva").pack(anchor="w")
    ttk.Radiobutton(check_frame, text="Retención ISR", variable=opcion, value="ret_isr").pack(anchor="w")
    ttk.Radiobutton(check_frame, text="Ambas", variable=opcion, value="ambas").pack(anchor="w")
    ttk.Radiobutton(check_frame, text="Ninguna", variable=opcion, value="ninguna").pack(anchor="w")
    ttk.Radiobutton(check_frame, text="Deducciones", variable=opcion, value="deducciones").pack(anchor="w")
    ttk.Radiobutton(check_frame, text="Nómina", variable=opcion, value="nomina").pack(anchor="w")

    mostrar_opciones()

    def continuar():
        ventana.withdraw()
        if tipo.get() != "ingresos":
            opcion.set("ninguna")
        procesar_zip_y_guardar_excel(ventana, opcion.get())

    ttk.Button(ventana, text="Procesar Archivos", command=continuar).pack(pady=20)

    ventana.mainloop()

if __name__ == "__main__":
    mostrar_ventana_principal()
