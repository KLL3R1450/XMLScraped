import zipfile
import xml.etree.ElementTree as ET
from openpyxl import Workbook
from tkinter import filedialog
import tkinter as tk
from tkinter import messagebox, Toplevel, Label, ttk
from zeep import Client
import threading

def extraer_datos_xml(archivo_xml,opcion):
    try:
        tree = ET.parse(archivo_xml)
        root = tree.getroot()

        
         
        nombre = archivo_xml.name
        nombre = nombre.replace('.xml', '').replace('.XML', '')  

        version = root.attrib.get("Version", root.attrib.get("version"))
      
        if version.startswith("3."):
            namespaces = {'cfdi': 'http://www.sat.gob.mx/cfd/3'}
            namespace_uuid = {'tfd': 'http://www.sat.gob.mx/TimbreFiscalDigital'}
            
        elif version.startswith("4."):
            namespaces = {'cfdi': 'http://www.sat.gob.mx/cfd/4'}
            namespace_uuid = {'tfd': 'http://www.sat.gob.mx/TimbreFiscalDigital'}
        else:
            raise ValueError("Version no registrada en el sistema")
        
        try:
            uuid = root.find('.//tfd:TimbreFiscalDigital', namespaces=namespace_uuid).attrib['UUID']
        except Exception as e:
            print(f"Error al obtener UUID: {e}")
            return None

        comprobante = root
        folio = comprobante.attrib.get('Folio')
        fecha = comprobante.attrib.get('Fecha')
        subtotal = comprobante.attrib.get('SubTotal') or comprobante.attrib.get('subtotal')
        total = comprobante.attrib.get('Total') or comprobante.attrib.get('total')
        tipoComprobante = comprobante.attrib.get('TipoDeComprobante')
        rfc_emisor = root.find('.//cfdi:Emisor', namespaces).attrib.get('Rfc')
        rfc_receptor = root.find('.//cfdi:Receptor', namespaces).attrib.get('Rfc')
        iva_total = 0.00
        estado = estaCancelado(uuid, rfc_emisor, rfc_receptor, total)
        metodoPago = comprobante.attrib.get('MetodoPago')
        formaPago = get_forma_pago(str(comprobante.attrib.get('FormaPago')))
        uso = get_uso_cfdi(str(root.find('.//cfdi:Receptor',namespaces).attrib.get('UsoCFDI')))

        conceptos = root.findall('.//cfdi:Concepto', namespaces)

        for concepto in conceptos:
            traslados = concepto.findall('.//cfdi:Traslado', namespaces)
            for traslado in traslados:
                if traslado.attrib.get('Impuesto') == '002':  # IVA
                    importe = float(traslado.attrib.get('Importe', '0'))
                    iva_total += importe
        
        concepto = ""
       
        if conceptos:
            for con in conceptos:
                concepto += ","+ con.attrib.get('Descripcion') or con.attrib.get('descripcion')

        if opcion == "ambas":
            retencion_iva = get_retIva(conceptos, namespaces)
            retencion_isr = get_retIsr(conceptos, namespaces)
            return [nombre, folio, fecha, concepto, subtotal, iva_total,retencion_iva,retencion_isr, total, estado,tipoComprobante,metodoPago,formaPago,uso]
        elif opcion == "ret_isr":
            retencion_Isr = get_retIsr(root, namespaces)
            return [nombre, folio, fecha, concepto, subtotal, iva_total,retencion_Isr, total, estado,tipoComprobante,metodoPago,formaPago,uso]
        elif  opcion == "ret_iva" :
            retencion_iva = get_retIva(root, namespaces)
            return [nombre, folio, fecha, concepto, subtotal, iva_total,retencion_iva, total, estado,tipoComprobante,metodoPago,formaPago,uso]
        elif opcion == "ninguna":
            return [nombre, folio, fecha, concepto, subtotal, iva_total, total, estado,tipoComprobante,metodoPago,formaPago,uso]

        

    except Exception as e:
        print(f"Error al procesar archivo XML: {e}")
        return None

def get_forma_pago(codigo: str):
    match codigo:
        case '01' : return 'Efectivo'
        case '02' : return 'Cheque'
        case '03' : return 'Transferencia'
        case '04' : return 'Tarjeta de Credito'
        case '05' : return 'Monedero Electronico'
        case '06' : return 'Dinero Electronico'
        case '07' : return 'Tarjetas Digitales'
        case '08' : return 'Vales de Despensa'
        case '09' : return 'Bienes'
        case '10' : return 'Servicio'
        case '11' : return 'Por Cuenta de Tercero'
        case '12' : return 'Dacion de Pago'
        case '13' : return 'Pago de Subrogacion'
        case '14' : return 'Pago de Consignacion'
        case '15' : return 'Condonacion'
        case '16' : return 'Cancelacion'
        case '17' : return 'Compensacion'
        case '98' : return 'NA'
        case '99' : return 'Parcialidades o diferido'
        case _: return "NA"

def get_uso_cfdi(codigo:str):
    match codigo:
        case 'G01' : return 'Adquisicion de Mercancias'
        case 'G02' : return 'Devoluciones, descuento o bonificaciones'
        case 'G03' : return 'Gastos en General'
        case 'I01' : return 'Construcciones'
        case 'I02' : return 'Mobiliarion y Equipo de Oficina por construcciones'
        case 'I03' : return 'Equipo de Transporte'
        case 'I04' : return 'Equipo de computo y accesorios'
        case 'I05' : return 'Dados, troqueles, modeles, matrices y herramientas'
        case 'I06' : return 'Comunicaciones telefonicas'
        case 'I07' : return 'Comunicaciones satelitales'
        case 'I08' : return 'Otras maquinas y equipo'
        case 'D01' : return 'Honorarios medicos, dentales y hospitalarios'
        case 'D02' : return 'Gastos medicos por incapacidad o discapacidad'
        case 'D03' : return 'Gastos funerales'
        case 'D04' : return 'Donativos'
        case 'D05' : return 'Intereses reales efectivamente pagados por creditos hipotecarios (Casa Habitacion)'
        case 'D06' : return 'Aportaciones voluntarias al SAR'
        case 'D07' : return 'Primas por seguros de gastos medicos'
        case 'D08' : return 'Gastos por transportacion escolar obligatoria'
        case 'D09' : return 'Depositos en cuentas de ahorro, primas que tengan como base planes de pensiones'
        case 'D10' : return 'Pagos por servicios educativos(colegiaturas)'
        case 'P01' : return 'Por Definir' 
        case _ : return f"No registrado ante el sat el codigo es {codigo}"
        
def extraer_datos_nomina(archivo_xml):
    import xml.etree.ElementTree as ET
    try:
        tree = ET.parse(archivo_xml)
        root = tree.getroot()
        namespaces = {
            'cfdi': 'http://www.sat.gob.mx/cfd/4',
            'nomina12': 'http://www.sat.gob.mx/nomina12',
            'tfd': 'http://www.sat.gob.mx/TimbreFiscalDigital'
        }

        nombre_archivo = archivo_xml.name

        comprobante = root
        uuid = root.find('.//tfd:TimbreFiscalDigital', namespaces).attrib.get('UUID')
        estado_sat = estaCancelado(uuid,
                                   root.find('.//cfdi:Emisor', namespaces).attrib.get('Rfc'),
                                   root.find('.//cfdi:Receptor', namespaces).attrib.get('Rfc'),
                                   comprobante.attrib.get('Total'))

        datos = {
            "EstadoSAT": estado_sat,
            "FechaEmision": comprobante.attrib.get("Fecha"),
            "FechaTimbrado": root.find('.//tfd:TimbreFiscalDigital', namespaces).attrib.get("FechaTimbrado"),
            "Serie": comprobante.attrib.get("Serie"),
            "Folio": comprobante.attrib.get("Folio"),
            "UUID": uuid,
            "RFC Receptor": root.find('.//cfdi:Receptor', namespaces).attrib.get("Rfc"),
            "NombreReceptor": root.find('.//cfdi:Receptor', namespaces).attrib.get("Nombre"),
            "RFC Emisor": root.find('.//cfdi:Emisor', namespaces).attrib.get("Rfc"),
            "NombreEmisor": root.find('.//cfdi:Emisor', namespaces).attrib.get("Nombre"),
        }

        nomina = root.find('.//nomina12:Nomina', namespaces)
        receptor_n = nomina.find('.//nomina12:Receptor', namespaces)
        percepciones = nomina.find('.//nomina12:Percepciones', namespaces)
        deducciones = nomina.find('.//nomina12:Deducciones', namespaces)

        datos.update({
            "RegistroPatronal": nomina.find('.//nomina12:Emisor', namespaces).attrib.get("RegistroPatronal"),
            "TipoNomina": nomina.attrib.get("TipoNomina"),
            "FechaPago": nomina.attrib.get("FechaPago"),
            "FechaInicialPago": nomina.attrib.get("FechaInicialPago"),
            "FechaFinalPago": nomina.attrib.get("FechaFinalPago"),
            "NumDiasPagados": nomina.attrib.get("NumDiasPagados"),
            "TotalPercepciones": nomina.attrib.get("TotalPercepciones"),
            "TotalDeducciones": nomina.attrib.get("TotalDeducciones"),
            "TotalOtrosPagos": nomina.attrib.get("TotalOtrosPagos"),
            "SubTotal": nomina.attrib.get("Subtotal"),
            "Descuento": nomina.attrib.get("TotalDeducciones"),
            "Total": comprobante.attrib.get("Total"),
            "MetodoPago": comprobante.attrib.get("MetodoPago"),
            "Regimen": receptor_n.attrib.get("TipoRegimen"),
            "ArchivoXML": nombre_archivo,
            "Conceptos": root.find('.//cfdi:Concepto', namespaces).attrib.get("Descripcion"),
            "TipoComprobante": comprobante.attrib.get("TipoDeComprobante"),
            "Version": comprobante.attrib.get("Version"),
            "Moneda": comprobante.attrib.get("Moneda"),
            "ReceptorCurp": receptor_n.attrib.get("Curp"),
            "NumSeguridadSocial": receptor_n.attrib.get("NumSeguridadSocial"),
            "FechaInicioRelLaboral": receptor_n.attrib.get("FechaInicioRelLaboral"),
            "TotalSueldosPer": percepciones.attrib.get("TotalSueldos"),
            "TotalGravadoPercepcion": percepciones.attrib.get("TotalGravado"),
            "TotalExentoPercepcion": percepciones.attrib.get("TotalExento"),
            "TotalOtrasDeducciones": deducciones.attrib.get("TotalOtrasDeducciones"),
            "TotalImpuestosRetenidosDed": deducciones.attrib.get("TotalImpuestosRetenidos"),
        })

        # Mapear percepciones específicas
        percepcion_map = {
            '001': ('P01_SueldoSalarioGra', 'P01_SueldoSalarioExe'),
            '002': ('P02_AguinaldoGra', 'P02_AguinaldoExe'),
            '005': ('P05_FondodeAhorroGra', 'P05_FondodeAhorroExe'),
            '021': ('P21_PrimVacacionalGra', 'P21_PrimVacacionalExe'),
            '029': ('P29_ValesDespensaGra', 'P29_ValesDespensaExe'),
            '038': ('P38_OtrosingresosGra', None)
        }

        for code, (gra, exe) in percepcion_map.items():
            datos[gra] = "0.00"
            if exe:
                datos[exe] = "0.00"

        for per in percepciones.findall('.//nomina12:Percepcion', namespaces):
            tipo = per.attrib.get('TipoPercepcion')
            if tipo in percepcion_map:
                gra, exe = percepcion_map[tipo]
                datos[gra] = per.attrib.get('ImporteGravado', '0.00')
                if exe:
                    datos[exe] = per.attrib.get('ImporteExento', '0.00')

        # Mapear deducciones específicas
        deduccion_map = {
            '001': 'D01_SeguroSocial',
            '002': 'D02_ISR',
            '004': 'D04_OtrasDeducciones',
            '010': 'D10_CreditoVivienda'
        }

        for val in deduccion_map.values():
            datos[val] = "0.00"

        for ded in deducciones.findall('.//nomina12:Deduccion', namespaces):
            tipo = ded.attrib.get('TipoDeduccion')
            if tipo in deduccion_map:
                datos[deduccion_map[tipo]] = ded.attrib.get('Importe', '0.00')

        return datos

    except Exception as e:
        print(f"Error al procesar XML de nómina: {e}")
        return None

def extraer_datos_deducciones(archivo_xml):
    tree = ET.parse(archivo_xml)
    root = tree.getroot()
    
    namespaces = {
    'cfdi': 'http://www.sat.gob.mx/cfd/4',
    'tfd': 'http://www.sat.gob.mx/TimbreFiscalDigital'
    }

    comprobante = root
    emisor = root.find('cfdi:Emisor', namespaces)
    receptor = root.find('cfdi:Receptor', namespaces)
    timbre = root.find('.//tfd:TimbreFiscalDigital', namespaces)

    conceptos = root.findall('.//cfdi:Concepto', namespaces)
    conceptos_list = [c.attrib.get('Descripcion', '').strip() for c in conceptos]

    # IVA solo dentro de conceptos
    iva_total = 0.0
    for concepto in conceptos:
        for traslado in concepto.findall('.//cfdi:Traslado', namespaces):
            if traslado.attrib.get('Impuesto') == '002':  # IVA 16%
                iva_total += float(traslado.attrib.get('Importe', '0'))

    datos = {
        "TIPO": comprobante.attrib.get('TipoDeComprobante', ''),
        "Estado SAT": "Pendiente o Cancelado (requiere consulta al SAT)",
        "Version": comprobante.attrib.get('Version', ''),
        "Tipo": comprobante.attrib.get('TipoDeComprobante', ''),
        "Fecha Emision": comprobante.attrib.get('Fecha', ''),
        "Fecha Timbrado": timbre.attrib.get('FechaTimbrado', ''),
        "Serie": comprobante.attrib.get('Serie', ''),
        "Folio": comprobante.attrib.get('Folio', ''),
        "UUID": timbre.attrib.get('UUID', ''),
        "UUID Relacion": "",  # Solo si hay nodo <cfdi:CfdiRelacionados>
        "RFC Emisor": emisor.attrib.get('Rfc', ''),
        "Nombre Emisor": emisor.attrib.get('Nombre', ''),
        "LugarDeExpedicion": comprobante.attrib.get('LugarExpedicion', ''),
        "RFC Receptor": receptor.attrib.get('Rfc', ''),
        "Nombre Receptor": receptor.attrib.get('Nombre', ''),
        "UsoCFDI": receptor.attrib.get('UsoCFDI', ''),
        "SubTotal": comprobante.attrib.get('SubTotal', ''),
        "Descuento": comprobante.attrib.get('Descuento', ''),
        "IVA 16%": round(iva_total, 2),
        "Total": comprobante.attrib.get('Total', ''),
        "Moneda": comprobante.attrib.get('Moneda', ''),
        "FormaDePago": comprobante.attrib.get('FormaPago', ''),
        "Metodo de Pago": comprobante.attrib.get('MetodoPago', ''),
        "Conceptos": " | ".join(conceptos_list),
        "OBSERVACIONES": ""  # Puedes rellenar manualmente o dejarlo vacío
    }

    return datos

def get_retIva(conceptos,namespaces):
    try:
        ret_iva_total = 0.00
        
        for concepto in conceptos:
            retenciones = concepto.findall('.//cfdi:Retencion',namespaces)
            for retencion in retenciones:
                if retencion.attrib.get('Impuesto') == '002':
                    importe = float(retencion.attrib.get('Importe'))
                    ret_iva_total += importe
             
        return ret_iva_total
    except Exception as e:
        return "No aplica"
    
def get_retIsr(conceptos, namespaces):
    try:
        ret_isr_total = 0.00
        
        for concepto in conceptos:
            retenciones = concepto.findall('.//cfdi:Retencion',namespaces)
            for retencion in retenciones:
                if retencion.attrib.get('Impuesto') == '001':
                    importe = float(retencion.attrib.get('Importe'))
                    ret_isr_total += importe
                
        return ret_isr_total
    except Exception as e:
        return "No aplica"

def estaCancelado(uuid, rfc_emisor, rfc_receptor, total):
    try:
        total_formateado = format(float(total), '.6f').zfill(17)
        client = Client('https://consultaqr.facturaelectronica.sat.gob.mx/ConsultaCFDIService.svc?wsdl')
        respuesta = client.service.Consulta(
            expresionImpresa=f"?re={rfc_emisor}&rr={rfc_receptor}&tt={total_formateado}&id={uuid}"
        )
        return respuesta.Estado
    except Exception as e:
        print(f"Error al consultar estado de CFDI: {e}")
        return 'Error'

def pantalla_progreso(root):
    ventana = Toplevel(root)
    ventana.title("Procesando archivos")
    ventana.geometry("600x400")
    label = Label(ventana, text="Procesando archivos, por favor espera...")
    label.pack(expand=False)
    ventana.grab_set()
    return ventana

def procesar_zip_y_guardar_excel(root,opcion):
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
        
        campos_deducciones = ["TIPO",
                            "Estado SAT",
                            "Version",
                            "Tipo",
                            "Fecha Emision",
                            "Fecha Timbrado",
                            "Serie",
                            "Folio",
                            "UUID",
                            "UUID Relacion",
                            "RFC Emisor",
                            "Nombre Emisor",
                            "LugarDeExpedicion",
                            "RFC Receptor",
                            "Nombre Receptor",
                            "UsoCFDI",
                            "SubTotal",
                            "Descuento",
                            "IVA 16%",
                            "Total",
                            "Moneda",
                            "FormaDePago",
                            "Metodo de Pago",
                            "Conceptos",
                            "OBSERVACIONES"]
        ventana_cargando = pantalla_progreso(root)

        wb = Workbook()
        ws = wb.active
        ws.title = "Datos XML"

        if opcion == "ambas":
            ws.append(["Folio Fiscal", "Folio", "Fecha", "Concepto", "Subtotal", "IVA", "Retención IVA", "Retención ISR", "Total", "Estado del CFDI","Tipo CFDI","Metodo Pago", "Forma Pago", "Uso CFDI"])
        elif opcion == "ret_iva":
            ws.append(["Folio Fiscal", "Folio", "Fecha", "Concepto", "Subtotal", "IVA", "Retención IVA", "Total", "Estado del CFDI","Tipo CFDI","Metodo Pago", "Forma Pago", "Uso CFDI"])
        elif opcion == "ret_isr":
            ws.append(["Folio Fiscal", "Folio", "Fecha", "Concepto", "Subtotal", "IVA", "Retención ISR", "Total", "Estado del CFDI","Tipo CFDI","Metodo Pago", "Forma Pago", "Uso CFDI"])
        elif opcion == "ninguna":
            ws.append(["Folio Fiscal", "Folio", "Fecha", "Concepto", "Subtotal", "IVA", "Total", "Estado del CFDI","Tipo CFDI","Metodo Pago", "Forma Pago", "Uso CFDI"])
        elif opcion == "nomina":
            ws.append(campos_nomina)
        elif opcion == "deducciones":
            ws.append(campos_deducciones)

        with zipfile.ZipFile(ruta_zip, 'r') as zip_ref:
            archivos_xml = [f for f in zip_ref.namelist() if f.endswith('.xml')]

            for nombre_archivo in archivos_xml:
                with zip_ref.open(nombre_archivo) as archivo:
                    archivo.name = nombre_archivo
                    
                    if(opcion == "nomina"):
                        datos = extraer_datos_nomina(archivo)
                        if datos:
                            fila = [str(datos.get(campo, "")) for campo in campos_nomina]
                            ws.append(fila)
                    elif(opcion == "deducciones"):
                        datos = extraer_datos_deducciones(archivo)
                        if datos:
                            fila = [str(datos.get(campo, "")) for campo in campos_deducciones]
                            ws.append(fila)
                    else:    
                        datos = extraer_datos_xml(archivo, opcion)
                        if datos:
                            ws.append(datos)

        ventana_cargando.destroy()

        ruta_salida = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if ruta_salida:
            wb.save(ruta_salida)
            wb.close()

        messagebox.showinfo("Éxito", f"Datos extraídos y guardados en Excel correctamente en la ruta {ruta_salida}")
        root.quit()
        
    threading.Thread(target=tarea).start()
    

def mostrar_ventana_principal():
    ventana = tk.Tk()
    ventana.title("Tipo de CFDI")
    ventana.geometry("300x300")

    tipo = tk.StringVar(value="ingresos")
    opcion = tk.StringVar(value="ninguna")

    ttk.Label(ventana, text="Selecciona el tipo de CFDI:").pack(pady=10)

    def mostrar_opciones():
        check_frame.pack()

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
    ttk.Radiobutton(check_frame, text="Nómina", variable=opcion, value="nomina").pack(anchor = "w")

    mostrar_opciones()

    def continuar():
        ventana.withdraw()
        if tipo.get() != "ingresos":
            opcion.set("ninguna")
        procesar_zip_y_guardar_excel(ventana, opcion.get())

    ttk.Button(ventana, text="Procesar Archivos", command=continuar).pack(pady=20)

    ventana.mainloop()

mostrar_ventana_principal()
