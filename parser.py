import xml.etree.ElementTree as ET
from sat_client import esta_cancelado

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

def get_float_or_zero(val):
    if not val:
        return 0.0
    try:
        return float(val)
    except ValueError:
        return 0.0

def get_retIva(conceptos, namespaces):
    try:
        ret_iva_total = 0.00
        for concepto in conceptos:
            retenciones = concepto.findall('.//cfdi:Retencion', namespaces)
            for retencion in retenciones:
                if retencion.attrib.get('Impuesto') == '002':
                    importe = get_float_or_zero(retencion.attrib.get('Importe'))
                    ret_iva_total += importe
        return ret_iva_total
    except Exception:
        return 0.0

def get_retIsr(conceptos, namespaces):
    try:
        ret_isr_total = 0.00
        for concepto in conceptos:
            retenciones = concepto.findall('.//cfdi:Retencion', namespaces)
            for retencion in retenciones:
                if retencion.attrib.get('Impuesto') == '001':
                    importe = get_float_or_zero(retencion.attrib.get('Importe'))
                    ret_isr_total += importe
        return ret_isr_total
    except Exception:
        return 0.0

def extraer_datos_xml(archivo_xml, opcion):
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
        
        # Parse numeric inputs initially
        subtotal = get_float_or_zero(comprobante.attrib.get('SubTotal') or comprobante.attrib.get('subtotal'))
        total = get_float_or_zero(comprobante.attrib.get('Total') or comprobante.attrib.get('total'))
        
        tipoComprobante = comprobante.attrib.get('TipoDeComprobante')
        
        rfc_emisor = root.find('.//cfdi:Emisor', namespaces).attrib.get('Rfc')
        rfc_receptor = root.find('.//cfdi:Receptor', namespaces).attrib.get('Rfc')
        iva_total = 0.00
        
        # We need to determine the total for SAT cancellation query first
        if tipoComprobante == 'P':
            # For Payments, get total from pagos20 namespace if available
            pago20_ns = {'pago20': 'http://www.sat.gob.mx/Pagos20', 'cfdi': namespaces['cfdi']}
            totales_node = root.find('.//pago20:Totales', pago20_ns)
            if totales_node is not None:
                total = get_float_or_zero(totales_node.attrib.get('MontoTotalPagos'))
            else:
                total = 0.0
        else:
            total = get_float_or_zero(comprobante.attrib.get('Total') or comprobante.attrib.get('total'))

        estado = esta_cancelado(uuid, rfc_emisor, rfc_receptor, total)
        
        if tipoComprobante == 'P':
            pago20_ns = {'pago20': 'http://www.sat.gob.mx/Pagos20', 'cfdi': namespaces['cfdi']}
            pago_node = root.find('.//pago20:Pago', pago20_ns)
            totales_node = root.find('.//pago20:Totales', pago20_ns)
            
            fecha = pago_node.attrib.get('FechaPago') if pago_node is not None else fecha
            metodoPago = "PPD"  # Typically Pago is PPD (Pago en parcialidades o diferido), or we can leave as "P"
            formaPago = get_forma_pago(str(pago_node.attrib.get('FormaDePagoP'))) if pago_node is not None else "NA"
            moneda = pago_node.attrib.get('MonedaP') or (pago_node.attrib.get('Moneda') if pago_node is not None else 'MXN')
            
            if totales_node is not None:
                subtotal = get_float_or_zero(totales_node.attrib.get('TotalTrasladosBaseIVA16'))
                iva_total = get_float_or_zero(totales_node.attrib.get('TotalTrasladosImpuestoIVA16'))
                retencion_iva = get_float_or_zero(totales_node.attrib.get('TotalRetencionesIVA'))
                retencion_isr = get_float_or_zero(totales_node.attrib.get('TotalRetencionesISR'))
                total = get_float_or_zero(totales_node.attrib.get('MontoTotalPagos'))
            else:
                subtotal = 0.0
                iva_total = 0.0
                retencion_iva = 0.0
                retencion_isr = 0.0

            # Get all DoctoRelacionado IdDocumento values
            docto_relacionados = root.findall('.//pago20:DoctoRelacionado', pago20_ns)
            id_docs = [doc.attrib.get('IdDocumento') for doc in docto_relacionados if doc.attrib.get('IdDocumento')]
            concepto_str ="Folios pagados: " +  ",".join(id_docs)
            
            receptor_node = root.find('.//cfdi:Receptor', namespaces)
            uso = get_uso_cfdi(str(receptor_node.attrib.get('UsoCFDI'))) if receptor_node is not None else "NA"
        else:
            metodoPago = comprobante.attrib.get('MetodoPago')
            formaPago = get_forma_pago(str(comprobante.attrib.get('FormaPago')))
            
            receptor_node = root.find('.//cfdi:Receptor', namespaces)
            uso = get_uso_cfdi(str(receptor_node.attrib.get('UsoCFDI'))) if receptor_node is not None else "NA"
            moneda = comprobante.attrib.get('Moneda') or comprobante.attrib.get('moneda')

            conceptos = root.findall('.//cfdi:Concepto', namespaces)

            for concepto in conceptos:
                traslados = concepto.findall('.//cfdi:Traslado', namespaces)
                for traslado in traslados:
                    if traslado.attrib.get('Impuesto') == '002':  # IVA
                        importe = get_float_or_zero(traslado.attrib.get('Importe', '0'))
                        iva_total += importe
            
            concepto_str = ""
            if conceptos:
                concepto_str = ",".join([con.attrib.get('Descripcion', con.attrib.get('descripcion', '')) for con in conceptos])

            retencion_iva = get_retIva(conceptos, namespaces)
            retencion_isr = get_retIsr(conceptos, namespaces)

        # Check for Egreso (type E) to invert numeric values to negative
        multiplier = 1.0
        if tipoComprobante == "E":
            multiplier = -1.0
            subtotal *= multiplier
            total *= multiplier
            iva_total *= multiplier
            retencion_iva *= multiplier
            retencion_isr *= multiplier

        if opcion == "ambas":
            return [nombre, folio, fecha, concepto_str, subtotal, iva_total, retencion_iva, retencion_isr, total, estado, tipoComprobante, metodoPago, formaPago, uso, moneda, rfc_emisor]
        elif opcion == "ret_isr":
            return [nombre, folio, fecha, concepto_str, subtotal, iva_total, retencion_isr, total, estado, tipoComprobante, metodoPago, formaPago, uso, moneda, rfc_emisor]
        elif opcion == "ret_iva":
            return [nombre, folio, fecha, concepto_str, subtotal, iva_total, retencion_iva, total, estado, tipoComprobante, metodoPago, formaPago, uso, moneda, rfc_emisor]
        elif opcion == "ninguna":
            return [nombre, folio, fecha, concepto_str, subtotal, iva_total, total, estado, tipoComprobante, metodoPago, formaPago, uso, moneda, rfc_emisor]

    except Exception as e:
        print(f"Error al procesar archivo XML: {e}")
        return None

def extraer_datos_nomina(archivo_xml):
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
        
        rfc_emisor = root.find('.//cfdi:Emisor', namespaces).attrib.get('Rfc')
        rfc_receptor = root.find('.//cfdi:Receptor', namespaces).attrib.get('Rfc')
        total_cfdi = get_float_or_zero(comprobante.attrib.get('Total'))
        
        estado_sat = esta_cancelado(uuid, rfc_emisor, rfc_receptor, total_cfdi)

        datos = {
            "EstadoSAT": estado_sat,
            "FechaEmision": comprobante.attrib.get("Fecha"),
            "FechaTimbrado": root.find('.//tfd:TimbreFiscalDigital', namespaces).attrib.get("FechaTimbrado"),
            "Serie": comprobante.attrib.get("Serie"),
            "Folio": comprobante.attrib.get("Folio"),
            "UUID": uuid,
            "RFC Receptor": rfc_receptor,
            "NombreReceptor": root.find('.//cfdi:Receptor', namespaces).attrib.get("Nombre"),
            "RFC Emisor": rfc_emisor,
            "NombreEmisor": root.find('.//cfdi:Emisor', namespaces).attrib.get("Nombre"),
        }

        nomina = root.find('.//nomina12:Nomina', namespaces)
        receptor_n = nomina.find('.//nomina12:Receptor', namespaces)
        percepciones = nomina.find('.//nomina12:Percepciones', namespaces)
        deducciones = nomina.find('.//nomina12:Deducciones', namespaces)

        # Inversion modifier check
        tipoComprobante = comprobante.attrib.get("TipoDeComprobante")
        multiplier = -1.0 if tipoComprobante == "E" else 1.0

        datos.update({
            "RegistroPatronal": nomina.find('.//nomina12:Emisor', namespaces).attrib.get("RegistroPatronal"),
            "TipoNomina": nomina.attrib.get("TipoNomina"),
            "FechaPago": nomina.attrib.get("FechaPago"),
            "FechaInicialPago": nomina.attrib.get("FechaInicialPago"),
            "FechaFinalPago": nomina.attrib.get("FechaFinalPago"),
            "NumDiasPagados": get_float_or_zero(nomina.attrib.get("NumDiasPagados")),
            "TotalPercepciones": get_float_or_zero(nomina.attrib.get("TotalPercepciones")) * multiplier,
            "TotalDeducciones": get_float_or_zero(nomina.attrib.get("TotalDeducciones")) * multiplier,
            "TotalOtrosPagos": get_float_or_zero(nomina.attrib.get("TotalOtrosPagos")) * multiplier,
            "SubTotal": get_float_or_zero(nomina.attrib.get("Subtotal")) * multiplier,
            "Descuento": get_float_or_zero(nomina.attrib.get("TotalDeducciones")) * multiplier,
            "Total": total_cfdi * multiplier,
            "MetodoPago": comprobante.attrib.get("MetodoPago"),
            "Regimen": receptor_n.attrib.get("TipoRegimen"),
            "ArchivoXML": nombre_archivo,
            "Conceptos": root.find('.//cfdi:Concepto', namespaces).attrib.get("Descripcion"),
            "TipoComprobante": tipoComprobante,
            "Version": comprobante.attrib.get("Version"),
            "Moneda": comprobante.attrib.get("Moneda"),
            "ReceptorCurp": receptor_n.attrib.get("Curp"),
            "NumSeguridadSocial": receptor_n.attrib.get("NumSeguridadSocial"),
            "FechaInicioRelLaboral": receptor_n.attrib.get("FechaInicioRelLaboral"),
            "TotalSueldosPer": get_float_or_zero(percepciones.attrib.get("TotalSueldos")) * multiplier if percepciones is not None else 0.0,
            "TotalGravadoPercepcion": get_float_or_zero(percepciones.attrib.get("TotalGravado")) * multiplier if percepciones is not None else 0.0,
            "TotalExentoPercepcion": get_float_or_zero(percepciones.attrib.get("TotalExento")) * multiplier if percepciones is not None else 0.0,
            "TotalOtrasDeducciones": get_float_or_zero(deducciones.attrib.get("TotalOtrasDeducciones")) * multiplier if deducciones is not None else 0.0,
            "TotalImpuestosRetenidosDed": get_float_or_zero(deducciones.attrib.get("TotalImpuestosRetenidos")) * multiplier if deducciones is not None else 0.0,
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
            datos[gra] = 0.00
            if exe:
                datos[exe] = 0.00

        if percepciones is not None:
            for per in percepciones.findall('.//nomina12:Percepcion', namespaces):
                tipo = per.attrib.get('TipoPercepcion')
                if tipo in percepcion_map:
                    gra, exe = percepcion_map[tipo]
                    datos[gra] = get_float_or_zero(per.attrib.get('ImporteGravado', '0.00')) * multiplier
                    if exe:
                        datos[exe] = get_float_or_zero(per.attrib.get('ImporteExento', '0.00')) * multiplier

        # Mapear deducciones específicas
        deduccion_map = {
            '001': 'D01_SeguroSocial',
            '002': 'D02_ISR',
            '004': 'D04_OtrasDeducciones',
            '010': 'D10_CreditoVivienda'
        }

        for val in deduccion_map.values():
            datos[val] = 0.00

        if deducciones is not None:
            for ded in deducciones.findall('.//nomina12:Deduccion', namespaces):
                tipo = ded.attrib.get('TipoDeduccion')
                if tipo in deduccion_map:
                    datos[deduccion_map[tipo]] = get_float_or_zero(ded.attrib.get('Importe', '0.00')) * multiplier

        return datos

    except Exception as e:
        print(f"Error al procesar XML de nómina: {e}")
        return None

def extraer_datos_deducciones(archivo_xml):
    try:
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

        iva_total = 0.0
        for concepto in conceptos:
            for traslado in concepto.findall('.//cfdi:Traslado', namespaces):
                if traslado.attrib.get('Impuesto') == '002':  # IVA 16%
                    iva_total += get_float_or_zero(traslado.attrib.get('Importe', '0'))

        tipoComprobante = comprobante.attrib.get('TipoDeComprobante', '')
        multiplier = -1.0 if tipoComprobante == "E" else 1.0

        subtotal = get_float_or_zero(comprobante.attrib.get('SubTotal', '0')) * multiplier
        descuento = get_float_or_zero(comprobante.attrib.get('Descuento', '0')) * multiplier
        iva_total = round(iva_total, 2) * multiplier
        total = get_float_or_zero(comprobante.attrib.get('Total', '0')) * multiplier

        datos = {
            "TIPO": tipoComprobante,
            "Estado SAT": "Pendiente o Cancelado (requiere consulta al SAT)",
            "Version": comprobante.attrib.get('Version', ''),
            "Tipo": tipoComprobante,
            "Fecha Emision": comprobante.attrib.get('Fecha', ''),
            "Fecha Timbrado": timbre.attrib.get('FechaTimbrado', '') if timbre is not None else '',
            "Serie": comprobante.attrib.get('Serie', ''),
            "Folio": comprobante.attrib.get('Folio', ''),
            "UUID": timbre.attrib.get('UUID', '') if timbre is not None else '',
            "UUID Relacion": "",  
            "RFC Emisor": emisor.attrib.get('Rfc', '') if emisor is not None else '',
            "Nombre Emisor": emisor.attrib.get('Nombre', '') if emisor is not None else '',
            "LugarDeExpedicion": comprobante.attrib.get('LugarExpedicion', ''),
            "RFC Receptor": receptor.attrib.get('Rfc', '') if receptor is not None else '',
            "Nombre Receptor": receptor.attrib.get('Nombre', '') if receptor is not None else '',
            "UsoCFDI": receptor.attrib.get('UsoCFDI', '') if receptor is not None else '',
            "SubTotal": subtotal,
            "Descuento": descuento,
            "IVA 16%": iva_total,
            "Total": total,
            "Moneda": comprobante.attrib.get('Moneda', ''),
            "FormaDePago": comprobante.attrib.get('FormaPago', ''),
            "Metodo de Pago": comprobante.attrib.get('MetodoPago', ''),
            "Conceptos": " | ".join(conceptos_list),
            "OBSERVACIONES": ""
        }

        return datos
    except Exception as e:
        print(f"Error al procesar deducciones: {e}")
        return None
