from zeep import Client

def esta_cancelado(uuid, rfc_emisor, rfc_receptor, total):
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
