Attribute VB_Name = "ConstantesNegocio"
Option Explicit
'*** Acceso ****
Public Const Clave_Registro_Sistema = "Software\TAM Consulting\Spectrum Fondos CX"
Public Const Clave_Registro_Sistema_Seguridad = "Software\TAM Consulting\Seguridad CX"
Public Const Clave_Registro_TipoAdministradora = "Tipo Administradora"
Public Const Clave_Registro_TipoAdministradoraContable = "Tipo Administradora Contable"
Public Const Clave_Registro_Administradora = "Administradora"
Public Const Clave_Registro_AdministradoraContable = "Administradora Contable"
Public Const Clave_Registro_NombreSistema = "Sistema"
Public Const Clave_Registro_Version = "Version"
Public Const Clave_Registro_Empresa = "Empresa"
Public Const Clave_Registro_Servidor = "Servidor"
Public Const Clave_Registro_BaseDatos = "Base de Datos"
Public Const Clave_Registro_RutaReportes = "Ruta Reportes"
Public Const Clave_Registro_RutaLogoTAM = "Ruta Logo"
Public Const Clave_Registro_RutaTemporal = "Ruta Temporal"
Public Const Clave_Registro_RutaBackups = "Ruta Backups"
Public Const Clave_Registro_RutaCargaOperaciones = "Ruta Carga Operaciones"
Public Const Clave_Registro_FormatoFechaCliente = "Formato Fecha Cliente"
Public Const Clave_Registro_FormatoFechaInterno = "Formato Fecha Interno"
Public Const Clave_Registro_Cliente = "Cliente"

'*** Estados Mantenimiento ***
Public Const Reg_Adicion = "ADICION"
Public Const Reg_Edicion = "EDICION"
Public Const Reg_Consulta = "CONSULTA"
Public Const Reg_Eliminacion = "ELIMINACION" 'acr : 06-09-2009
Public Const Reg_Defecto = ""

'*** Tipos de Campos *** ACR: 19/10/2013
Public Const Tipo_Campo_Input = "01"
Public Const Tipo_Campo_Output = "02"
Public Const Tipo_Campo_Filtro = "03"
Public Const Tipo_Campo_Input_Output = "04"

'*** Estado de un Registro ***
Public Const Modo_Cobro_Interes_Vencimiento = "01"
Public Const Modo_Cobro_Interes_Adelantado = "02"

'*** Estado de un Registro ***
Public Const Estado_Activo = "01"
Public Const Estado_Inactivo = "02"
Public Const Estado_Eliminado = "03"

'*** Estado de una Solicitud de Participacion
Public Const Estado_Solicitud_Ingresada = "01"
Public Const Estado_Solicitud_Confirmada = "02"
Public Const Estado_Solicitud_Anulada = "03"
Public Const Estado_Solicitud_Liquidada = "04"
Public Const Estado_Solicitud_Procesada = "05"

'*** Estado de una Solicitud de Flujos ***
Public Const Estado_Solicitud_Flujos_Ingresada = "01"
Public Const Estado_Solicitud_Flujos_EnProceso = "02"
Public Const Estado_Solicitud_Flujos_Aprobada = "03"
Public Const Estado_Solicitud_Flujos_Formalizada = "04"
Public Const Estado_Solicitud_Flujos_Anulada = "05"
Public Const Estado_Solicitud_Flujos_Denegada = "06"
Public Const Estado_Solicitud_Flujos_Observada = "07"

'*** Estado de Registro de C/V ***
Public Const Estado_Registro_Ingresado = "01"
Public Const Estado_Registro_Anulado = "03"
Public Const Estado_Registro_Contabilizado = "04"

'*** Tipos de desplazamiento ***
Public Const Tipo_Desplazamiento_Ningun_Desplazamiento = "00"
Public Const Tipo_Desplazamiento_Siguiente_Dia_Laborable = "01"
Public Const Tipo_Desplazamiento_Siguiente_laborable_modificado = "02"
Public Const Tipo_Desplazamiento_Dia_Laborable_Anterior = "03"
Public Const Tipo_Desplazamiento_Dia_Laborable_Anterior_Modificado = "04"


'**** Tipo Impuesto
Public Const Codigo_Impuesto_IGV = "01"

'*** Tipo de Documento
Public Const Codigo_Documento_Emitido = "01"
Public Const Codigo_Documento_No_Emitido = "02"

'*** Tipos de calculo de dias ***
Public Const Tipo_Calculo_Dias_Base_Anual = "01"
Public Const Tipo_Calculo_Dias_Diferencia = "02"

'*** Tipos de calculo ***
Public Const Tipo_Calculo_Fijo = "01"
Public Const Tipo_Calculo_Variable = "02"

'*** Tipos de Formula ***
Public Const Tipo_Formula_Fija = "03"

'*** Estado de un Certificado ***
Public Const Estado_Certificado_Vigente = "01"
Public Const Estado_Certificado_NoVigente = "02"

'*** Tipos de �rdenes de inversi�n ***
Public Const Tipo_Orden_DepositoPlazo = "003"

'*** CodFile de Operaciones de Acreencias ***
Public Const CodFile_Descuento_Comprobantes_Pago = "014"
Public Const CodFile_Descuento_Documentos_Cambiario = "015"
Public Const CodFile_Descuento_Flujos_Dinerarios = "016"

'*** CodFile de Operaciones de Financiamiento ***
Public Const CodFile_Financiamiento_Prestamos = "070"

'*** Estado de una Orden de Inversi�n ***
Public Const Estado_Orden_Anulada = "01"
Public Const Estado_Orden_Enviada = "02"
Public Const Estado_Orden_Ingresada = "03"
Public Const Estado_Orden_Procesada = "04"
Public Const Estado_Orden_PorAutorizar = "05"
Public Const Estado_Orden_Autorizada = "06"

'*** Estado de un Acuerdo (Evento) ***
Public Const Estado_Acuerdo_Ingresado = "01"
Public Const Estado_Acuerdo_Procesado = "02"
Public Const Estado_Acuerdo_Anulado = "03"

'*** Estado de una Entrega (Evento) ***
Public Const Estado_Entrega_Generado = "01"
Public Const Estado_Entrega_Procesado = "02"
Public Const Estado_Entrega_Anulado = "03"

'*** Estado de Orden CobroPago ***
Public Const Estado_Caja_NoConfirmado = "01"
Public Const Estado_Caja_Confirmado = "02"
Public Const Estado_Caja_Anulado = "03"

'*** Valores Selecci�n ***
Public Const Sel_Todos = "[ TODOS ]"
Public Const Sel_Defecto = "[ NO SELECCIONADO ]"
Public Const Sel_NoAplicable = "[ NO APLICABLE ]" 'NUEVO ACR

'*** C�digo Tipo Comisionista *** NUEVO ACR
Public Const Codigo_Tipo_Comisionista_Participe = "01"
Public Const Codigo_Tipo_Comisionista_Inversion = "02"

'*** Valores Nulos ***
Public Const Valor_Fecha = "01/01/1900"
Public Const Valor_Fecha_Fin = "31/12/9999"
Public Const Valor_Numero = 0
Public Const Valor_Caracter = ""
Public Const Valor_Indicador = "X"
Public Const Valor_Comodin = "%"

'*** Tratamiento Contable de Comisiones e Impuestos - Operaciones de Inversion
Public Const Valor_Tratamiento_Contable_Costo = "COSTO"
Public Const Valor_Tratamiento_Contable_Gasto = "GASTO"
Public Const Valor_Tratamiento_Contable_Credito = "CREDITO"

'*** Tipo Aporte ***
Public Const Tipo_Aporte_Dinerario = "01"
Public Const Tipo_Aporte_NoDinerario = "02"

'*** Valores Par�metros Fondo ***
Public Const Valor_NumOperacion = "NUMOPE"
Public Const Valor_NumOpeCertificado = "NUMOCR"
Public Const Valor_NumSolicitud = "NUMSOL"
Public Const Valor_NumComprobante = "NUMCOM"
Public Const Valor_NumCertificado = "NUMCER"
Public Const Valor_NumOrdenCaja = "NUMCAJ"
Public Const Valor_NumInt = "NUMINT"
Public Const Valor_NumOrdenInversion = "NUMORD"
Public Const Valor_NumKardex = "NUMKAR"
Public Const Valor_NumEntregaEvento = "NUMENT"
Public Const Valor_NumCobertura = "NUMCOB"
Public Const Valor_NumOpeCajaBancos = "NUMOPC" 'modifica ACR: 27/01/2009
Public Const Valor_NumOpeTesoreria = "NUMTES"
Public Const Valor_NumRegistroCompra = "NUMREC"
Public Const Valor_NumPrestamo = "NUMPRE"
Public Const Valor_NumOrdenPago = "NUMORP"
Public Const Valor_NumComprobantePago = "NUMCOP"

'*** Tipo de File ***
Public Const Valor_File_Inversiones = "01"
Public Const Valor_File_Ingresos = "02"
Public Const Valor_File_Gastos = "03"
Public Const Valor_File_CajaBancos = "04"
Public Const Valor_File_Generico = "05" 'modifica ACR: 27/01/2009



'Nuevo ACR: 27/01/2009
Public Enum ParametrosFondo
    PARAM_INICIA = 1
    PARAM_NUMOPE = PARAM_INICIA
    PARAM_NUMOCR
    PARAM_NUMSOL
    PARAM_NUMCOM
    PARAM_NUMCER
    PARAM_NUMCAJ
    PARAM_NUMINT
    PARAM_NUMORD
    PARAM_NUMKAR
    PARAM_NUMENT
    PARAM_NUMCOB
    PARAM_NUMOPC
    PARAM_NUMREC
    PARAM_NUMPRE
    PARAM_NUMORP
    PARAM_NUMCOP
    PARAM_FINAL = PARAM_NUMOPC
End Enum

Public astrParametrosFondo() As String

'*** Valores Mensajes ***
Public Const Mensaje_Adicion_Exitosa = "Se guardaron los datos en forma exitosa..."
Public Const Mensaje_Edicion_Exitosa = "Se actualizaron los datos en forma exitosa..."
Public Const Mensaje_Eliminacion_Exitosa = "Se eliminaron los datos en forma exitosa..."
Public Const Mensaje_Adicion = "Seguro de adicionar la informaci�n ?"
Public Const Mensaje_Edicion = "Seguro de actualizar la informaci�n ?"
Public Const Mensaje_Eliminacion = "Seguro de eliminar la informaci�n ?"
Public Const Mensaje_Proceso_Exitoso = "El proceso termin� en forma exitosa..."
Public Const Mensaje_Proceso_NoExitoso = "El proceso tuvo errores, se procedi� a revertir..."
Public Const Mensaje_Confirmacion_Exitoso = "El proceso de confirmaci�n termin� en forma exitosa..."
Public Const Mensaje_Desconfirmacion_Exitoso = "El proceso de desconfirmaci�n termin� en forma exitosa..."
Public Const Mensaje_Envio_Exitoso = "El Proceso de Env�o termin� en forma exitosa..."
Public Const Mensaje_Desenvio_Exitoso = "El Proceso de Desenv�o termin� en forma exitosa..."
Public Const Mensaje_Adicion_Periodo = "Seguro de generar el nuevo periodo contable ?"
Public Const Mensaje_Adicion_Periodo_Exitoso = "Se gener� el nuevo periodo contable en forma exitosa..."
Public Const Mensaje_Registro_Duplicado = "Registro ya existe..."
Public Const Mensaje_Error_Inesperado = "Error inesperado..."
Public Const Mensaje_Carga_Exitosa = "Proceso de carga exitoso..."
Public Const Mensaje_Modificar = "Se envio la peticion de Modificacion.."


'*** N�mero de Decimales ***
Public Const Decimales_Monto = 2
Public Const Decimales_ValorCuota = 5
Public Const Decimales_ValorCuota_Cierre = 8
Public Const Decimales_CantCuota = 5
Public Const Decimales_Tasa = 6
Public Const Decimales_Tasa2 = 2
Public Const Decimales_Precio = 9
Public Const Decimales_TasaDiaria = 12
Public Const Decimales_TipoCambio = 12

'*** Valor Tipo de Cambio ***
Public Const Valor_TipoCambio_Compra = "COMPRA"
Public Const Valor_TipoCambio_Venta = "VENTA"
Public Const Codigo_TipoCambio_Compra = "01"
Public Const Codigo_TipoCambio_Venta = "02"

'*** C�digo Tipo Administradora ***
Public Const Codigo_Tipo_Fondo_Mutuo = "01"
Public Const Codigo_Tipo_Fondo_Inversion = "02"
Public Const Codigo_Tipo_Fondo_Portafolio = "03"
Public Const Codigo_Tipo_Fondo_Administradora = "04"

'*** C�digo Tipo Fondo ***
Public Const Codigo_Fondo_Abierto = "01"
Public Const Codigo_Fondo_Cerrado = "02"
Public Const Administradora_Fondos = "03"

'*** C�digo Tipo Frecuencia ***
Public Const Codigo_Tipo_Frecuencia_Anual = "01"
Public Const Codigo_Tipo_Frecuencia_Semestral = "02"
Public Const Codigo_Tipo_Frecuencia_Trimestral = "03"
Public Const Codigo_Tipo_Frecuencia_Bimestral = "04"
Public Const Codigo_Tipo_Frecuencia_Mensual = "05"
Public Const Codigo_Tipo_Frecuencia_Quincenal = "06"
Public Const Codigo_Tipo_Frecuencia_Diaria = "07"

'*** Codigo de Asiento TIPASI ****
Public Const Codigo_Tipo_Asiento_Apertura_Cierre = "00"
Public Const Codigo_Tipo_Asiento_Gasto = "01"
Public Const Codigo_Tipo_Asiento_Inversion = "02"
Public Const Codigo_Tipo_Asiento_Cobranza = "03"
Public Const Codigo_Tipo_Asiento_Caja_Chica = "04"
Public Const Codigo_Tipo_Asiento_Diario = "05"
Public Const Codigo_Tipo_Asiento_Planilla = "06"
Public Const Codigo_Tipo_Asiento_Regularizacion_Intereses = "07"
Public Const Codigo_Tipo_Asiento_Operaciones_Participes_Valor_Conocido = "08"
Public Const Codigo_Tipo_Asiento_Operaciones_Participes_Valor_Desconocido = "09"
Public Const Codigo_Tipo_Asiento_Bancos_Cargo = "10"
Public Const Codigo_Tipo_Asiento_Bancos_Abono = "11"
Public Const Codigo_Tipo_Asiento_Distribucion_Utilidades_Reparto_Reinversion = "12"
Public Const Codigo_Tipo_Asiento_Distribucion_Utilidades_Reversion = "13"
Public Const Codigo_Tipo_Asiento_Provision_Comisiones_Promotores = "14"
Public Const Codigo_Tipo_Asiento_Provision_Gastos_Proveedores = "15"
Public Const Codigo_Tipo_Asiento_Provision_Comisiones_SAF = "16"
Public Const Codigo_Tipo_Asiento_Valorizacion_Inversiones = "17"
Public Const Codigo_Tipo_Asiento_Ajuste_Diferencia_Cambio = "18"
Public Const Codigo_Tipo_Asiento_Ajuste_Traslacion = "19"
Public Const Codigo_Tipo_Asiento_Resultados_Ejercicio = "20"
Public Const Codigo_Tipo_Asiento_Apertura_Periodo_Contable = "21"
Public Const Codigo_Tipo_Asiento_Cierre_Periodo_Contable = "22"
Public Const Codigo_Tipo_Asiento_Devolucion = "23"
Public Const Codigo_Tipo_Asiento_Automaticos = "99"

'*** Estructuras XML ***
Public Const XML_MontoMovimientoContable = "<MontoMovimientoContable />"
Public Const XML_TipoCambioReemplazo = "<TipoCambioReemplazo />"
Public Const XML_MonedaContable = "<MonedaContable />"

'*** C�digo Tipo Remunerada ***
Public Const Codigo_Tipo_Remunerada_Monto = "01"

'*** C�digo Comprobante Pago
Public Const Codigo_Tipo_Comprobante_Pago_Poliza = "11"

'*** C�digo Tipo Cuenta Fondo ***
Public Const Codigo_Tipo_Cuenta_Ahorro = "02"
Public Const Codigo_Tipo_Cuenta_Corriente = "01"

'*** C�digo Tipo Direcci�n Postal ***
Public Const Codigo_Tipo_Direccion_Domicilio = "01"
Public Const Codigo_Tipo_Direccion_Oficina = "02"
Public Const Codigo_Tipo_Direccion_Otro = "03"
Public Const Codigo_Tipo_Direccion_Retencion = "04"

'*** C�digo Tipo Documento Identidad *** 'Agregado ACR
Public Const Codigo_Tipo_Otro_Documento_Natural = "00"
Public Const Codigo_Tipo_Documento_Nacional_Identidad = "01"
Public Const Codigo_Tipo_Carnet_Identidad_Militar = "02"
Public Const Codigo_Tipo_Ficha_Registro_Publico = "03"
Public Const Codigo_Tipo_Carnet_Extranjeria = "04"
Public Const Codigo_Tipo_Registro_Conasev = "05"
Public Const Codigo_Tipo_Registro_Unico_Contribuyente = "06"
Public Const Codigo_Tipo_Pasaporte = "07"
Public Const Codigo_Tipo_Registro_SAFP = "08"
Public Const Codigo_Tipo_Otro_Documento_Juridico = "21"
Public Const Codigo_Tipo_Numero_Participe = "30"

'*** C�digo Tipo Mancomuno ***
Public Const Codigo_Tipo_Contrato_Individual = "01"
Public Const Codigo_Tipo_Contrato_Mancomuno = "02"

'*** C�digo Tipo Mancomuno ***
Public Const Codigo_Tipo_Mancomuno_Individual = "01"
Public Const Codigo_Tipo_Mancomuno_Indistinto = "03"
Public Const Codigo_Tipo_Mancomuno_Conjunto = "02"

'*** C�digo Base Anual ***
Public Const Codigo_Base_Actual_Actual = "01"
Public Const Codigo_Base_Actual_360 = "02"
Public Const Codigo_Base_Actual_365 = "03"
Public Const Codigo_Base_30_360 = "04"
Public Const Codigo_Base_30_365 = "05"

'*** C�digo Forma de C�lculo Intereses ***
Public Const Codigo_Calculo_Normal = "01"
Public Const Codigo_Calculo_Prorrateo = "02"

'*** C�digo Frecuencia ***
Public Const Codigo_Frecuencia_Anual = "01"
Public Const Codigo_Frecuencia_Semestral = "02"
Public Const Codigo_Frecuencia_Trimestral = "03"
Public Const Codigo_Frecuencia_Bimestral = "04"
Public Const Codigo_Frecuencia_Mensual = "05"
Public Const Codigo_Frecuencia_Quincenal = "06"
Public Const Codigo_Frecuencia_Diaria = "07"

'*** C�digo Modalidad de Pago ***
Public Const Codigo_Modalidad_Pago_Vencimiento = "01"
Public Const Codigo_Modalidad_Pago_Adelantado = "02"

'*** C�digo Tipo de Pago ***
Public Const Codigo_Tipo_Pago_Periodico = "01"
Public Const Codigo_Tipo_Pago_Unico = "02"

'*** C�digo de Moneda ***
Public Const Codigo_Moneda_Local = "01"
Public Const Codigo_Moneda_Dolar_Americano = "02"
Public Const Codigo_Moneda_Extranjero = "02"

'*** Signo Moneda ***
Public Const Signo_Moneda_Local = "PEN"
Public Const Signo_Moneda_Dolar_Americano = "USD"


'*** C�digo Tipo Movimiento Cuenta Fondo ***
Public Const Codigo_Movimiento_Deposito = "01"
Public Const Codigo_Movimiento_Retiro = "02"
Public Const Codigo_Movimiento_Deposito_No_Identificado = "04"



'*** C�digo Tipo de Operaci�n Captaci�n ***
Public Const Codigo_Operacion_Suscripcion = "01"
Public Const Codigo_Operacion_Rescate = "02"
Public Const Codigo_Operacion_Transferencia = "03"

'*** C�digo Clase Operaci�n Captaci�n ***
Public Const Codigo_Clase_SuscripcionConocida = "01"
Public Const Codigo_Clase_SuscripcionDesconocida = "02"
Public Const Codigo_Clase_RescateTotalConocido = "03"
Public Const Codigo_Clase_RescateTotalDesconocido = "05"
Public Const Codigo_Clase_RescateParcialConocido = "04"
Public Const Codigo_Clase_RescateParcialDesconocido = "06"
Public Const Codigo_Clase_TransferenciaTotal = "07"
Public Const Codigo_Clase_TransferenciaParcial = "08"

'*** C�digo Clase Persona *** ' cambios ACR
Public Const Codigo_Persona_Natural = "01"
Public Const Codigo_Persona_Juridica = "02"
Public Const Codigo_Persona_Mancomuno = "03"

'*** C�digo Tipo Respuesta ***
Public Const Codigo_Respuesta_Si = "01"
Public Const Codigo_Respuesta_No = "02"

'*** C�digo Tipo Cuenta Contable ***
Public Const Codigo_Tipo_Cuenta_Cuenta = "01"
Public Const Codigo_Tipo_Cuenta_Grupo = "02"

'*** C�digo Tipo Naturaleza Cuenta Contable ***
Public Const Codigo_Tipo_Naturaleza_Debe = "D"
Public Const Codigo_Tipo_Naturaleza_Haber = "H"

'*** C�digo Mercado Local Extranjero ***
Public Const Codigo_Mercado_Local = "01" 'ACR: 24/05/2012
Public Const Codigo_Mercado_Extranjero = "02" 'ACR: 24/05/2012

'*** C�digo Din�mica Contable *** ' OJO : Estos codigos corresponden tambien a tabla virtual OPECAJ
Public Const Codigo_Dinamica_Compra = "01"
Public Const Codigo_Dinamica_Venta = "02"
Public Const Codigo_Dinamica_Vencimiento = "03"
Public Const Codigo_Dinamica_Apertura = "04"
Public Const Codigo_Dinamica_Cierre = "05"
Public Const Codigo_Dinamica_Cupon = "06"
Public Const Codigo_Dinamica_Renovacion = "07"
Public Const Codigo_Dinamica_Dividendos = "08"
Public Const Codigo_Dinamica_Provision = "09"
Public Const Codigo_Dinamica_PrePago = "10"
Public Const Codigo_Dinamica_Comision = "18"
Public Const Codigo_Dinamica_Gasto = "19"
Public Const Codigo_Dinamica_Impuesto = "20"
Public Const Codigo_Dinamica_Detraccion = "21" 'ACR:28-11-2008
Public Const Codigo_Dinamica_Retencion = "22" 'ACR:27-05-2009

Public Const Codigo_Dinamica_Detraccion_Ajuste_Redondeo_Ganancia = "23" 'ACR:28-11-2008
Public Const Codigo_Dinamica_Detraccion_Ajuste_Redondeo_Perdida = "24" 'ACR:27-05-2009
Public Const Codigo_Dinamica_Gasto_Emitida = "25"   'JJCC: 20-05-2012
Public Const Codigo_Dinamica_Facturacion = "26"   'JJCC: 20-05-2012
Public Const Codigo_Dinamica_Dividendos_Percibidos = "27"   'ACR: 24-05-2012

'JAFR 2/07/2014
Public Const Codigo_Dinamica_ProvisionInteresesDiferido = "28"
Public Const Codigo_Dinamica_ProvisionInteresesAdicionales = "29"
Public Const Codigo_Dinamica_DevolucionIntereses = "30"
Public Const Codigo_Dinamica_DevolucionMargenCobertura = "31"
Public Const Codigo_Dinamica_ProvisionGananciaPerdidaCapital = "32"
Public Const Codigo_Dinamica_CargaDiferidaGasto = "33"
Public Const Codigo_Dinamica_ContabilizacionComprobantePago = "34"
Public Const Codigo_Dinamica_Facturacion_Vcto_Operacion = "35"
Public Const Codigo_Dinamica_Facturacion_Intereses_Adicionales = "36"
Public Const Codigo_Dinamica_Facturacion_Intereses_Adelantados = "37"
Public Const Codigo_Dinamica_Registro_Intereses_X_Devengar = "38"
Public Const Codigo_Dinamica_Desembolso = "39"
Public Const Codigo_Dinamica_ContabilizacionNotaCredito = "40"

'*** C�digo Tipo de Cierre ***
Public Const Codigo_Cierre_Definitivo = "0"
Public Const Codigo_Cierre_Simulacion = "1"

'*** C�digo Tipo de Proceso Contable ***
Public Const Codigo_Proceso_Apertura_Periodo_Contable = "0"
Public Const Codigo_Proceso_Cierre_Periodo_Contable = "1"

'*** C�digo Tipo Asignaci�n Valor Cuota ***
Public Const Codigo_Asignacion_TMenos1 = "01"
Public Const Codigo_Asignacion_T = "02"
Public Const Codigo_Asignacion_TMas1 = "03"

'*** C�digo Tipo Persona ***
Public Const Codigo_Tipo_Persona_Relacionado = "01"
Public Const Codigo_Tipo_Persona_Emisor = "02"
Public Const Codigo_Tipo_Persona_Agente = "03"
Public Const Codigo_Tipo_Persona_Proveedor = "04"
Public Const Codigo_Tipo_Persona_Cliente = "05"
Public Const Codigo_Tipo_Persona_Participe = "06"
Public Const Codigo_Tipo_Persona_Portafolio = "07"
Public Const Codigo_Tipo_Persona_Organismo = "08"
Public Const Codigo_Tipo_Persona_Contratante = "09"
Public Const Codigo_Tipo_Persona_Comisionista = "10"
Public Const Codigo_Tipo_Persona_Obligado = "11"

'** C�digo Tipo Ajuste ***
Public Const Codigo_Tipo_Ajuste_Vac = "01"
Public Const Codigo_Tipo_Ajuste_Tamex = "02"
Public Const Codigo_Tipo_Ajuste_Libor = "03"

'*** C�digo Fecha Inicial Indice ***
Public Const Codigo_Vac_Emision = "01"
Public Const Codigo_Vac_InicioPrimerCupon = "02"
Public Const Codigo_Vac_InicioCuponAnterior = "03"
Public Const Codigo_Vac_InicioCuponVigente = "04"

'*** C�digo Fecha Final Indice ***
Public Const Codigo_Vac_Liquidacion = "01"
Public Const Codigo_Vac_FinPrimerCupon = "02"
Public Const Codigo_Vac_FinCuponAnterior = "03"
Public Const Codigo_Vac_FinCuponVigente = "04"

'*** C�digo Tipo D�a ***
Public Const Codigo_Tipo_Dia_Calendario = "01"

'*** C�digo Tipo Valor de Inversi�n ***
Public Const Codigo_Valor_RentaFija = "01"
Public Const Codigo_Valor_RentaVariable = "02"

'** C�digo Tipo de Plazo de Valor Renta Fija ***
Public Const Codigo_Valor_LargoPlazo = "01"
Public Const Codigo_Valor_CortoPlazo = "02"
Public Const Codigo_Valor_MedianoPlazo = "03"

'*** C�digo Tipo Comisi�n Empresa ***
Public Const Codigo_Comision_Empresa_Safi = "01"
Public Const Codigo_Comision_Empresa_Participes = "02"

'*** C�digo Tipo Comisi�n Administraci�n ***
Public Const Codigo_Tipo_Comision_Fija = "01"

'*** C�digo Variables para Reportes ***
Public Const Codigo_Listar_Todos = "T"
Public Const Codigo_Listar_Individual = "I"

'*** C�digo Grupo de Reporte ***
Public Const Codigo_Grupo_Reporte_Limite = "L"      '*** L�mites Legales  ***
Public Const Codigo_Grupo_Reporte_Analisis = "A"    '*** An�lisis Cartera ***
Public Const Codigo_Grupo_Reporte_Reglamento = "R"  '*** Reglamento       ***
Public Const Codigo_Grupo_Reporte_Control = "D"     '*** Control Diario   ***
Public Const Codigo_Grupo_Reporte_Conasev = "C"     '*** Conasev          ***
Public Const Codigo_Grupo_Reporte_Sistema = "S"     '*** Sistema          ***
Public Const Codigo_Grupo_Reporte_Otros = "O"       '*** Control Diario   ***

'*** C�digo Tipo Orden de Inversi�n ***
Public Const Codigo_Orden_Compra = "01"
Public Const Codigo_Orden_Venta = "02"
Public Const Codigo_Orden_Pacto = "03"
Public Const Codigo_Orden_Renovacion = "07"
Public Const Codigo_Orden_Prepago = "10"
Public Const Codigo_Orden_Quiebre = "10"
Public Const Codigo_Orden_Compromiso = "15"
Public Const Codigo_Orden_PagoCancelacion = "30"

'*** C�digo de Costos de Negociaci�n ***
Public Const Codigo_Costo_Bolsa = "01"
Public Const Codigo_Costo_Conasev = "02"
Public Const Codigo_Costo_Cavali = "03"
Public Const Codigo_Costo_FGarantia = "04"
Public Const Codigo_Costo_FLiquidacion = "05"
Public Const Codigo_Costo_Agente = "08"

'*** C�digo de Tipo de Costos ***
Public Const Codigo_Tipo_Costo_Monto = "01"
Public Const Codigo_Tipo_Costo_Porcentaje = "02"

'*** C�digo de Tipo de Costos ***
Public Const Codigo_Tipo_Gasto_Periodico = "01"
Public Const Codigo_Tipo_Gasto_Unico = "02"

Public Const Codigo_Tipo_Tasa_Efectiva = "01"
Public Const Codigo_Tipo_Tasa_Nominal = "02"
Public Const Codigo_Tipo_Tasa_Flat = "03"

'*** C�digo de aplicacion de devengo ***
Public Const Codigo_Aplica_Devengo_Inmediata = "01"
Public Const Codigo_Aplica_Devengo_Periodica = "02"

'*** C�digo de modalidad de devengo ***
Public Const Codigo_Modalidad_Devengo_Provision = "01"
Public Const Codigo_Modalidad_Devengo_Ganancia_Diferida = "02"
Public Const Codigo_Modalidad_Devengo_Inmediata = "03" 'JCB R01

'*** C�digo de Tipo de devengo ***
Public Const Codigo_Tipo_Devengo_Alicuota_Lineal = "01"
Public Const Codigo_Tipo_Devengo_Alicuota_Incremental = "02"
Public Const Codigo_Tipo_Devengo_Valor_Total = "03"
Public Const Codigo_Tipo_Devengo_Valor_Total_Incremental = "04"
Public Const Codigo_Tipo_Devengo_Provision_Periodica = "05" 'NUEVO ACR

'*** C�digo de Tipo de devengo ***
Public Const Codigo_Tipo_Devengo_Provision = "01"
Public Const Codigo_Tipo_Devengo_Ganancia_Diferida = "02"

'*** C�digo Mercado de Negociacion ***
Public Const Codigo_Negociacion_Local = "01"
Public Const Codigo_Negociacion_Extranjera = "02"

'*** C�digo Tipo Evento ***
Public Const Codigo_Evento_Liberacion = "01"
Public Const Codigo_Evento_Dividendo = "02"
Public Const Codigo_Evento_Nominal = "03"
Public Const Codigo_Evento_Preferente = "04"


'*** C�digo Tipo de Cuenta de Inversi�n *** 'Estos codigos corresponden a la tabla virtual "TIPCTA"
Public Const Codigo_CtaInversion = "01"
Public Const Codigo_CtaProvInteres = "02"
Public Const Codigo_CtaInteres = "03"
Public Const Codigo_CtaCosto = "04"
Public Const Codigo_CtaIngresoOperacional = "05"
Public Const Codigo_CtaInteresVencido = "06"
Public Const Codigo_CtaVacCorrido = "07"
Public Const Codigo_CtaXPagar = "08"
Public Const Codigo_CtaXCobrar = "09"
Public Const Codigo_CtaInteresCorrido = "10"
Public Const Codigo_CtaProvReajusteK = "11"
Public Const Codigo_CtaReajusteK = "12"
Public Const Codigo_CtaProvFlucMercado = "13"
Public Const Codigo_CtaFlucMercado = "14"
Public Const Codigo_CtaProvInteresVac = "15"
Public Const Codigo_CtaInteresVac = "16"
Public Const Codigo_CtaIntCorridoK = "17"
Public Const Codigo_CtaProvFlucK = "18"
Public Const Codigo_CtaFlucK = "19"
Public Const Codigo_CtaInversionTransito = "20"
Public Const Codigo_CtaProvGasto = "21"
Public Const Codigo_CtaCostoSAB = "22"
Public Const Codigo_CtaCostoBVL = "23"
Public Const Codigo_CtaCostoCavali = "24"
Public Const Codigo_CtaCostoFondoLiquidacion = "33"
Public Const Codigo_CtaCostoGastosBancarios = "35"
Public Const Codigo_CtaCostoComisionEspecial = "34"
Public Const Codigo_CtaCostoFondoGarantia = "25"
Public Const Codigo_CtaCostoConasev = "26"
Public Const Codigo_CtaImpuesto = "27"
Public Const Codigo_CtaCompromiso = "28"
Public Const Codigo_CtaResponsabilidad = "29"
Public Const Codigo_CtaME = "30"
Public Const Codigo_CtaMN = "31"
Public Const Codigo_CtaComision = "32"
Public Const Codigo_CtaDetraccion = "36"
Public Const Codigo_CtaRetencion = "37"
Public Const Codigo_CtaIngresoOperacional_AjusteRedondeo = "38"
Public Const Codigo_Perdida_AjusteRedondeo = "39"

'Public Const Codigo_CtaIngresoRendimientoPrestamo = "38"  'ACR:14-05-2010
'Public Const Codigo_CtaGastoRendimientoPrestamo = "39"  'ACR:14-05-2010
Public Const Codigo_CtaXCobrarDividendos = "40"

Public Const Codigo_CtaInversionCostoSAB = "41"
Public Const Codigo_CtaInversionCostoBVL = "42"
Public Const Codigo_CtaInversionCostoCavali = "43"
Public Const Codigo_CtaInversionCostoFondoGarantia = "44"
Public Const Codigo_CtaInversionCostoConasev = "45"
Public Const Codigo_CtaInversionCostoIGV = "46"
Public Const Codigo_CtaInversionCostoCompromiso = "47"
Public Const Codigo_CtaInversionCostoResponsabilidad = "48"
Public Const Codigo_CtaInversionCostoFondoLiquidacion = "49"
'******
Public Const Codigo_CtaGananciaDiferidaIntereses = "50"
'*****
Public Const Codigo_CtaInversionCostoComisionEspecial = "50"
Public Const Codigo_CtaInversionCostoGastosBancarios = "51"
Public Const Codigo_CtaProvFlucMercado_Perdida = "52" '40
Public Const Codigo_CtaFlucMercado_Perdida = "53"     '41
Public Const Codigo_CtaXPagarEmitida = "54"
Public Const Codigo_CtaIngresoOperacionalPercibido = "55"
Public Const Codigo_CtaGastoIGV = "56"
Public Const Codigo_CtaInteresCastigado = "57"

'*** C�digo Tipo de Pago de Intereses ***
Public Const Codigo_Interes_Vencimiento = "001"
Public Const Codigo_Interes_Descuento = "002"

'*** C�digo Tipo Operaci�n de Caja *** 'Estos codigos corresponden a la tabla virtual "OPECAJ"
Public Const Codigo_Caja_Compra = "01"
Public Const Codigo_Caja_Venta = "02"
Public Const Codigo_Caja_Vencimiento = "03"
Public Const Codigo_Caja_Apertura = "04"
Public Const Codigo_Caja_Cierre = "05"
Public Const Codigo_Caja_Cupon = "06"
Public Const Codigo_Caja_Renovacion = "07"
Public Const Codigo_Caja_Dividendos = "08"
Public Const Codigo_Caja_Provision = "09"
Public Const Codigo_Caja_PrePago = "10"
Public Const Codigo_Caja_TransferenciaCta = "11"
Public Const Codigo_Caja_Suscripcion = "12"
Public Const Codigo_Caja_Rescate = "13"
Public Const Codigo_Caja_Transferencia = "14"
Public Const Codigo_Caja_Compromiso = "15"
Public Const Codigo_Caja_CtaME = "16"
Public Const Codigo_Caja_CtaMN = "17"
Public Const Codigo_Caja_Comision = "18"
Public Const Codigo_Caja_Gasto = "19"
Public Const Codigo_Caja_Impuesto = "20"
Public Const Codigo_Caja_Detraccion = "21"
Public Const Codigo_Caja_Retencion = "22"
Public Const Codigo_Caja_Detracci�n_AjusteRedondeoGanancia = "23"
Public Const Codigo_Caja_Detracci�n_AjusteRedondeoPerdida = "24"
Public Const Codigo_Caja_Gasto_Emitidas = "25"
Public Const Codigo_Caja_Facturacion = "26"
Public Const Codigo_Caja_Dividendos_Percibidos = "27"
Public Const Codigo_Caja_Abono_Participe = "28"
Public Const Codigo_Caja_Extorno_Provision = "29"
Public Const Codigo_Caja_Cancelacion = "30"
Public Const Codigo_Caja_Cancelacion_Susc_Cuotas = "31"
Public Const Codigo_Caja_Liq_Operacion_Inversion = "32"
Public Const Codigo_Caja_Liq_Quiebre_Inversion = "33"
Public Const Codigo_Caja_Liq_Cancelacion_Inversion = "34"
Public Const Codigo_Caja_Provision_Gasto = "35"
Public Const Codigo_Caja_Provision_ComisionP = "36"
Public Const Codigo_Caja_Provision_Intereses_Adicionales = "37"
Public Const Codigo_Caja_Provision_Intereses_Moratorios = "39"
Public Const Codigo_Caja_Operacion_Cambio = "49"

'*** C�digo Tipo Forma de Pago *** MEDPAG
Public Const Codigo_FormaPago_Cheque = "01"
Public Const Codigo_FormaPago_Cuenta = "02"
Public Const Codigo_FormaPago_Efectivo = "03"
Public Const Codigo_FormaPago_Transferencia_Mismo_Banco = "04"
Public Const Codigo_FormaPago_Transferencia_Otro_Banco = "05"
Public Const Codigo_FormaPago_Transferencia_Exterior = "06"

'*** C�digo Forma Ingreso Calidad de Part�cipe ***
Public Const Codigo_FormaIngreso_Suscripcion = "01"

'*** C�digo Tipo Bloqueo ***
Public Const Codigo_Tipo_Bloqueo_Emision = "04"

'*** C�digo Tipo Direcci�n Postal ***
Public Const Codigo_Direcci�n_Domicilio = "01"
Public Const Codigo_Direcci�n_Trabajo = "02"

'*** C�digo Tipo de Limites ***
Public Const Codigo_Limite_Patrimonio = "01"
Public Const Codigo_Limite_CreditoVigente = "02"
Public Const Codigo_Limite_Riesgo = "03"
Public Const Codigo_Limite_Instrumento = "04"
Public Const Codigo_Limite_Activo = "05"

'*** C�digos de LimiteReglamentoEstructura  ***
Public Const Codigo_LimiteRE_Cliente = "09"

'*** C�digos de L�neas de Cliente ***
Public Const Linea_Descuento_Letras_Facturas = "42"
Public Const Linea_Financiamiento_Proveedores = "43"
Public Const Linea_Contrato_Flujo_Dinerario = "44"
Public Const Linea_Compra_Maquinarias = "45"

'*** C�digo Tipo de Dato Parametro ***
Public Const Codigo_TipoDato_Numerico = "01"
Public Const Codigo_TipoDato_AlfaNumerico = "02"
Public Const Codigo_TipoDato_Fecha = "03"

'*** C�digo Mecanismo Negociaci�n ***
Public Const Codigo_Mecanismo_Rueda = "01"

'*** C�digo Plazo Operaci�n ***
Public Const Codigo_Operacion_Contado = "01"
Public Const Codigo_Operacion_Plazo = "02"

'*** C�digo Afectaci�n Impuesto ***
Public Const Codigo_Afecto = "01"
Public Const Codigo_Inafecto = "02"

'*** C�digo Tipo de Pago Gasto ***
Public Const Codigo_Pago_Adelantado = "02"
Public Const Codigo_Pago_Vencimiento = "01"

'*** C�digo Tipo de Gasto ***
Public Const Codigo_Gasto_Provision = "01"
Public Const Codigo_Gasto_MismoDia = "02"
Public Const Codigo_Gasto_Devengado = "03" 'ACR

'*** Tipo de Movimiento Bancario
Public Const Tipo_Abono = "02"
Public Const Tipo_Retiro = "01"

'*** C�digo Aplicacion de Devengo ***
Public Const Codigo_Aplicacion_Devengo_Periodica = "01"
Public Const Codigo_Aplicacion_Devengo_Inmmediata = "02"

'*** C�digo Signos Aplicaci�n Costos Negociaci�n ***
Public Const Codigo_Signo_Igual = "01"
Public Const Codigo_Signo_Menor = "02"
Public Const Codigo_Signo_Mayor = "03"
Public Const Codigo_Signo_MenorIgual = "04"
Public Const Codigo_Signo_MayorIgual = "05"

'*** C�digo Tipo de Comprobantes de Pago *** 'Completar segun se requiera!
Public Const Codigo_Comprobante_Factura = "01"
Public Const Codigo_Comprobante_Recibo_Honorarios = "02"
Public Const Codigo_Comprobante_Boleta_Venta = "03"
Public Const Codigo_Comprobante_Documento_Emitido_Bancos = "13"

'*** JAFR 14/11/11 Constantes de Tipo de asiento contable ***
Public Const Tipo_Asiento_Apertura_Cierre = "00"
Public Const Tipo_Asiento_Gasto = "01"
Public Const Tipo_Asiento_Inversion = "02"
Public Const Tipo_Asiento_Cobranza = "03"
Public Const Tipo_Asiento_Caja_Chica = "04"
Public Const Tipo_Asiento_Diario = "05"
Public Const Tipo_Asiento_Planilla = "06"
Public Const Tipo_Asiento_Regularizacion_Intereses = "07"
Public Const Tipo_Asiento_Suscripcion_Rescate = "08"
Public Const Tipo_Asiento_Reparto_Utilidades = "09"
Public Const Tipo_Asiento_BancosME = "10"
Public Const Tipo_Asiento_BancosMN = "11"
Public Const Tipo_Asiento_Distribucion_Utilidades_Reparto_Reinversion = "12"
Public Const Tipo_Asiento_Distribucion_Utilidades_Reversion_Resultados = "13"
Public Const Tipo_Asiento_Provision_Comisiones_Promotores = "14"
Public Const Tipo_Asiento_Provision_Gastos_Proveedores = "15"
Public Const Tipo_Asiento_Provision_Comisiones_SAF = "16"
Public Const Tipo_Asiento_Valorizacion_Inversiones = "17"
Public Const Tipo_Asiento_Ajuste_Diferencia_Cambio = "18"
Public Const Tipo_Asiento_Ajuste_Traslacion = "19"
Public Const Tipo_Asiento_Resultados_Ejercicio = "20"
Public Const Tipo_Asiento_Automatico = "99"

'*** C�digo Valor Tipo de Cambio ***
Public Const Codigo_Valor_TipoCambioCompra = "01"
Public Const Codigo_Valor_TipoCambioVenta = "02"

'*** C�digo Clasificaci�n Tipo de Cambio ***
Public Const Codigo_TipoCambio_SBS = "01"
Public Const Codigo_TipoCambio_Bancario = "02"
Public Const Codigo_TipoCambio_Reuters = "03"
Public Const Codigo_TipoCambio_Conasev = "04"
Public Const Codigo_TipoCambio_Sunat = "05"

'*** C�digo Tipo Cr�dito Fiscal ***
Public Const Codigo_Tipo_Credito_RentaGravada = "01"
Public Const Codigo_Tipo_Credito_RentaNoGravada = "02"
Public Const Codigo_Tipo_Credito_RentaGravadaNoGravada = "03"
Public Const Codigo_Tipo_Credito_AdquisicionesNoGravada = "04"

'*** C�digo Tipo Indicador Cobertura ***
Public Const Codigo_IndCobertura_Amortizacion = "04"
Public Const Codigo_IndCobertura_Cupon = "02"
Public Const Codigo_IndCobertura_CuponPrincipal = "03"
Public Const Codigo_IndCobertura_Principal = "01"

'*** C�digo Tipo Cobertura ***
Public Const Codigo_Tipo_Cobertura_Sintetico = "01"
Public Const Codigo_Tipo_Cobertura_Independiente = "02"

'*** C�digo de File de Gastos ***
Public Const Codigo_File_Comision_Inversion = "097"
Public Const Codigo_File_Comision = "098"
Public Const Codigo_File_Gasto = "099"

'*** Tipo de Calculo ***
Public Const Codigo_Tipo_Calculo_Fijo = "01"
Public Const Codigo_Tipo_Calculo_Variable = "02"

'*** Cuentas de Comisi�n ***
Public Const Codigo_Cuenta_Comision_Fija = "632911"
Public Const Codigo_Cuenta_Comision_Variable = "632912"

'*** Tipo de Objeto Sistema ***
Public Const Codigo_Tipo_Objeto_Modulo = "01"
Public Const Codigo_Tipo_Objeto_Menu = "02"
Public Const Codigo_Tipo_Objeto_Formulario = "03"
Public Const Codigo_Tipo_Objeto_Control = "04"

'*** Atributos Objeto Sistema ***
Public Const Codigo_Atributo_NoDefinido = "001"
Public Const Codigo_Atributo_Enabled = "002"
Public Const Codigo_Atributo_Visible = "003"
Public Const Codigo_Atributo_Acceso = "004"

'*** Separador de nombres en el Codigo de Objeto ***
Public Const Separador_Codigo_Objeto = "."

'** Variable **
Public gstrNombreObjetoMenuPulsado          As String

'*** Nombre del Formulario **
Global frmFormulario As Form

'*** Dinamicas 24/05/13***
Public Const Codigo_Dinamica_Liquidacion_Ingreso_Comitente = "43"
Public Const Codigo_Dinamica_Liquidacion_Ingreso_Cliente = "44"

'*** Estados de Registro 27/05/13****
Public Const Codigo_Estado_Registro_Activo = "01"
Public Const Codigo_Estado_Registro_Inactivo = "02"
Public Const Codigo_Estado_Registro_Eliminado = "03"


