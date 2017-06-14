Attribute VB_Name = "NegocioInversiones"
Option Explicit

Public strNumAsiento                As String
Public strNumOperacion              As String
Public strNumKardex                 As String
Public strNumCaja                   As String
Public curIntCapRea                 As Currency
Public intNemotecnicoInd            As Integer
Public strNemotecnicoVal            As String

Public Function calcularValorDias(ByVal valorBase As Double, ByVal valorDias As Double) As Double
'comentario!!
    Dim result As Double
    result = (valorDias * valorBase) / 360
    calcularValorDias = result
End Function

Public Function CalculaVAN(strpCodTitulo As String, strpFecha As String, TIRDiario As Double) As Double

    Dim adoResultAux As New ADODB.Recordset, adoRegistro As ADODB.Recordset
    Dim FechaRede As String, ValNomi As Double, nDiasAcum As Integer
    Dim IntCorrido As Double, AcumFlujo As Double
    Dim TasDiar As Double, FchIntCor As String, DiasIntCorr As Integer
    Dim FchCupoMenos1 As Variant, res As Integer
    Dim n_TasDia As Double, n_Tasa As Double
   
    ' FUNCION QUE CALCULA EL PRECIO A
    ' PARTIR DE LA TIR BRUTA DE LA OPERACION.
    
    Set adoRegistro = New ADODB.Recordset
    With adoComm
        'Obtener datos del Título
        .CommandText = "SELECT * FROM InstrumentoInversion " & _
            "WHERE CodTitulo='" & strpCodTitulo & "'"
        Set adoRegistro = .Execute

        'Para Intereses Corridos obtener Cupon Actual donde este incluido strpFecha, capturar FCH_INIC, TAS_DIAR
        'Si el Cupón Vigente es el Primero tomar FCH_INIC, caso contrario tomar FCH_INIC Menos 1 DIA
        .CommandText = "SELECT NumCupon, FechaInicio, FactorDiario FROM InstrumentoInversionCalendario " & _
            "WHERE CodTitulo='" & strpCodTitulo & "' AND " & _
            "FechaInicio<='" & strpFecha & "' AND " & _
            "FechaVencimiento>='" & strpFecha & "'"
        Set adoResultAux = .Execute
        
        If Not adoResultAux.EOF Then  ' Si tiene cupon vigente
            TasDiar = adoResultAux("FactorDiario")
            If adoResultAux("NumCupon") = "001" Then
                FchIntCor = adoResultAux("FechaInicio")
            Else
                FchIntCor = Convertyyyymmdd(DateAdd("d", -1, adoResultAux("FechaInicio")))
            End If
        Else
            TasDiar = 0
            FchIntCor = strpFecha   'Como no existe asumo la strpFecha de cierre para que de 0 dias
        End If
        adoResultAux.Close ': Set adoResultAux = Nothing
   
        nDiasAcum = 0: AcumFlujo = 0
        n_TasDia = 0: n_Tasa = 0

        .CommandText = "SELECT NumCupon, FactorInteres,ValorCupon, IndVencido, FechaInicio, FechaVencimiento,CantDiasPeriodo,ValorAmortizacion " & _
            "FROM InstrumentoInversionCalendario " & _
            "WHERE CodTitulo='" & strpCodTitulo & "' AND " & _
            "IndVencido <> 'X'"
        Set adoResultAux = .Execute
        
        Do Until adoResultAux.EOF   ' Si el bono tiene cupones vigentes
            nDiasAcum = DateDiff("d", Convertddmmyyyy(strpFecha), adoResultAux("FechaVencimiento"))
            If adoResultAux("FactorInteres") = 0 Then
            'If adoRegistro!FLG_AMORT = "X" Then
            '    n_Tasa = ((n_TasDia# ^ adoresultAux!CNT_DIAS) - 1) + adoresultAux!VAL_AMOR
            '    AcumFlujo = AcumFlujo + format$((n_Tasa / ((1 + TIRDiario) ^ nDiasAcum)), "0.0000000000") + adoresultAux!VAL_AMOR
            'Else
                n_Tasa = ((n_TasDia ^ adoResultAux("CantDiasPeriodo")) - 1)
                AcumFlujo = AcumFlujo + (n_Tasa / ((1 + TIRDiario) ^ nDiasAcum))
            'End If
            Else
            'If adoRegistro!FLG_AMORT = "X" Then
            '    AcumFlujo = AcumFlujo + format$(((adoresultAux!TAS_INTE + adoresultAux!VAL_AMOR) / ((1 + TIRDiario) ^ nDiasAcum)), "0.0000000000") + adoresultAux!VAL_AMOR
            'Else
                AcumFlujo = AcumFlujo + (adoResultAux("FactorInteres") / ((1 + TIRDiario * 0.01) ^ nDiasAcum))
            'End If
                n_TasDia = ((1 + adoResultAux("FactorInteres")) ^ (1 / adoResultAux("CantDiasPeriodo")))
            End If
            adoResultAux.MoveNext
        Loop
   
        AcumFlujo = AcumFlujo + (adoRegistro("ValorNominal") / ((1 + TIRDiario * 0.01) ^ nDiasAcum)) ' Incluir Valor Nominal

        If FchIntCor <> strpFecha Then 'Existen Intereses Corridos
            DiasIntCorr = DateDiff("d", Convertddmmyyyy(FchIntCor), Convertddmmyyyy(strpFecha))
            IntCorrido = (((1 + TasDiar) ^ DiasIntCorr) - 1) * adoRegistro("ValorNominal")
        Else
            IntCorrido = Format$(0, "0.0000000000")
        End If

        AcumFlujo = ((AcumFlujo - IntCorrido) / adoRegistro("ValorNominal"))
   
        adoResultAux.Close: Set adoResultAux = Nothing
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    CalculaVAN = AcumFlujo

End Function

'JAFR: Calculo iterativo de cuota constante
Public Function CalculoCuotaConstante(ByVal CodTitulo As String, ByVal SubDetalleFile As String, ByVal cti_igv As Boolean, ByVal Principal As Double, _
                                        ByVal cantCupones As Double, ByVal fechaEmision As Date, ByVal fechavencimiento As Date, ByVal tolerancia As Double) As Double

    'Variables generales para la definicion del calendario
    Dim numCuota                 As Integer
    Dim fechaInicio              As Date
    Dim fechaCorte               As Date
    Dim fechaCorteAnterior       As Date
    Dim fechaCorteDesplazada     As Date
    Dim fechaPago                As Date
    Dim interes                  As Double
    Dim total                    As Double
    Dim amortizacion             As Double
    Dim acumuladoAmortizacion    As Double
    Dim cuota                    As Double
    Dim saldoDeudorInicial       As Double
    Dim saldoDeudorFinal         As Double
    Dim blnFaltanCupones         As Boolean
    Dim Tasa                     As Double
    Dim TasaRecalculadaIgv       As Double
    Dim igvIntereses             As Double
    Dim saldoIgvIntereses        As Double
    
    'condiciones financieras tomadas de la BD
    Dim fechaPrimerCorte         As Date
    Dim tasaIni                  As Double
    Dim strTipoTasa              As String
    Dim unidadesPeriodo          As Integer
    Dim indCorteAFinPeriodo      As Boolean
    Dim indPeriodoPersonalizable As Integer
    Dim indCortePrimerCupon      As Integer
    Dim indexDesplazamientoCorte As Integer
    
    Dim strCodDesplazamientoCorte   As String
    Dim strCodDesplazamientoPago    As String
    
    Dim indexDesplazamientoPago  As Integer
    Dim indexPeriodoTasa         As Integer
    Dim indexTipoCupon           As Integer
    Dim indexBaseCalculo         As Integer
    Dim indexPeriodoCupon        As Integer
    Dim indexUnidadPeriodo       As Integer
    'Control de la cuota y ajuste
    Dim anteriorCuota            As Double
    Dim proximaCuota             As Double
    Dim ajusteCuota              As Double
    Dim iteraciones              As Integer
 

    Dim comm                     As ADODB.Command
    Set comm = New ADODB.Command
    
    Dim i As Integer
    
    ReDim listaCupones(cantCupones)
    
    Dim adoRegistroCondicionesFinancieras As New ADODB.Recordset
        
    'query de la tabla instrumentoInversionCondicionesFinancieras
    adoComm.CommandText = "SELECT IndNumCuotas, NumCuotas, FechaEmision, FechaVencimiento, ValorNominal,Tasa, TipoTasa, PeriodoTasa, TipoCupon, BaseCalculo, TipoAmortizacion, PeriodoCupon, " & _
                          "IndPeriodoPersonalizable, CantUnidadesPeriodo, UnidadPeriodo, DesplazamientoCorte, DesplazamientoPago, IndCortePrimerCupon, FechaPrimerCorte " & _
                          "from InstrumentoInversionCondicionesFinancieras where CodTitulo = '" & CodTitulo & "'"
    Set adoRegistroCondicionesFinancieras = adoComm.Execute
    
    'Procesamiento de adoRegistroCondicionesFinancieras
    
    tasaIni = adoRegistroCondicionesFinancieras.Fields.Item("Tasa")
    
    If adoRegistroCondicionesFinancieras.Fields.Item("TipoTasa") = "01" Then
        strTipoTasa = "Efectiva"
    ElseIf adoRegistroCondicionesFinancieras.Fields.Item("TipoTasa") = "02" Then
        strTipoTasa = "Nominal"
    End If
    
    indexPeriodoTasa = CInt(adoRegistroCondicionesFinancieras.Fields.Item("PeriodoTasa")) - 1
    indexTipoCupon = CInt(adoRegistroCondicionesFinancieras.Fields.Item("TipoCupon")) - 1
    indexBaseCalculo = CInt(adoRegistroCondicionesFinancieras.Fields.Item("BaseCalculo"))
    
    'ajuste de indexbasecalculo
    If indexBaseCalculo = 1 Then
        indexBaseCalculo = 4
    ElseIf indexBaseCalculo > 3 Then
        indexBaseCalculo = indexBaseCalculo - 4
    End If
    
    indexPeriodoCupon = CInt(adoRegistroCondicionesFinancieras.Fields.Item("PeriodoCupon")) - 1
    indPeriodoPersonalizable = CInt(adoRegistroCondicionesFinancieras.Fields.Item("IndPeriodoPersonalizable"))
    unidadesPeriodo = adoRegistroCondicionesFinancieras.Fields.Item("CantUnidadesPeriodo")
    indexUnidadPeriodo = adoRegistroCondicionesFinancieras.Fields.Item("UnidadPeriodo") - 1
    indexDesplazamientoCorte = CInt(adoRegistroCondicionesFinancieras.Fields.Item("DesplazamientoCorte"))
    strCodDesplazamientoCorte = adoRegistroCondicionesFinancieras.Fields.Item("DesplazamientoCorte")
    indexDesplazamientoPago = CInt(adoRegistroCondicionesFinancieras.Fields.Item("DesplazamientoPago"))
    strCodDesplazamientoPago = adoRegistroCondicionesFinancieras.Fields.Item("DesplazamientoPago")
    indCortePrimerCupon = adoRegistroCondicionesFinancieras.Fields.Item("IndCortePrimerCupon")
    
    fechaPrimerCorte = adoRegistroCondicionesFinancieras.Fields.Item("FechaPrimerCorte")
    saldoDeudorInicial = Principal
    
    proximaCuota = Round(saldoDeudorInicial / cantCupones, 2)
    anteriorCuota = 0
    ajusteCuota = tolerancia + 1
    iteraciones = 0
    
    While Abs(ajusteCuota) > tolerancia And iteraciones < 25
        iteraciones = iteraciones + 1
        'inicializacion de variables
        fechaInicio = fechaEmision
        fechaCorte = fechaInicio
        fechaCorteDesplazada = fechaInicio
        
        cuota = proximaCuota
        numCuota = 0
        interes = 0
        total = 0
        amortizacion = 0
        acumuladoAmortizacion = 0
        saldoDeudorInicial = 0
        saldoDeudorFinal = 0
        igvIntereses = 0
        blnFaltanCupones = True
        
 
    
        Tasa = tasaIni / 100
    '    Tasa = recalcularTasa(tasaIni, strTipoTasa, indexPeriodoTasa, indexPeriodoCupon, indexBaseCalculo, unidadesPeriodo, fechaEmision)
    '    Tasa = recalcularTasaDifDias(tasaIni / 100, indexPeriodoTasa, indexBaseCalculo, fechaEmision, indexPeriodoCupon, strTipoTasa, DateDiff("d", fechaEmision, fechavencimiento))
        saldoDeudorFinal = Principal

        Tasa = recalcularTasa(tasaIni, strTipoTasa, indexPeriodoTasa, indexPeriodoCupon, indexBaseCalculo, unidadesPeriodo, fechaEmision)

        While blnFaltanCupones
            saldoDeudorInicial = saldoDeudorFinal
            numCuota = numCuota + 1
            fechaCorteAnterior = fechaCorteDesplazada
            
            'calculo de la fecha de corte
            If (numCuota = 1) And (indCortePrimerCupon = 1) And (fechaEmision <= fechaPrimerCorte) Then
                fechaCorte = fechaPrimerCorte
                fechaCorteAnterior = fechaInicio
                fechaPago = fechaCorte
            Else

                If indCorteAFinPeriodo = True Then
                    Dim tempfechainicio As Date
                    tempfechainicio = DateAdd("d", 1, fechaInicio)
                    tempfechainicio = DateAdd("d", 1, tempfechainicio)

                    If numCuota > 1 Then
                        fechaCorte = ultimaFechaPeriodo(tempfechainicio, indexPeriodoCupon)
                    Else
                        fechaCorte = ultimaFechaPeriodo(fechaInicio, indexPeriodoCupon)
                    End If

                Else
                    fechaCorteAnterior = fechaInicio
                    fechaCorte = CalculaFechaSiguienteCalendario(fechaInicio, indexBaseCalculo, indexPeriodoCupon, indexUnidadPeriodo, unidadesPeriodo)
                End If

                If (fechaCorte >= fechavencimiento) Or (numCuota = cantCupones) Then
                    If (fechaCorte >= fechavencimiento) Then
                        fechaCorte = fechavencimiento
                    End If
                End If
            End If

            'fin del cálculo de la fecha de corte
            
            'si no es el primer cupon con fecha de corte especifica....
            'se realiza el desplazamiento especificado tanto para fecha de corte como para fecha de pago
            If Not ((numCuota = 1) And (indCortePrimerCupon = 1)) Then
                fechaCorteDesplazada = desplazamientoDiaLaborable(fechaCorte, strCodDesplazamientoCorte)
                fechaPago = desplazamientoDiaLaborable(fechaCorte, strCodDesplazamientoPago)
                fechaCorte = fechaCorteDesplazada
             End If
            
            'Teniendo la fecha de corte desplazada, se calcula tasa entre las dos fechas de corte pertinentes
            Tasa = recalcularTasaDifDias(tasaIni / 100, indexPeriodoTasa, indexBaseCalculo, fechaEmision, indexPeriodoCupon, strTipoTasa, DateDiff("d", fechaCorteAnterior, fechaCorte))
            
            If fechaCorte >= fechavencimiento Then blnFaltanCupones = False
            'hasta aqui ya se tienen calculadas las fechas.
            
            'aqui inicia el cálculo del monto a pagar en el cupon


    '-------------------------------------'
            interes = Round(saldoDeudorInicial * Tasa, 2)
            
            If SubDetalleFile = "001" Then

                If cti_igv Then
                    igvIntereses = Round(interes * gdblTasaIgv, 2)
                Else
                    igvIntereses = 0
                End If

            Else
                igvIntereses = 0
            End If
            
            amortizacion = cuota - (interes + igvIntereses)

    '--------------------------------------------------------'

            saldoDeudorFinal = saldoDeudorInicial - amortizacion
                   
            acumuladoAmortizacion = acumuladoAmortizacion + amortizacion
                    
            Dim cuotaTmp As Double
            Dim intRegistro As Integer


            fechaInicio = fechaCorte

        Wend
        anteriorCuota = cuota
       
        ajusteCuota = Round(saldoDeudorFinal / cantCupones, 2)
        proximaCuota = ajusteCuota + anteriorCuota
    Wend
    
    CalculoCuotaConstante = proximaCuota
End Function

'Caso para el cual se deba calcular el igv primero.
Public Function calculoCuotaConstanteIgv(ByVal Principal As Double, ByVal TasaRecalculadaIgv As Double, ByVal periodos As Double) As Double
    Dim cuota As Double
    cuota = Principal * (potencia(1 + TasaRecalculadaIgv, periodos) * TasaRecalculadaIgv) / (potencia(1 + TasaRecalculadaIgv, periodos) - 1)
    calculoCuotaConstanteIgv = cuota
End Function 'Fin JJCC
Public Sub ConfirmarOrden(ByVal strFechaRegistro As String, ByVal strFechaMas1Dia As String, ByVal strNumOrden As String, ByVal strCodFondo As String, ByVal strObservacion As String, ByVal strpNumCobertura As String)
    
    '*** Liquidación de Ordenes de Compra / Venta ***
    Dim adoRegistro                 As ADODB.Recordset, adoTemporal         As ADODB.Recordset
    Dim curKarSldInic               As Currency, curKarSldFina              As Currency
    Dim curKarValSald               As Currency, dblKarValProm              As Double
    Dim curValComi                  As Currency, dblKarSldAmort             As Double
    Dim curKarIadSald               As Currency, dblKarIadProm              As Double
    Dim curVAN                      As Currency, curValorAmortizacion       As Currency
    Dim curVANLimpio                As Currency
    Dim curValorNominal             As Currency, curCantOrden               As Currency
    Dim curCantMovimiento           As Currency, curValorMovimiento         As Currency
    Dim curSaldoInicialKardex       As Currency, curSaldoFinalKardex        As Currency
    Dim curValorSaldoKardex         As Currency, curSaldoInteresCorrido     As Currency
    Dim curSaldoAmortizacion        As Currency, curVacCorrido              As Currency
    Dim curMontoMN                  As Currency, curMontoME                 As Currency
    Dim curMontoContable            As Currency, curValorCupon              As Currency
    Dim dblTirProm                  As Double, dblTirPromAnt                As Double
    Dim dblTirPromLimpia            As Double
    Dim dblPrecioUnitario           As Double, dblTirNeta                   As Double
    Dim dblTirNetaKardex            As Double, dblTirOperacionKardex        As Double
    Dim dblTirPromedioKardex        As Double, dblValorPromedioKardex       As Double
    Dim dblInteresCorridoPromedio   As Double, dblTipoCambio                As Double
    Dim dblTipoCambioMonedaPago     As Double, dblTipoCambioArbitraje       As Double
    Dim dblMontoVencimiento         As Double, dblTirBruta                  As Double
    Dim dblTasaInteres              As Double, dblFactor1                   As Double
    Dim dblFactor2                  As Double, dblFactorAnualCupon          As Double
    Dim dblTasaCuponNormal          As Double, dblFactorDiario              As Double
    Dim dblSaldoAmortizacion        As Double, dblTasaAnual                 As Double
    Dim dblValorInteres             As Double, dblAcumuladoAmortizacion     As Double
    Dim dblFactorDiarioNormal       As Double, dblValorAmortizacion         As Double
    Dim lngNumAnalitica             As Long, lngNumOperacion                As Long
    Dim lngNumAsiento               As Long, lngNumKardex                   As Long
    Dim lngNumOrdenCaja             As Long, lngDiasTitulo                  As Long
    Dim intContador                 As Integer
    Dim intCantMovAsiento           As Integer, intDiasPlazo                As Integer
    Dim intRegistro                 As Integer, intDiasPeriodo              As Integer
    Dim strIndAmortizacion          As String, strFlgTvac                   As String
    Dim strCodMoneda                As String, strIndTitulo                 As String
    Dim strCodMonedaPago            As String
    Dim strCodAnalitica             As String, strCodTitulo                 As String
    Dim strCodClaseInstrumento      As String, strCodSubClaseInstrumento    As String
    Dim strCodFile                  As String, strIndInversion              As String
    Dim strDescripTitulo            As String, strFechaOrden                As String
    Dim strFechaLiquidacion         As String, strFechaEmision              As String
    Dim strFechaVencimiento         As String, strCodEmisor                 As String
    Dim strCodCiiu                  As String, strCodGrupo                  As String
    Dim strCodSector                As String, strCodTipoTasa               As String
    Dim strCodBaseAnual             As String, strIndGenerado               As String
    Dim strIndCuponCero             As String, strTipoMovimientoKardex      As String
    Dim strSQLTitulo                As String, strSQLOperacion              As String
    Dim strSQLKardex                As String, strSQLOrdenCaja              As String
    Dim strSQLOrdenCajaDetalle      As String, strSQLAsientoContable        As String
    Dim strSQLCupon                 As String, strCodGarantia               As String
    Dim strCodAgente                As String, strDescripOrden              As String
    Dim strCodOperacion             As String, strCodReportado              As String
    Dim strCodGirador               As String, strCodAceptante              As String
    Dim strIndUltimoMovimiento      As String, strIndDebeHaber              As String
    Dim strDescripMovimiento        As String, strCodTipoOrden              As String
    Dim strCodOperacionCaja         As String, strCodTipoOperacion          As String
    Dim strIndPorcenPrecio          As String, strCodNemonico               As String
    Dim strCodSubDetalleFile        As String, strCodNegociacion            As String
    Dim strCodOrigen                As String, strIndCustodia               As String
    Dim strIndKardex                As String, strTipoTasa                  As String
    Dim strBaseAnual                As String, strRiesgo                    As String
    Dim strSubRiesgo                As String, strIndGasto                  As String
    Dim strIndGastoImpuesto         As String
    
    Dim strIndContableComision    As String
    Dim strIndContableImpuesto    As String
    
    Dim strFechaGrabar              As String, strFechaPago                 As String
    Dim datFchCalc                  As Date
    Dim blnEstado                   As Boolean
    Dim strCodMercado               As String
    
    'ACR: 11-06-2010
    Dim curSaldoInversion                       As Currency
    Dim curSaldoInversionCostoSAB               As Currency
    Dim curSaldoInversionCostoBVL               As Currency
    Dim curSaldoInversionCostoCavali            As Currency
    Dim curSaldoInversionCostoFondoGarantia     As Currency
    Dim curSaldoInversionCostoConasev           As Currency
    Dim curSaldoInversionCostoIGV               As Currency
    Dim curSaldoInversionCostoCompromiso        As Currency
    Dim curSaldoInversionCostoResponsabilidad   As Currency
    Dim curSaldoInversionCostoFondoLiquidacion  As Currency
    Dim curSaldoInversionCostoComisionEspecial  As Currency
    Dim curSaldoInversionCostoGastosBancarios   As Currency
    Dim curSaldoProvInteres                     As Currency
    Dim curSaldoInteres                         As Currency
    Dim curSaldoInteresVencido                  As Currency
    Dim curSaldoInteresCorridoAcum              As Currency
    Dim curSaldoProvFlucMercado                 As Currency
    Dim curSaldoFlucMercado                     As Currency
    Dim curSaldoProvFlucMercadoPerdida          As Currency
    Dim curSaldoFlucMercadoPerdida              As Currency
    Dim curSaldoProvFlucK                       As Currency
    Dim curSaldoFlucK                           As Currency
    'ACR: 11-06-2010
    
    Dim dblTirOperacion                         As Double
    Dim dblTirOperacionLimpia                   As Double
    Dim dblTirPromLimpiaAnt                     As Double

    Dim dblTirOperacionKardexLimpia             As Double
    Dim dblTirBrutaLimpia                       As Double
    Dim dblTirPromedioKardexLimpia              As Double
    Dim curVANLimpia                            As Double
    Dim dblPrecioPromedio                       As Double
    Dim dblPrecioPromedioLimpio                 As Double
    Dim dblPrecioUnitarioSucio                  As Double
    
    
    Dim dblKarPrecioPromedio                    As Double
    Dim dblKarPrecioPromedioLimpio              As Double
    
    
    'ACR: 24-08-2010
    Dim strTipoDocumento                        As String
    Dim strNumDocumento                         As String
    'ACR: 24-08-2010
        
    Dim strTipoPersona                          As String
    Dim strCodPersona                           As String
    Dim strNumRucFondo                          As String
        
    'On Error GoTo Ctrl_Error
        
    With adoComm
        Set adoRegistro = New ADODB.Recordset
        
        curValComi = 0
        
        '*** Obtener Secuenciales ***
        strNumAsiento = ObtenerSecuencialInversionOperacion(strCodFondo, Valor_NumComprobante)
        strNumOperacion = ObtenerSecuencialInversionOperacion(strCodFondo, Valor_NumOperacion)
        strNumKardex = ObtenerSecuencialInversionOperacion(strCodFondo, Valor_NumKardex)
        strNumCaja = ObtenerSecuencialInversionOperacion(strCodFondo, Valor_NumOrdenCaja)

        '*** Consultar Orden ***
        .CommandText = "SELECT * FROM InversionOrden WHERE (FechaOrden >='" & strFechaRegistro & "' AND FechaOrden <'" & strFechaMas1Dia & "') AND " & _
            "NumOrden='" & strNumOrden & "' AND CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
        Set adoRegistro = .Execute

        If Not adoRegistro.EOF Then
            If Trim(adoRegistro("TipoOrden")) = Codigo_Orden_Pacto Then
                strCodTipoOrden = Codigo_Orden_Compra
            Else
                strCodTipoOrden = Trim(adoRegistro("TipoOrden"))
            End If
            
            Set adoTemporal = New ADODB.Recordset
            
            '*** Numero de movimientos de dinamica contable ***
            adoComm.CommandText = "SELECT COUNT(*) AS NumRegistros FROM DinamicaContable " & _
                "WHERE TipoOperacion='" & strCodTipoOrden & "' AND CodFile='" & Trim(adoRegistro("CodFile")) & "' AND " & _
                "(CodDetalleFile = '" & Trim(adoRegistro("CodDetalleFile")) & "' OR CodDetalleFile='000') AND " & _
                "(CodSubDetalleFile = '" & Trim(adoRegistro("CodSubDetalleFile")) & "' OR CodSubDetalleFile = '000') AND " & _
                "CodAdministradora='" & gstrCodAdministradora & "' AND CodMoneda = '" & _
                IIf(Trim(adoRegistro("CodMoneda")) = Codigo_Moneda_Local, Codigo_Moneda_Local, Codigo_Moneda_Extranjero) & "'"
            Set adoTemporal = adoComm.Execute
            
            If Not adoTemporal.EOF Then
                If adoTemporal("NumRegistros") > 0 Then
                    intCantMovAsiento = CInt(adoTemporal("NumRegistros"))
                Else
                    adoTemporal.Close: Set adoTemporal = Nothing: Exit Sub   'GoTo Ctrl_Error:
                End If
            Else
                adoTemporal.Close: Set adoTemporal = Nothing: Exit Sub  'GoTo Ctrl_Error:
            End If
            
            adoTemporal.Close
            
            '*** El Precio es % ? ***
            strIndPorcenPrecio = Valor_Caracter
            
            .CommandText = "SELECT IndPorcenPrecio,IndGasto FROM InversionFile " & _
                "WHERE CodFile='" & Trim(adoRegistro("CodFile")) & "'"
            Set adoTemporal = .Execute
            
            If Not adoTemporal.EOF Then
                strIndPorcenPrecio = Trim(adoTemporal("IndPorcenPrecio"))
            End If
            adoTemporal.Close
           
            'OBTIENE NUMERO DE RUC DEL FONDO
            .CommandText = "{ call up_ACSelDatosParametro(24,'" & strCodFondo & "','" & gstrCodAdministradora & "') }"
            
            Set adoTemporal = .Execute
            
            If Not adoTemporal.EOF Then
                strNumRucFondo = CStr(adoTemporal("NumRucFondo"))
                strTipoPersona = Codigo_Tipo_Persona_Portafolio
            End If
            
            adoTemporal.Close
            
            'OBTIENE CODIGO DE PERSONA DEL FONDO
            .CommandText = "SELECT CodPersona FROM InstitucionPersona"
            .CommandText = .CommandText & " WHERE "
            .CommandText = .CommandText & " TipoPersona   = '" & strTipoPersona & "' AND "
            .CommandText = .CommandText & " TipoIdentidad = '" & Codigo_Tipo_Registro_Unico_Contribuyente & "' AND "
            .CommandText = .CommandText & " NumIdentidad  = '" & strNumRucFondo & "'"
        
            Set adoTemporal = .Execute
            
            If Not adoTemporal.EOF Then
                strCodPersona = adoTemporal("CodPersona")
            Else
                strCodPersona = "00000000"
            End If
            
            adoTemporal.Close
            
            strCodOrigen = Trim(adoRegistro("CodOrigen"))
           
            If strCodOrigen = Codigo_Mercado_Local Then
                'COMISIONES
                If gstrTratamientoContableComisionValorLocal = Valor_Tratamiento_Contable_Gasto Then
                    strIndContableComision = Valor_Tratamiento_Contable_Gasto
                ElseIf gstrTratamientoContableComisionValorLocal = Valor_Tratamiento_Contable_Costo Then
                    strIndContableComision = Valor_Tratamiento_Contable_Costo
                End If
                
                'IMPUESTOS
                If gstrTratamientoContableIGVValorLocal = Valor_Tratamiento_Contable_Gasto Then
                    strIndContableImpuesto = Valor_Tratamiento_Contable_Gasto
                ElseIf gstrTratamientoContableIGVValorLocal = Valor_Tratamiento_Contable_Costo Then
                    strIndContableImpuesto = Valor_Tratamiento_Contable_Costo
                ElseIf gstrTratamientoContableIGVValorLocal = Valor_Tratamiento_Contable_Credito Then
                    strIndContableImpuesto = Valor_Tratamiento_Contable_Credito
                End If
            End If
            
            If strCodOrigen = Codigo_Mercado_Extranjero Then
                'COMISIONES
                If gstrTratamientoContableComisionValorExtranjero = Valor_Tratamiento_Contable_Gasto Then
                    strIndContableComision = Valor_Tratamiento_Contable_Gasto
                ElseIf gstrTratamientoContableComisionValorExtranjero = Valor_Tratamiento_Contable_Costo Then
                    strIndContableComision = Valor_Tratamiento_Contable_Costo
                End If
                
                'IMPUESTOS
                If gstrTratamientoContableIGVValorExtranjero = Valor_Tratamiento_Contable_Gasto Then
                    strIndContableImpuesto = Valor_Tratamiento_Contable_Gasto
                ElseIf gstrTratamientoContableIGVValorExtranjero = Valor_Tratamiento_Contable_Costo Then
                    strIndContableImpuesto = Valor_Tratamiento_Contable_Costo
                ElseIf gstrTratamientoContableIGVValorExtranjero = Valor_Tratamiento_Contable_Credito Then
                    strIndContableImpuesto = Valor_Tratamiento_Contable_Credito
                End If
            End If
            
'            If gstrTratamientoContableComision = Valor_Tratamiento_Contable_Costo Then
'                strIndGasto = Valor_Caracter
'            End If
            
            '*** Asignar Valores ***
            blnEstado = False
            
            strCodAnalitica = Trim(adoRegistro("CodAnalitica"))
            strIndTitulo = Trim(adoRegistro("IndTituloMaestro"))
            strCodTitulo = Trim(adoRegistro("CodTitulo"))
            strCodGarantia = Trim(adoRegistro("CodGarantia"))
            strIndInversion = Valor_Caracter
            strIndGenerado = Valor_Caracter
            strIndCuponCero = Valor_Caracter
            dblPrecioUnitario = CDbl(adoRegistro("PrecioUnitarioMFL1"))
            dblPrecioUnitarioSucio = CDbl(adoRegistro("PrecioUnitarioSucioMFL1"))
            
            'ACR:24/08/2010
            strTipoDocumento = adoRegistro("TipoDocumento")
            strNumDocumento = adoRegistro("NumDocumento")
            'ACR:24/08/2010
            
            '*** Valores Comunes ***
            If strCodTipoOrden = Codigo_Orden_Compra Or strCodTipoOrden = Codigo_Orden_Pacto Or strCodTipoOrden = Codigo_Orden_Compromiso Then
                curCantOrden = CCur(adoRegistro("CantOrden"))
                curValorMovimiento = CCur(adoRegistro("MontoSubTotalMFL1"))
                
                If strIndContableComision = Valor_Tratamiento_Contable_Costo And strIndContableImpuesto = Valor_Tratamiento_Contable_Costo Then
                    curValorMovimiento = CCur(adoRegistro("MontoTotalMFL1"))
                End If
                
                If strIndContableComision = Valor_Tratamiento_Contable_Costo And strIndContableImpuesto <> Valor_Tratamiento_Contable_Costo Then
                    curValorMovimiento = CCur(adoRegistro("MontoTotalMFL1") - adoRegistro("MontoIgvMFL1"))
                End If
            Else
                curCantOrden = CCur(adoRegistro("CantOrden")) * -1
                curValorMovimiento = CCur(adoRegistro("MontoTotalMFL1"))
            End If
            
            curValorNominal = CCur(adoRegistro("ValorNominal"))
            'curCantMovimiento = curCantOrden * curValorNominal
            curVacCorrido = CCur(adoRegistro("VacCorrido"))
'            curValorMovimiento = CCur(adoRegistro("MontoSubTotalMFL1"))
            
            If strCodTipoOrden = Codigo_Orden_Compromiso Then curValorMovimiento = curValorNominal * curCantOrden * dblPrecioUnitario * 0.01
            
            
            strDescripOrden = Trim(adoRegistro("DescripOrden"))
            strCodMoneda = Trim(adoRegistro("CodMoneda"))
            strCodMonedaPago = Trim(adoRegistro("CodMoneda")) 'Trim(adoRegistro("CodMoneda"))
            
            'dblTipoCambio = CDbl(adoRegistro("ValorTipoCambio"))
            
            dblTipoCambio = ObtenerValorTipoCambio(adoRegistro("CodMoneda"), Codigo_Moneda_Local, gstrFechaActual, gstrFechaActual, gstrCodClaseTipoCambioOperacionFondo, Codigo_Valor_TipoCambioCompra)
            
            'dblTipoCambioMonedaPago = CDbl(adoRegistro("ValorTipoCambioMonedaPago"))
            
            dblTipoCambioMonedaPago = ObtenerValorTipoCambio(adoRegistro("CodMoneda"), Codigo_Moneda_Local, gstrFechaActual, gstrFechaActual, gstrCodClaseTipoCambioOperacionFondo, Codigo_Valor_TipoCambioCompra)
            
            dblTipoCambioArbitraje = CStr(ObtenerTipoCambioMoneda("04", "01", adoRegistro("FechaOrden"), adoRegistro("CodMoneda"), Codigo_Moneda_Local, "3"))
                        
            strCodFile = Trim(adoRegistro("CodFile"))
            strCodClaseInstrumento = Trim(adoRegistro("CodDetalleFile"))
            strCodSubDetalleFile = Trim(adoRegistro("CodSubDetalleFile"))
            strFechaOrden = Convertyyyymmdd(adoRegistro("FechaOrden"))
            strFechaGrabar = gstrFechaActual & Space(1) & Format$(Time, "hh:mm")
            
'            If gdatFechaActual >= adoRegistro("FechaLiquidacion") Then
'                strFechaLiquidacion = gstrFechaActual
'            Else
                strFechaLiquidacion = Convertyyyymmdd(adoRegistro("FechaLiquidacion"))
'            End If
            
            strFechaEmision = Convertyyyymmdd(adoRegistro("FechaEmision"))
            strFechaVencimiento = Convertyyyymmdd(adoRegistro("FechaVencimiento"))
            strFechaPago = Convertyyyymmdd(adoRegistro("FechaConfirmacion"))
            strCodEmisor = Trim(adoRegistro("CodEmisor"))
            strCodCiiu = Valor_Caracter
            strCodGrupo = Valor_Caracter
            strCodSector = Valor_Caracter
            strCodTipoTasa = Trim(adoRegistro("TipoTasa"))
            strCodBaseAnual = Trim(adoRegistro("BaseAnual"))
            dblTasaInteres = CDbl(adoRegistro("TasaInteres"))
            
            dblTirBruta = CDbl(adoRegistro("TirBruta"))
            dblTirBrutaLimpia = CDbl(adoRegistro("TirBrutaLimpia"))
            dblTirNeta = CDbl(adoRegistro("TirNeta"))
            
            dblMontoVencimiento = CDbl(adoRegistro("MontoVencimiento"))
            intDiasPlazo = CInt(adoRegistro("CantDiasPlazo"))
            strCodAgente = Trim(adoRegistro("CodAgente"))
            strCodOperacion = Trim(adoRegistro("CodOperacion"))
            strCodNegociacion = Trim(adoRegistro("CodNegociacion"))
            strCodReportado = Trim(adoRegistro("CodReportado"))
            strCodGirador = Trim(adoRegistro("CodGirador"))
            strCodAceptante = Trim(adoRegistro("CodAceptante"))
            strIndCustodia = Trim(adoRegistro("IndCustodia"))
            strIndKardex = Trim(adoRegistro("IndKardex"))
            strTipoTasa = Trim(adoRegistro("TipoTasa"))
            strBaseAnual = Trim(adoRegistro("BaseAnual"))
            strRiesgo = Trim(adoRegistro("TipoRiesgo"))
            strSubRiesgo = Trim(adoRegistro("SubRiesgo"))
            strCodNemonico = Trim(adoRegistro("Nemotecnico"))
            
            curCtaCostoSAB = 0
            curCtaCostoBVL = 0
            curCtaCostoCavali = 0
            curCtaCostoConasev = 0
            curCtaCostoFondoLiquidacion = 0
            curCtaCostoFondoGarantia = 0
            curCtaGastoBancario = 0
            curCtaComisionEspecial = 0
            curCtaImpuesto = 0  'gasto
            'REVISAR
            curCtaImpuestoCredito = 0
            
            curCtaInversionCostoSAB = 0
            curCtaInversionCostoBVL = 0
            curCtaInversionCostoCavali = 0
            curCtaInversionCostoConasev = 0
            curCtaInversionCostoFondoLiquidacion = 0
            curCtaInversionCostoFondoGarantia = 0
            curCtaInversionCostoGastosBancarios = 0
            curCtaInversionCostoComisionEspecial = 0
            curCtaInversionCostoIGV = 0
            
            If strCodTipoOrden = Codigo_Orden_Compra Then
                If strIndContableComision = Valor_Tratamiento_Contable_Costo Then
                    curCtaInversionCostoSAB = CCur(adoRegistro("MontoAgenteMFL1"))
                    curCtaInversionCostoBVL = CCur(adoRegistro("MontoBolsaMFL1"))
                    curCtaInversionCostoCavali = CCur(adoRegistro("MontoCavaliMFL1"))
                    curCtaInversionCostoConasev = CCur(adoRegistro("MontoConasevMFL1"))
                    curCtaInversionCostoFondoLiquidacion = CCur(adoRegistro("MontoFondoLiquidacionMFL1"))
                    curCtaInversionCostoFondoGarantia = CCur(adoRegistro("MontoFondoGarantiaMFL1"))
                    curCtaInversionCostoGastosBancarios = CCur(adoRegistro("MontoGastoBancarioMFL1"))
                    curCtaInversionCostoComisionEspecial = CCur(adoRegistro("MontoComisionEspecialMFL1"))
                ElseIf strIndContableComision = Valor_Tratamiento_Contable_Gasto Then
                    curCtaCostoSAB = CCur(adoRegistro("MontoAgenteMFL1"))
                    curCtaCostoBVL = CCur(adoRegistro("MontoBolsaMFL1"))
                    curCtaCostoCavali = CCur(adoRegistro("MontoCavaliMFL1"))
                    curCtaCostoConasev = CCur(adoRegistro("MontoConasevMFL1"))
                    curCtaCostoFondoLiquidacion = CCur(adoRegistro("MontoFondoLiquidacionMFL1"))
                    curCtaCostoFondoGarantia = CCur(adoRegistro("MontoFondoGarantiaMFL1"))
                    curCtaGastoBancario = CCur(adoRegistro("MontoGastoBancarioMFL1"))
                    curCtaComisionEspecial = CCur(adoRegistro("MontoComisionEspecialMFL1"))
                End If
                
                If strIndContableImpuesto = Valor_Tratamiento_Contable_Costo Then
                    curCtaInversionCostoIGV = CCur(adoRegistro("MontoIgvMFL1"))
                ElseIf strIndContableImpuesto = Valor_Tratamiento_Contable_Credito Then
                    curCtaImpuestoCredito = CCur(adoRegistro("MontoIgvMFL1"))
                ElseIf strIndContableImpuesto = Valor_Tratamiento_Contable_Gasto Then
                    curCtaImpuesto = CCur(adoRegistro("MontoIgvMFL1"))
                End If
                
                
            ElseIf strCodTipoOrden = Codigo_Orden_Venta Then
                curCtaCostoSAB = CCur(adoRegistro("MontoAgenteMFL1"))
                curCtaCostoBVL = CCur(adoRegistro("MontoBolsaMFL1"))
                curCtaCostoCavali = CCur(adoRegistro("MontoCavaliMFL1"))
                curCtaCostoConasev = CCur(adoRegistro("MontoConasevMFL1"))
                curCtaCostoFondoLiquidacion = CCur(adoRegistro("MontoFondoLiquidacionMFL1"))
                curCtaCostoFondoGarantia = CCur(adoRegistro("MontoFondoGarantiaMFL1"))
                curCtaGastoBancario = CCur(adoRegistro("MontoGastoBancarioMFL1"))
                curCtaComisionEspecial = CCur(adoRegistro("MontoComisionEspecialMFL1"))
                
                If strIndContableImpuesto = Valor_Tratamiento_Contable_Credito Then
                    curCtaImpuestoCredito = CCur(adoRegistro("MontoIgvMFL1"))
                ElseIf strIndContableImpuesto = Valor_Tratamiento_Contable_Gasto Then
                    curCtaImpuesto = CCur(adoRegistro("MontoIgvMFL1"))
                Else
                    curCtaImpuesto = CCur(adoRegistro("MontoIgvMFL1"))
                End If
                
            End If
            
            If strCodTipoOrden = Codigo_Orden_Compra Or _
               strCodTipoOrden = Codigo_Orden_Pacto Or _
               strCodTipoOrden = Codigo_Orden_Compromiso Then
                
                strCodOperacionCaja = Codigo_Caja_Compra
                strCodTipoOperacion = Codigo_Caja_Compra
                
                If strCodTipoOrden = Codigo_Orden_Pacto Then
                    strCodOperacionCaja = Codigo_Caja_Compra
                    strCodTipoOperacion = Codigo_Caja_Compra
                ElseIf strCodTipoOrden = Codigo_Orden_Compromiso Then
                    strCodOperacionCaja = Codigo_Caja_Compromiso
                    strCodTipoOperacion = Codigo_Caja_Compromiso
                End If
                
                If strIndTitulo = Valor_Caracter Then
                    '*** Obtener Analítica ***
                    .CommandText = "{call up_ACSelDatosParametro(21,'" & strCodFile & "') }"
                    Set adoTemporal = .Execute
    
                    If Not adoTemporal.EOF Then
                        lngNumAnalitica = CLng(adoTemporal("NumUltimo")) + 1
                        strCodAnalitica = Format$(lngNumAnalitica, "00000000")
                    End If
                    adoTemporal.Close
                                
                    strCodTitulo = Trim(gstrInicialTitulo) & Format$(CLng(gstrPeriodoActual) & CLng(gstrCodAdministradora) & CLng(adoRegistro("CodFondo")) & CLng(adoRegistro("CodFile")) & CLng(strCodAnalitica), String(12, "0"))
                    strIndInversion = "X"
                    strIndCuponCero = "X"
                    strIndGenerado = "X"
                    
                    .CommandText = "SELECT DescripFile FROM InversionFile WHERE CodFile='" & strCodFile & "'"
                    Set adoTemporal = .Execute
                    
                    If Not adoTemporal.EOF Then
                        strDescripTitulo = Trim(adoTemporal("DescripFile"))
                    End If
                    adoTemporal.Close
                    
                    .CommandText = "SELECT DescripPersona,CodCiiu,CodGrupo,CodSector FROM InstitucionPersona WHERE CodPersona='" & adoRegistro("CodEmisor") & "' AND " & _
                        "TipoPersona='" & Codigo_Tipo_Persona_Emisor & "'"
                    Set adoTemporal = .Execute
                    
                    If Not adoTemporal.EOF Then
                        strDescripTitulo = strDescripTitulo & Space(1) & Trim(adoTemporal("DescripPersona"))
                        strCodCiiu = Trim(adoTemporal("CodCiiu"))
                        strCodGrupo = Trim(adoTemporal("CodGrupo"))
                        strCodSector = Trim(adoTemporal("CodSector"))
                    End If
                    adoTemporal.Close
                    
                End If
            Else
            
                If strCodTipoOrden = Codigo_Orden_Venta Then
                    strCodOperacionCaja = Codigo_Caja_Venta
                    strCodTipoOperacion = Codigo_Caja_Venta
                End If
                
                If strCodTipoOrden = Codigo_Orden_Quiebre Then
                    strCodOperacionCaja = Codigo_Caja_PrePago
                    strCodTipoOperacion = Codigo_Caja_PrePago
                End If
            
            End If

            .CommandText = "SELECT IndAmortizacion,IndTasaAjustada,CodMoneda FROM InstrumentoInversion WHERE CodTitulo='" & Trim(adoRegistro("CodTitulo")) & "'"
            Set adoTemporal = .Execute
            If Not adoTemporal.EOF Then
                strIndAmortizacion = Trim(adoTemporal("IndAmortizacion"))
                strFlgTvac = Trim(adoTemporal("IndTasaAjustada"))
                strCodMoneda = Trim(adoTemporal("CodMoneda"))
            Else
                strIndAmortizacion = Valor_Caracter
                strFlgTvac = Valor_Caracter
                strCodMoneda = Trim(adoRegistro("CodMoneda"))
            End If
            adoTemporal.Close
            
            '*** Si no es título inscrito, registrarlo ***
            If strIndTitulo = Valor_Caracter Then
                lngDiasTitulo = DateDiff("d", adoRegistro("FechaEmision"), adoRegistro("FechaVencimiento"))
            
                '*** Datos del título ***
'                strSQLTitulo = "{ call up_IVManInstrumentoInversion('" & strCodTitulo & "','" & _
'                    strCodFile & "','" & strCodAnalitica & "','" & strCodClaseInstrumento & "','" & strCodSubClaseInstrumento & "','" & _
'                    strCodAnalitica & "','" & strCodFondo & "','" & gstrCodAdministradora & "','" & strIndInversion & "','" & strCodNemonico & "','" & strDescripTitulo & "','','" & _
'                    "','" & strFechaEmision & "','" & strFechaVencimiento & "'," & lngDiasTitulo & ",'" & _
'                    strCodEmisor & "','','" & strCodCiiu & "','" & strCodGrupo & "','" & strCodSector & "'," & CDec(curValorNominal) & ",'" & _
'                    strCodMoneda & "','" & strCodMoneda & "',0,0,0,'" & strCodTipoTasa & "','" & Codigo_Tipo_Comision_Fija & "','" & strCodBaseAnual & "','" & _
'                    Codigo_Frecuencia_Anual & "','',''," & CDec(dblTIRNETA) & ",0,0,0,'" & _
'                    "','','','','" & strIndGenerado & "','X','','" & _
'                    "','" & strIndCuponCero & "','01','" & Trim(adoRegistro("TipoRiesgo")) & "','" & Trim(adoRegistro("SubRiesgo")) & "','" & _
'                    gstrLogin & "','" & Convertyyyymmdd(gdatFechaActual) & Space(1) & format$(Time, "hh:mm") & "','" & _
'                    gstrLogin & "','" & Convertyyyymmdd(gdatFechaActual) & Space(1) & format$(Time, "hh:mm") & "','I') }"
                                    
                '*** Calculando factores ***
                If strCodFile = "008" Then
                    dblTasaAnual = dblTirNeta
                    dblFactor1 = dblTirNeta
                    dblFactor2 = dblTirNeta
                Else
                    dblTasaAnual = dblTasaInteres
                    dblFactor1 = dblTasaInteres
                    dblFactor2 = dblTasaInteres
                End If
'                curValorCupon = CDbl(adoRegistro("ValorNominal")) * CDbl(adoRegistro("CantOrden")) * dblTasaAnual * 0.01
                
                If strCodFile = "008" Then
                    dblSaldoAmortizacion = CDbl(adoRegistro("MontoTotalMFL1"))
                Else
                    dblSaldoAmortizacion = CDbl(adoRegistro("ValorNominal")) * CDbl(adoRegistro("CantOrden"))
                End If
                
                '*** Base de Cálculo ***
                intDiasPeriodo = 365
                Select Case strCodBaseAnual
                    Case Codigo_Base_30_360: intDiasPeriodo = 360
                    Case Codigo_Base_Actual_365: intDiasPeriodo = 365
                    Case Codigo_Base_Actual_360: intDiasPeriodo = 360
                    Case Codigo_Base_30_365: intDiasPeriodo = 365
                End Select
                
                '*** Calculando factores ***
                dblFactorAnualCupon = FactorAnual(dblFactor1, intDiasPlazo, intDiasPeriodo, strCodTipoTasa, Valor_Indicador, Codigo_Calculo_Normal, 0, intDiasPlazo, 1)
                dblTasaCuponNormal = FactorAnualNormal(dblFactor2, intDiasPlazo, intDiasPeriodo, strCodTipoTasa, Valor_Indicador, Codigo_Calculo_Normal, 0, intDiasPlazo, 1)
                
                dblFactorDiario = FactorDiario(dblFactorAnualCupon, intDiasPlazo, strCodTipoTasa, Valor_Indicador, intDiasPlazo)
                dblFactorDiarioNormal = FactorDiarioNormal(dblTasaCuponNormal, intDiasPlazo, strCodTipoTasa, Valor_Indicador, intDiasPlazo)
            
                If strCodTipoTasa = Codigo_Tipo_Tasa_Efectiva Then '*** Efectiva ***
                   dblFactorAnualCupon = ((1 + (0.01 * dblFactor1)) ^ (intDiasPlazo / intDiasPeriodo)) - 1
                   dblTasaCuponNormal = ((1 + (0.01 * dblFactor2)) ^ (intDiasPlazo / intDiasPeriodo)) - 1
                   dblFactorDiario = ((1 + dblFactorAnualCupon) ^ (1 / intDiasPlazo)) - 1
                Else '*** Nominal Capitalizable ***
                   dblFactorAnualCupon = (0.01 * dblFactor1) / (intDiasPeriodo / intDiasPlazo)
                   dblTasaCuponNormal = (0.01 * dblFactor2) / (intDiasPeriodo / intDiasPlazo)
                   dblFactorDiario = dblFactorAnualCupon / intDiasPlazo
                End If
                
                
                If adoRegistro("CodFile") = "008" Then
                    dblFactorDiario = FactorDiarioImplicito(adoRegistro("MontoTotalMFL1"), adoRegistro("MontoTotalMFL2"), intDiasPlazo)
                    curValorCupon = CDbl(adoRegistro("MontoTotalMFL1")) + (CDbl(adoRegistro("MontoTotalMFL2")) - CDbl(adoRegistro("MontoTotalMFL1")))
                    dblValorInteres = (CDbl(adoRegistro("MontoTotalMFL2")) - CDbl(adoRegistro("MontoTotalMFL1")))
                Else
                    curValorCupon = CDbl(adoRegistro("ValorNominal")) * CDbl(adoRegistro("CantOrden")) * dblFactorDiario
                    dblValorInteres = dblSaldoAmortizacion * dblFactorAnualCupon
                End If
                
                curValorAmortizacion = curValorCupon - dblValorInteres
                dblSaldoAmortizacion = dblSaldoAmortizacion - curValorAmortizacion
                dblAcumuladoAmortizacion = dblAcumuladoAmortizacion + curValorAmortizacion
            
            End If

            '*** Obtener las cuentas de inversión ***
            Call ObtenerCuentasInversion(strCodFile, strCodClaseInstrumento, Trim(adoRegistro("CodMoneda")), strCodSubDetalleFile)
            
            If ((strFechaLiquidacion = strFechaOrden Or strCodTipoOrden = Codigo_Orden_Venta) And adoRegistro("CodFile") = "005") Or adoRegistro("CodFile") <> "005" Then
                '*** Obtener Inventario Actual del Kardex ***
                '*** NO tomar en cuenta los Mov. Anulados (IndAnulado<>'X') ***
                .CommandText = "SELECT SaldoInicial,SaldoFinal,MontoSaldo,SaldoInteresCorrido,PromedioInteresCorrido,ValorPromedio,SaldoAmortizacion,TirOperacion,TirOperacionLimpia, TirPromedio,TirPromedioLimpia,PrecioPromedio, PrecioPromedioLimpio FROM InversionKardex " & _
                    "WHERE CodAnalitica='" & strCodAnalitica & "' AND CodFile='" & adoRegistro("CodFile") & "' AND " & _
                    "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                    "IndUltimoMovimiento='X' " & _
                    "ORDER BY NumKardex"
                Set adoTemporal = .Execute

                If adoTemporal.EOF Then
                    curKarSldInic = 0: curKarSldFina = 0: curKarValSald = 0: dblKarValProm = 0
                Else
                    curKarSldInic = CCur(adoTemporal("SaldoFinal"))
                    curKarSldFina = CCur(adoTemporal("SaldoFinal"))
                    curKarValSald = CCur(adoTemporal("MontoSaldo"))
                    curKarIadSald = CCur(adoTemporal("SaldoInteresCorrido"))
                    dblKarIadProm = CDbl(adoTemporal("PromedioInteresCorrido"))
                    dblKarValProm = CDbl(adoTemporal("ValorPromedio"))
                    dblKarSldAmort = CCur(adoTemporal("SaldoAmortizacion"))
                    
                    dblTirPromAnt = CDbl(adoTemporal("TirPromedio"))
                    dblTirPromLimpiaAnt = CDbl(adoTemporal("TirPromedioLimpia"))
                    
                    dblKarPrecioPromedio = CDbl(adoTemporal("PrecioPromedio"))
                    dblKarPrecioPromedioLimpio = CDbl(adoTemporal("PrecioPromedioLimpio"))
                    
                   
                End If
                adoTemporal.Close: Set adoTemporal = Nothing

                '*** Verificar en caso de venta si hay suficiente stock ***
                If strCodTipoOrden = Codigo_Orden_Venta Or strCodTipoOrden = Codigo_Orden_Quiebre Or strCodTipoOrden = Codigo_Orden_Prepago Then
                    If curKarSldFina < CCur(adoRegistro("CantOrden")) Then
    '                    gblnErr = True
                        MsgBox "No hay suficiente stock para vender " & CStr(adoRegistro("CantOrden")) & " títulos " & adoRegistro("CodTitulo")
                        Error 100
                    End If
                End If

                curValComi = CCur(adoRegistro("MontoAgenteMFL1")) + CCur(adoRegistro("MontoBolsaMFL1")) + CCur(adoRegistro("MontoConasevMFL1")) + CCur(adoRegistro("MontoCavaliMFL1")) + CCur(adoRegistro("MontoFondoGarantiaMFL1")) + CCur(adoRegistro("MontoFondoLiquidacionMFL1")) + CCur(adoRegistro("MontoGastoBancarioMFL1")) + CCur(adoRegistro("MontoComisionEspecialMFL1")) + CCur(adoRegistro("MontoIgvMFL1"))

                strIndUltimoMovimiento = "X"
                strTipoMovimientoKardex = "E"
                
                If strCodTipoOrden = Codigo_Orden_Venta Or strCodTipoOrden = Codigo_Orden_Quiebre Or strCodTipoOrden = Codigo_Orden_Prepago Then strTipoMovimientoKardex = "S"
                
                If strCodTipoOrden = Codigo_Orden_Compra Or _
                   strCodTipoOrden = Codigo_Orden_Compromiso Then '*** Compra ***
                    dblPrecioUnitario = CDbl(adoRegistro("PrecioUnitarioMFL1"))
                    curValorMovimiento = curValorMovimiento
                Else '*** Venta ***
                    curValorMovimiento = dblKarValProm * curCantOrden
                End If
                
                If curKarSldInic = 0 And strCodTipoOperacion <> Codigo_Orden_Venta Then  '*** Primera Compra ***
                    curSaldoInicialKardex = 0
                    curSaldoFinalKardex = curCantOrden
                    curValorSaldoKardex = curValorMovimiento '+ curValComi
                    curSaldoInteresCorrido = CCur(adoRegistro("InteresCorridoMFL1"))
                    curSaldoAmortizacion = curSaldoFinalKardex
                    'acr
'                    If strIndAmortizacion = Valor_Indicador Then
'                        Set adoTemporal = New ADODB.Recordset
'
'                        .CommandText = "SELECT SUM(ValorAmortizacion) ValorAmortizacion " & _
'                        "FROM InstrumentoInversionCalendario " & _
'                        "WHERE CodTitulo='" & strCodTitulo & "' AND FechaVencimiento<'" & strFechaLiquidacion & "'"
'                        Set adoTemporal = .Execute
'
'                        If Not adoRegistro.EOF Then
'                            If Not IsNull(adoTemporal("ValorAmortizacion")) Then
'                                dblValorAmortizacion = CDbl(adoTemporal("ValorAmortizacion")) * 0.01
'                            End If
'                        End If
'                        adoTemporal.Close: Set adoTemporal = Nothing
'
'                        curSaldoAmortizacion = curSaldoFinalKardex - (curSaldoFinalKardex * CDbl(dblValorAmortizacion))
'                    End If
                Else
                    curSaldoInicialKardex = curKarSldFina
                    curSaldoFinalKardex = curKarSldFina + curCantOrden
                    'acr
'                    If strIndAmortizacion <> Valor_Caracter Then
'                        If dblKarSldAmort = 0 Or IsNull(dblKarSldAmort) Then
'                            curSaldoAmortizacion = curSaldoFinalKardex
'                        Else
'                            curSaldoAmortizacion = dblKarSldAmort + ((curCantOrden / (curKarSldFina / dblKarSldAmort)))
'                        End If
'                    Else
                    curSaldoAmortizacion = curSaldoFinalKardex
                    'End If
                End If
                
                If strCodTipoOrden = Codigo_Orden_Compra Or strCodTipoOrden = Codigo_Orden_Compromiso Then '*** Compra ***
                    If curKarSldFina = 0 Then
                        curSaldoInteresCorrido = CCur(adoRegistro("InteresCorridoMFL1"))
                        curValorSaldoKardex = curValorMovimiento '+ curValComi
                    Else
                        curSaldoInteresCorrido = curKarIadSald + CCur(adoRegistro("InteresCorridoMFL1"))
                        curValorSaldoKardex = curKarValSald + curValorMovimiento '+ curValComi
                    End If
                Else
                    curSaldoInteresCorrido = curKarIadSald + (dblKarIadProm * curCantOrden)
                    If strIndPorcenPrecio = Valor_Indicador Then
                        curValorSaldoKardex = dblKarValProm * curSaldoFinalKardex * adoRegistro("ValorNominal")    'ACR: dblPrecioUnitario * 0.01 * curSaldoFinalKardex
                    Else
                        curValorSaldoKardex = dblKarValProm * curSaldoFinalKardex        'ACR: dblPrecioUnitario * curSaldoFinalKardex
                    End If
                End If
                
                '*** Obtener Saldos ***
'                curCtaInversion = ObtenerSaldoFinalCuenta(strCodFondo, gstrCodAdministradora, strCodFile, strCodAnalitica, strFechaRegistro, strFechaMas1Dia, strCtaInversion, strCodMoneda)
'                curCtaProvInteres = ObtenerSaldoFinalCuenta(strCodFondo, gstrCodAdministradora, strCodFile, strCodAnalitica, strFechaRegistro, strFechaMas1Dia, strCtaProvInteres, strCodMoneda)
'                curCtaInteres = curCtaProvInteres
'                curCtaInteresVencido = curCtaProvInteres
'                curCtaInteresCorrido = ObtenerSaldoFinalCuenta(strCodFondo, gstrCodAdministradora, strCodFile, strCodAnalitica, strFechaRegistro, strFechaMas1Dia, strCtaInteresCorrido, strCodMoneda)
'                curCtaProvFlucMercado = ObtenerSaldoFinalCuenta(strCodFondo, gstrCodAdministradora, strCodFile, strCodAnalitica, strFechaRegistro, strFechaMas1Dia, strCtaProvFlucMercado, strCodMoneda)
'                curCtaFlucMercado = curCtaProvFlucMercado
'                curCtaProvFlucK = ObtenerSaldoFinalCuenta(strCodFondo, gstrCodAdministradora, strCodFile, strCodAnalitica, strFechaRegistro, strFechaMas1Dia, strCtaProvFlucK, strCodMoneda)
'                curCtaFlucK = curCtaProvFlucK
                                
                curSaldoInversion = 0
                curSaldoInversionCostoSAB = 0
                curSaldoInversionCostoBVL = 0
                curSaldoInversionCostoCavali = 0
                curSaldoInversionCostoFondoGarantia = 0
                curSaldoInversionCostoConasev = 0
                curSaldoInversionCostoIGV = 0
                curSaldoInversionCostoCompromiso = 0
                curSaldoInversionCostoResponsabilidad = 0
                curSaldoInversionCostoFondoLiquidacion = 0
                curSaldoInversionCostoComisionEspecial = 0
                curSaldoInversionCostoGastosBancarios = 0
                curSaldoProvInteres = 0
                curSaldoInteres = 0
                curSaldoInteresVencido = 0
                curSaldoInteresCorridoAcum = 0
                curSaldoProvFlucMercado = 0
                curSaldoFlucMercado = 0
                'acr
                curSaldoProvFlucMercadoPerdida = 0
                curSaldoFlucMercadoPerdida = 0
                'acr
                curSaldoProvFlucK = 0
                curSaldoFlucK = 0
                
'                strCtaProvFlucMercadoPerdida = 0
'                strCtaFlucMercadoPerdida = 0

                curSaldoInversion = ObtenerSaldoFinalCuenta(strCodFondo, gstrCodAdministradora, strCodFile, strCodAnalitica, strFechaRegistro, strFechaMas1Dia, strCtaInversion, strCodMoneda)
                curSaldoProvInteres = ObtenerSaldoFinalCuenta(strCodFondo, gstrCodAdministradora, strCodFile, strCodAnalitica, strFechaRegistro, strFechaMas1Dia, strCtaProvInteres, strCodMoneda)
                curSaldoInteres = curSaldoProvInteres
                curSaldoInteresVencido = curSaldoProvInteres
                curSaldoInteresCorridoAcum = ObtenerSaldoFinalCuenta(strCodFondo, gstrCodAdministradora, strCodFile, strCodAnalitica, strFechaRegistro, strFechaMas1Dia, strCtaInteresCorrido, strCodMoneda)
                curSaldoProvFlucMercado = ObtenerSaldoFinalCuenta(strCodFondo, gstrCodAdministradora, strCodFile, strCodAnalitica, strFechaRegistro, strFechaMas1Dia, strCtaProvFlucMercado, strCodMoneda)
                curSaldoFlucMercado = ObtenerSaldoFinalCuenta(strCodFondo, gstrCodAdministradora, strCodFile, strCodAnalitica, strFechaRegistro, strFechaMas1Dia, strCtaFlucMercado, strCodMoneda) 'curSaldoProvFlucMercado
                'acr
                curSaldoProvFlucMercadoPerdida = 0 'ObtenerSaldoFinalCuenta(strCodFondo, gstrCodAdministradora, strCodFile, strCodAnalitica, strFechaRegistro, strFechaMas1Dia, strCtaProvFlucMercadoPerdida, strCodMoneda)
                curSaldoFlucMercadoPerdida = ObtenerSaldoFinalCuenta(strCodFondo, gstrCodAdministradora, strCodFile, strCodAnalitica, strFechaRegistro, strFechaMas1Dia, strCtaFlucMercadoPerdida, strCodMoneda) 'curSaldoProvFlucMercadoPerdida
                'acr
                curSaldoProvFlucK = ObtenerSaldoFinalCuenta(strCodFondo, gstrCodAdministradora, strCodFile, strCodAnalitica, strFechaRegistro, strFechaMas1Dia, strCtaProvFlucK, strCodMoneda)
                curSaldoFlucK = curSaldoProvFlucK

                '*** Obtener Saldos de Comisiones***
                If strIndContableComision = Valor_Tratamiento_Contable_Costo Then
                    curSaldoInversionCostoSAB = ObtenerSaldoFinalCuenta(strCodFondo, gstrCodAdministradora, strCodFile, strCodAnalitica, strFechaRegistro, strFechaMas1Dia, strCtaInversionCostoSAB, strCodMoneda)
                    curSaldoInversionCostoBVL = ObtenerSaldoFinalCuenta(strCodFondo, gstrCodAdministradora, strCodFile, strCodAnalitica, strFechaRegistro, strFechaMas1Dia, strCtaInversionCostoBVL, strCodMoneda)
                    curSaldoInversionCostoCavali = ObtenerSaldoFinalCuenta(strCodFondo, gstrCodAdministradora, strCodFile, strCodAnalitica, strFechaRegistro, strFechaMas1Dia, strCtaInversionCostoCavali, strCodMoneda)
                    curSaldoInversionCostoFondoGarantia = ObtenerSaldoFinalCuenta(strCodFondo, gstrCodAdministradora, strCodFile, strCodAnalitica, strFechaRegistro, strFechaMas1Dia, strCtaInversionCostoFondoGarantia, strCodMoneda)
                    curSaldoInversionCostoConasev = ObtenerSaldoFinalCuenta(strCodFondo, gstrCodAdministradora, strCodFile, strCodAnalitica, strFechaRegistro, strFechaMas1Dia, strCtaInversionCostoConasev, strCodMoneda)
                    curSaldoInversionCostoCompromiso = ObtenerSaldoFinalCuenta(strCodFondo, gstrCodAdministradora, strCodFile, strCodAnalitica, strFechaRegistro, strFechaMas1Dia, strCtaInversionCostoCompromiso, strCodMoneda)
                    curSaldoInversionCostoResponsabilidad = ObtenerSaldoFinalCuenta(strCodFondo, gstrCodAdministradora, strCodFile, strCodAnalitica, strFechaRegistro, strFechaMas1Dia, strCtaInversionCostoResponsabilidad, strCodMoneda)
                    curSaldoInversionCostoFondoLiquidacion = ObtenerSaldoFinalCuenta(strCodFondo, gstrCodAdministradora, strCodFile, strCodAnalitica, strFechaRegistro, strFechaMas1Dia, strCtaInversionCostoFondoLiquidacion, strCodMoneda)
                    curSaldoInversionCostoComisionEspecial = ObtenerSaldoFinalCuenta(strCodFondo, gstrCodAdministradora, strCodFile, strCodAnalitica, strFechaRegistro, strFechaMas1Dia, strCtaInversionCostoComisionEspecial, strCodMoneda)
                    curSaldoInversionCostoGastosBancarios = ObtenerSaldoFinalCuenta(strCodFondo, gstrCodAdministradora, strCodFile, strCodAnalitica, strFechaRegistro, strFechaMas1Dia, strCtaInversionCostoGastosBancarios, strCodMoneda)
                End If
                
                '*** Obtener Saldos de Impuesto si es Costo***
                If strIndContableImpuesto = Valor_Tratamiento_Contable_Costo Then
                    curSaldoInversionCostoIGV = ObtenerSaldoFinalCuenta(strCodFondo, gstrCodAdministradora, strCodFile, strCodAnalitica, strFechaRegistro, strFechaMas1Dia, strCtaInversionCostoIGV, strCodMoneda)
                End If
                                
                If strCodTipoOrden = Codigo_Orden_Compra Or _
                   strCodTipoOrden = Codigo_Orden_Compromiso Then '*** Compra ***
                    If curSaldoInicialKardex > 0 Then
                        '*** Obtener VAN ***
                        
                        If strCodFile = "005" Then
                            If strIndContableComision = Valor_Tratamiento_Contable_Costo And strIndContableImpuesto = Valor_Tratamiento_Contable_Costo Then
                                'curVAN = curSaldoInversion + curSaldoProvInteres + curSaldoInteresCorridoAcum + curSaldoProvFlucMercado + curSaldoProvFlucK + Round((dblPrecioUnitario / 100 * curCantOrden * adoRegistro("ValorNominal")), 2) + curValComi
                                curVAN = curSaldoInversion + curSaldoInteresCorridoAcum + Round((dblPrecioUnitario / 100 * curCantOrden * adoRegistro("ValorNominal")), 2) + adoRegistro("InteresCorridoMFL1") + curValComi
                                curVANLimpio = curSaldoInversion + Round((dblPrecioUnitario / 100 * curCantOrden * adoRegistro("ValorNominal")), 2) + curValComi
                            ElseIf strIndContableComision = Valor_Tratamiento_Contable_Costo And strIndContableImpuesto <> Valor_Tratamiento_Contable_Costo Then
                                curVAN = curSaldoInversion + curSaldoInteresCorridoAcum + Round((dblPrecioUnitario / 100 * curCantOrden * adoRegistro("ValorNominal")), 2) + adoRegistro("InteresCorridoMFL1") + curValComi - adoRegistro("MontoIgvMFL1")
                                curVANLimpio = curSaldoInversion + Round((dblPrecioUnitario / 100 * curCantOrden * adoRegistro("ValorNominal")), 2) + curValComi - adoRegistro("MontoIgvMFL1")
                            ElseIf strIndContableComision = Valor_Tratamiento_Contable_Gasto Then
                                curVAN = curSaldoInversion + curSaldoInteresCorridoAcum + Round((dblPrecioUnitario / 100 * curCantOrden * adoRegistro("ValorNominal")), 2) + adoRegistro("InteresCorridoMFL1")
                                curVANLimpio = curSaldoInversion + Round((dblPrecioUnitario / 100 * curCantOrden * adoRegistro("ValorNominal")), 2)
                            End If
                        End If
                        
                        If strCodFile <> "004" And strCodFile <> "017" Then
                            '*** Convertir ***
                            datFchCalc = Convertddmmyyyy(strFechaLiquidacion)
    
                            '*** Hallar TIR Promedio ***
                            If strCodFile = "005" Then
                                dblTirProm = TirNoPer(strCodTitulo, datFchCalc, datFchCalc, curVAN, 0, (curSaldoFinalKardex * adoRegistro("ValorNominal")), curSaldoFinalKardex, 0.1, "", "", "") 'strCodTipoAjuste, strCodIndiceInicial, strCodIndiceFinal)
                                dblTirPromLimpia = TirNoPer(strCodTitulo, datFchCalc, datFchCalc, curVANLimpio, 0, (curSaldoFinalKardex * adoRegistro("ValorNominal")), curSaldoFinalKardex, 0.1, "", "", "") 'strCodTipoAjuste, strCodIndiceInicial, strCodIndiceFinal)
                                'PRECIO LIMPIO
                                dblPrecioPromedioLimpio = (curVANLimpio) / (curSaldoFinalKardex * adoRegistro("ValorNominal")) * 100
                                'PRECIO SUCIO
                                dblPrecioPromedio = (curVAN) / (curSaldoFinalKardex * adoRegistro("ValorNominal")) * 100
                            End If
                            
                            dblTirNetaKardex = dblTirNeta
                            dblTirOperacionKardex = dblTirBruta  '*** El que se ingresó ***
                            dblTirOperacionKardexLimpia = dblTirBrutaLimpia
                            dblTirPromedioKardex = dblTirProm   '*** Nuevo TIR Promedio ***
                            dblTirPromedioKardexLimpia = dblTirPromLimpia
                            
                        Else
                            dblTirNetaKardex = 0
                            dblTirOperacionKardex = 0
                            dblTirOperacionKardexLimpia = 0
                            dblTirPromedioKardex = 0
                            dblTirPromedioKardexLimpia = 0
                            dblPrecioPromedio = 0
                            dblPrecioPromedioLimpio = 0
                        End If
                    Else 'PRIMERA COMPRA
                        dblTirNetaKardex = dblTirNeta
                        dblTirOperacionKardex = dblTirBruta
                        dblTirOperacionKardexLimpia = dblTirBrutaLimpia
                        dblTirPromedioKardex = dblTirBruta
                        dblTirPromedioKardexLimpia = dblTirBrutaLimpia
                        dblPrecioPromedio = dblPrecioUnitarioSucio
                        dblPrecioPromedioLimpio = dblPrecioUnitario
                    End If
                Else
                    '*** Graba la TIR de la Operación y la TIR PROMEDIO del anterior movimiento ***
                    dblTirNetaKardex = dblTirNeta
                    dblTirOperacionKardex = dblTirOperacion
                    dblTirOperacionKardexLimpia = dblTirOperacionLimpia
                    dblTirPromedioKardex = dblTirPromAnt
                    dblTirPromedioKardexLimpia = dblTirPromLimpiaAnt
                    dblPrecioPromedio = dblKarPrecioPromedio
                    dblPrecioPromedioLimpio = dblKarPrecioPromedioLimpio
                
                End If
    
                If curSaldoFinalKardex <> 0 Then
                    If strCodTipoOrden = Codigo_Orden_Compra Or strCodTipoOrden = Codigo_Orden_Compromiso Then '*** Compra ***
                        If strIndPorcenPrecio = Valor_Indicador Then
                            dblValorPromedioKardex = curValorSaldoKardex / (curSaldoFinalKardex * CCur(adoRegistro("ValorNominal")))
                        Else
                            dblValorPromedioKardex = curValorSaldoKardex / (curSaldoFinalKardex * CCur(adoRegistro("Valornominal")))
                        End If
                        dblInteresCorridoPromedio = curSaldoInteresCorrido / curSaldoFinalKardex
                    Else
                        dblValorPromedioKardex = dblKarValProm
                        dblInteresCorridoPromedio = IIf(IsNull(dblKarIadProm), 0, dblKarIadProm)
                    End If
                Else
                    If strCodTipoOrden = Codigo_Orden_Venta Then '*** Para guardar Saldo del Int. cuando es Vta.Total en caso se desea recuperar ***
                        curSaldoInteresCorrido = IIf(IsNull(curKarIadSald), 0, curKarIadSald)
                    End If
                    dblValorPromedioKardex = 0
                    dblInteresCorridoPromedio = 0
                End If
            End If
                                        
            '*** Asignar valores a los montos del asiento contable ***
            If strCodTipoOrden = Codigo_Orden_Venta Then
                'If curSaldoFinalKardex > 0 Then
                
                    curCtaProvFlucMercado = 0: curCtaProvFlucMercadoPerdida = 0
                    
                    If curSaldoFlucMercado <> 0 Then
                        curCtaFlucMercado = curSaldoFlucMercado * Abs(curCantOrden / curSaldoInicialKardex)  '(Abs(dblValorPromedioKardex * curCantOrden) / curSaldoInversion)
                    End If
                    
                    If curSaldoFlucMercadoPerdida <> 0 Then
                        curCtaFlucMercadoPerdida = curSaldoFlucMercadoPerdida * Abs(curCantOrden / curSaldoInicialKardex) * -1 '(Abs(dblValorPromedioKardex * curCantOrden) / curSaldoInversion)
                    End If
                    
                    curCtaProvFlucK = curSaldoProvFlucK * Abs(curCantOrden / curSaldoInicialKardex)  '(Abs(dblValorPromedioKardex * curCantOrden) / curSaldoInversion)
                    curCtaInversion = Abs(curSaldoInversion * (curCantOrden / curSaldoInicialKardex))
                    
                    If strIndContableComision = Valor_Tratamiento_Contable_Costo Then
                        curCtaInversionCostoSAB = Abs(curSaldoInversionCostoSAB * (curCantOrden / curSaldoInicialKardex))
                        curCtaInversionCostoBVL = Abs(curSaldoInversionCostoBVL * (curCantOrden / curSaldoInicialKardex))
                        curCtaInversionCostoCavali = Abs(curSaldoInversionCostoCavali * (curCantOrden / curSaldoInicialKardex))
                        curCtaInversionCostoFondoGarantia = Abs(curSaldoInversionCostoFondoGarantia * (curCantOrden / curSaldoInicialKardex))
                        curCtaInversionCostoConasev = Abs(curSaldoInversionCostoConasev * (curCantOrden / curSaldoInicialKardex))
                        curCtaInversionCostoCompromiso = Abs(curSaldoInversionCostoCompromiso * (curCantOrden / curSaldoInicialKardex))
                        curCtaInversionCostoResponsabilidad = Abs(curSaldoInversionCostoResponsabilidad * (curCantOrden / curSaldoInicialKardex))
                        curCtaInversionCostoFondoLiquidacion = Abs(curSaldoInversionCostoFondoLiquidacion * (curCantOrden / curSaldoInicialKardex))
                        curCtaInversionCostoComisionEspecial = Abs(curSaldoInversionCostoComisionEspecial * (curCantOrden / curSaldoInicialKardex))
                        curCtaInversionCostoGastosBancarios = Abs(curSaldoInversionCostoGastosBancarios * (curCantOrden / curSaldoInicialKardex))
                    End If
                    
                    If strIndContableImpuesto = Valor_Tratamiento_Contable_Costo Then
                        curCtaInversionCostoIGV = Abs(curSaldoInversionCostoIGV * (curCantOrden / curSaldoInicialKardex))
                    End If

                'End If
            Else
                curCtaInversion = CCur(adoRegistro("MontoSubTotalMFL1"))
            End If
            
'            curCtaCosto = curCtaInversion
'            curCtaFlucMercado = curCtaProvFlucMercado
'            curCtaFlucK = curCtaProvFlucK
            
            curCtaCosto = curCtaInversion + curCtaInversionCostoSAB + curCtaInversionCostoBVL + curCtaInversionCostoCavali + curCtaInversionCostoFondoGarantia + curCtaInversionCostoConasev + curCtaInversionCostoIGV + curCtaInversionCostoCompromiso + curCtaInversionCostoResponsabilidad + curCtaInversionCostoFondoLiquidacion + curCtaInversionCostoComisionEspecial + curCtaInversionCostoGastosBancarios
                        
            
            curCtaProvFlucMercado = curCtaFlucMercado * -1
            curCtaProvFlucMercadoPerdida = curCtaFlucMercadoPerdida * -1
            
            If strCodFile = "005" Then
                curCtaCosto = curCtaInversion + curCtaInversionCostoSAB + curCtaInversionCostoBVL + curCtaInversionCostoCavali + curCtaInversionCostoFondoGarantia + curCtaInversionCostoConasev + curCtaInversionCostoIGV + curCtaInversionCostoCompromiso + curCtaInversionCostoResponsabilidad + curCtaInversionCostoFondoLiquidacion + curCtaInversionCostoComisionEspecial + curCtaInversionCostoGastosBancarios + curCtaProvFlucMercado + curCtaProvFlucMercadoPerdida
            End If
            
'            curCtaFlucMercado = curCtaProvFlucMercado * -1 'curSaldoProvFlucMercado 'curCtaProvFlucMercado
'            curCtaFlucMercadoPerdida = curCtaProvFlucMercadoPerdida * -1 'curSaldoProvFlucMercado 'curCtaProvFlucMercado

            curCtaFlucK = curSaldoProvFlucK * -1 'curCtaProvFlucK
            
            curCtaInteresCorrido = CCur(adoRegistro("InteresCorridoMFL1"))
            curCtaVacCorrido = CCur(adoRegistro("VacCorrido"))
            curCtaXPagar = CCur(adoRegistro("MontoTotalMFL1")) 'CCur(adoRegistro("MontoTotalMonedaPagoMFL1"))
            curCtaXCobrar = CCur(adoRegistro("MontoTotalMFL1")) 'CCur(adoRegistro("MontoTotalMonedaPagoMFL1"))

            curCtaIngresoOperacional = CCur(adoRegistro("MontoSubTotalMFL1")) 'curCtaXCobrar
            
            '*** Transacción ***
            gblnRollBack = False
        
            .CommandText = "{ call up_IVProcOrdenInversion1('" & strCodFondo & "','" & gstrCodAdministradora & "','" & strFechaLiquidacion & "','" & strFechaOrden & "','" & strFechaGrabar & "','" & strFechaMas1Dia & "','" & _
                gstrPeriodoActual & "','" & gstrMesActual & "','" & strCodTipoOrden & "','" & strCodFile & "','" & strCodAnalitica & "','" & strIndTitulo & "','" & strCodTitulo & "','" & strCodClaseInstrumento & "','" & _
                strCodMoneda & "','" & strCodMonedaPago & "','" & strCodMonedaPago & "','" & strNumOrden & "','" & strNumOperacion & "','" & strNumCaja & "','" & strNumKardex & "','" & strNumAsiento & "','" & strpNumCobertura & "','" & strTipoPersona & "','" & strCodPersona & "','" & strCodEmisor & "','" & strCodAgente & "','" & _
                strTipoMovimientoKardex & "'," & CDec(curCantOrden) & "," & CDec(dblPrecioUnitario) & "," & CDec(dblPrecioPromedio) & "," & CDec(dblPrecioPromedioLimpio) & "," & CDec(curValorMovimiento) & "," & CDec(curValComi) & "," & CDec(curSaldoInicialKardex) & "," & CDec(curSaldoFinalKardex) & "," & _
                CDec(curValorSaldoKardex) & ",'" & strDescripOrden & "'," & CDec(dblValorPromedioKardex) & ",'" & strIndUltimoMovimiento & "'," & CDec(dblInteresCorridoPromedio) & "," & CDec(curSaldoInteresCorrido) & "," & CDec(dblTirOperacionKardex) & "," & CDec(dblTirOperacionKardexLimpia) & "," & _
                CDec(dblTirPromedioKardex) & "," & CDec(dblTirPromedioKardexLimpia) & "," & CDec(curVacCorrido) & "," & CDec(dblTirNetaKardex) & "," & CDec(curSaldoAmortizacion) & ",'01','" & strCodSubDetalleFile & "','" & strCodTipoOperacion & "','" & strCodNegociacion & "','" & _
                strCodOrigen & "','" & strCodGarantia & "','" & strFechaVencimiento & "','" & strFechaEmision & "','" & strFechaPago & "'," & CDec(curValorNominal) & "," & CDec(adoRegistro("MontoSubTotalMFL1")) & "," & CDec(adoRegistro("MontoSubTotalMFL1")) & "," & _
                CDec(adoRegistro("InteresCorridoMFL1")) & "," & CDec(adoRegistro("MontoAgenteMFL1")) & "," & CDec(adoRegistro("MontoCavaliMFL1")) & "," & CDec(adoRegistro("MontoConasevMFL1")) & "," & CDec(adoRegistro("MontoBolsaMFL1")) & "," & CDec(adoRegistro("MontoFondoLiquidacionMFL1")) & "," & CDec(adoRegistro("MontoFondoGarantiaMFL1")) & "," & CDec(adoRegistro("MontoGastoBancarioMFL1")) & "," & CDec(adoRegistro("MontoComisionEspecialMFL1")) & "," & CDec(adoRegistro("MontoIgvMFL1")) & "," & CDec(adoRegistro("MontoTotalMFL1")) & "," & CDec(adoRegistro("MontoTotalMFL1")) & "," & _
                CDec(adoRegistro("PrecioUnitarioMFL2")) & "," & CDec(adoRegistro("MontoSubTotalMFL2")) & "," & CDec(adoRegistro("InteresCorridoMFL2")) & "," & CDec(adoRegistro("MontoAgenteMFL2")) & "," & CDec(adoRegistro("MontoCavaliMFL2")) & "," & CDec(adoRegistro("MontoConasevMFL2")) & "," & CDec(adoRegistro("MontoBolsaMFL2")) & "," & CDec(adoRegistro("MontoFondoLiquidacionMFL2")) & "," & CDec(adoRegistro("MontoFondoGarantiaMFL2")) & "," & CDec(adoRegistro("MontoGastoBancarioMFL2")) & "," & CDec(adoRegistro("MontoComisionEspecialMFL2")) & "," & _
                CDec(adoRegistro("MontoIgvMFL2")) & "," & CDec(adoRegistro("MontoTotalMFL2")) & "," & CDec(dblMontoVencimiento) & "," & intDiasPlazo & ",'" & strTipoDocumento & "','" & strNumDocumento & "','" & strCodGrupo & "','" & strCodOperacion & "','" & _
                strCodReportado & "','" & strCodGirador & "','" & strCodAceptante & "','" & strIndCustodia & "','" & strIndKardex & "','" & strTipoTasa & "','" & strBaseAnual & "'," & CDec(dblTasaInteres) & "," & _
                CDec(dblTirBruta) & "," & CDec(dblTirNeta) & ",'" & strRiesgo & "','" & strSubRiesgo & "','" & strObservacion & "','" & strCtaXCobrar & "','" & strCtaXPagar & "','" & strCodOperacionCaja & "'," & _
                CDec(dblTasaAnual) & "," & CDec(dblFactorAnualCupon) & "," & CDec(dblFactorDiario) & "," & CDec(dblValorInteres) & "," & CDec(curValorAmortizacion) & "," & CDec(dblSaldoAmortizacion) & "," & CDec(dblAcumuladoAmortizacion) & "," & _
                CDec(curValorCupon) & "," & CDec(dblTasaCuponNormal) & "," & CDec(dblFactorDiarioNormal) & "," & intCantMovAsiento & ",'" & strIndInversion & "','" & strCodNemonico & "','" & strDescripTitulo & "'," & lngDiasTitulo & ",'" & strCodCiiu & "','" & _
                strCodSector & "','" & strCodTipoTasa & "','" & strCodBaseAnual & "','" & Codigo_Calculo_Normal & "','" & strIndGenerado & "','" & strIndCuponCero & "'," & CDec(curCtaInversion) & "," & CDec(curCtaInversionCostoSAB) & "," & CDec(curCtaInversionCostoBVL) & "," & CDec(curCtaInversionCostoCavali) & "," & CDec(curCtaInversionCostoFondoGarantia) & "," & CDec(curCtaInversionCostoConasev) & "," & CDec(curCtaInversionCostoIGV) & "," & CDec(curCtaInversionCostoCompromiso) & "," & CDec(curCtaInversionCostoResponsabilidad) & "," & CDec(curCtaInversionCostoFondoLiquidacion) & "," & CDec(curCtaInversionCostoComisionEspecial) & "," & CDec(curCtaInversionCostoGastosBancarios) & "," & _
                CDec(curCtaProvInteres) & "," & CDec(curCtaInteres) & "," & _
                CDec(curCtaCosto) & "," & CDec(curCtaIngresoOperacional) & "," & CDec(curCtaInteresVencido) & "," & CDec(curCtaVacCorrido) & "," & CDec(curCtaXPagar) & "," & CDec(curCtaXCobrar) & "," & CDec(curCtaInteresCorrido) & "," & CDec(curCtaProvReajusteK) & "," & _
                CDec(curCtaReajusteK) & "," & CDec(curCtaProvFlucMercado) & "," & CDec(curCtaFlucMercado) & "," & CDec(curCtaProvFlucMercadoPerdida) & "," & CDec(curCtaFlucMercadoPerdida) & "," & CDec(curCtaProvInteresVac) & "," & CDec(curCtaInteresVac) & "," & CDec(curCtaIntCorridoK) & "," & CDec(curCtaProvFlucK) & "," & CDec(curCtaFlucK) & "," & _
                CDec(curCtaInversionTransito) & "," & CDec(curCtaCostoSAB) & "," & CDec(curCtaCostoBVL) & "," & CDec(curCtaCostoCavali) & "," & CDec(curCtaCostoConasev) & "," & CDec(curCtaCostoFondoLiquidacion) & "," & CDec(curCtaCostoFondoGarantia) & "," & CDec(curCtaGastoBancario) & "," & CDec(curCtaComisionEspecial) & "," & CDec(curCtaImpuesto) & "," & CDec(curCtaImpuestoCredito) & "," & CDec(dblTipoCambio) & "," & CDec(dblTipoCambioMonedaPago) & "," & CDec(gdblTasaIgv) & ",'" & gstrLogin & "') }"
            adoConn.Execute .CommandText

        End If
        adoRegistro.Close: Set adoRegistro = Nothing
               
    End With

'    Exit Sub
    
'Ctrl_Error:
'    Select Case Err.Number
'        Case -2147217871 '*** Tiempo de espera ***
'            Resume
'        Case Else
'            gblnRollBack = True
'            Resume Next
'    End Select
  
   
    
End Sub

Public Function convertirTasa(ByVal Tasa As Double, ByVal tipo As String, ByVal periodoInicial As Integer, ByVal periodoFinal As Integer, ByVal indexBaseCalculo As Integer, ByVal unidadesPeriodo As Integer, ByVal fecha As Date, Optional ByVal blnIndCapitalizable As Boolean = False) As Double
    Dim result As Double
    Dim vbase As Double
    Dim vdias As Double
    
    Dim valorParametro As Integer
    Dim valorPeriodoInicial As Integer
    Dim valorPeriodoFinal As Integer
        
    If tipo = "Efectiva" Then
        valorParametro = obtenerValorParametro(Format$(periodoInicial + 1, "00"), "TIPPDC")
        vbase = calcularValorDias(DevolverValorBase(indexBaseCalculo, fecha), valorParametro)
        valorParametro = obtenerValorParametro(Format$(periodoFinal + 1, "00"), "TIPPDC")
        vdias = calcularValorDias(DevolverValorBase(indexBaseCalculo, fecha), valorParametro)
        If unidadesPeriodo <> 0 Then
            vdias = vdias * unidadesPeriodo
        End If
        result = potencia(1 + Tasa, vdias / vbase) - 1
    ElseIf tipo = "Nominal" And blnIndCapitalizable = False Then
        valorPeriodoInicial = obtenerValorParametro("0" & Trim(Str(periodoInicial + 1)), "TIPPDC")
        valorPeriodoFinal = obtenerValorParametro("0" & Trim(Str(periodoFinal + 1)), "TIPPDC")
       'Result = Tasa * valorPeriodoFinal / valorPeriodoInicial
        valorParametro = obtenerValorParametro(Format$(periodoFinal + 1, "00"), "TIPPDC")
        vdias = calcularValorDias(DevolverValorBase(indexBaseCalculo, fecha), valorParametro)
        result = potencia(1 + Tasa / valorPeriodoInicial, vdias) - 1
        If unidadesPeriodo <> 0 Then
            result = result * unidadesPeriodo
        End If
    ElseIf tipo = "Nominal" And blnIndCapitalizable = True Then
        valorPeriodoInicial = obtenerValorParametro("0" & Trim(Str(periodoInicial + 1)), "TIPPDC")
        valorPeriodoFinal = obtenerValorParametro("0" & Trim(Str(periodoFinal + 1)), "TIPPDC")
       'Result = Tasa * valorPeriodoFinal / valorPeriodoInicial
        valorParametro = obtenerValorParametro(Format$(periodoFinal + 1, "00"), "TIPPDC")
        vdias = calcularValorDias(DevolverValorBase(indexBaseCalculo, fecha), valorParametro)
        result = potencia(1 + Tasa / valorPeriodoInicial, unidadesPeriodo) - 1
    End If
    
    convertirTasa = result
End Function

Public Function desplazamientoDiaLaborable(ByVal fecha As Date, ByVal strTipoDesplazamiento As String) As Date
'    *CASOS para tipo
'    * 0 sin despl
'    * 1 sgte dia laborable
'    * 2 sgte dia laborable modificado
'    * 3 dia ant. laborable
'    * 4 dia ant. laborable modificado
'    Dim direccion As Integer '1 hacia adelante, -1 hacia atras
'    Dim result As Date
'    result = fecha
'    If (tipo > 0) And (tipo < 3) Then
'        direccion = 1
'    ElseIf (tipo > 2) And (tipo < 5) Then
'        direccion = -1
'    End If
'    If tipo > 0 Then
'        While (Weekday(result, vbSunday) = vbSaturday) Or (Weekday(result, vbSunday) = vbSunday)
'            result = DateAdd("d", direccion, result)
'            If tipo Mod 2 = 0 Then
'                If Month(result) <> Month(fecha) Then
'                    direccion = direccion * (-1)
'                End If
'            End If
'        Wend
'    End If
'
    Dim adoConsulta As ADODB.Recordset
    Dim strFecha As String
    
    If strTipoDesplazamiento <> Valor_Caracter Then
        strFecha = Convertyyyymmdd(fecha)
    
    With adoComm
            .CommandText = "select dbo.uf_ACObtenerFechaUtil('" & strFecha & "','" & strTipoDesplazamiento & "') as FechaDesplazada"
            Set adoConsulta = .Execute
    End With
    
        desplazamientoDiaLaborable = adoConsulta("FechaDesplazada")
    Else
        desplazamientoDiaLaborable = fecha
    End If
    
End Function
Public Function DevolverValorBase(ByVal indexBaseCalculo As Integer, ByVal fecha As Date) As Double

    Dim valorBase As Integer
    If (indexBaseCalculo < 4) Then
        valorBase = 360 + (indexBaseCalculo Mod 2) * 5
    Else
        If Year(fecha) Mod 4 = 0 Then
            valorBase = 366
        Else
            valorBase = 365
        End If
    End If
    DevolverValorBase = valorBase
End Function

Public Function FormulaCuotaConstante(ByVal Principal As Double, ByVal Tasa As Double, ByVal periodos As Double) As Double
    Dim cuota As Double
    cuota = Principal * (potencia(1 + Tasa, periodos) * Tasa) / (potencia(1 + Tasa, periodos) - 1)
        
    FormulaCuotaConstante = cuota
End Function

'JAFR: Método para generar cuponeras para operaciones 016    Ultima mod: JAFR 14/11/2014
Public Sub GeneraCuponera(ByVal numOrigen As Integer, _
                          ByVal codAnalitica As String, _
                          ByVal CodTitulo As String, _
                          ByVal strCodFondo, _
                          ByVal codAdministradora As String, _
                          ByVal codSolicitud, _
                          ByVal SubDetalleFile As String, _
                          ByVal desembNuevo As Boolean, _
                          ByVal cti_igv As Boolean, _
                          Optional ByVal numDesembolsosOcurridos As Integer, _
                          Optional ByVal fechaDesembolsoNuevo As Date, _
                          Optional ByVal valorDesembolsoNuevo As Double)

    'Variables generales para la definicion del calendario
    Dim numCupon                 As Integer
    Dim numTramo                 As Integer
    Dim fechaInicio              As Date
    Dim fechaCorte               As Date
    Dim fechaCorteAnterior       As Date
    Dim fechaCorteDesplazada     As Date
    Dim fechaCorteDesplazadaAnt  As Date
    Dim fechaPago                As Date
    Dim interes                  As Double
    Dim total                    As Double
    Dim amortizacion             As Double
    Dim acumuladoAmortizacion    As Double
    Dim cuota                    As Double
    Dim saldoDeudorInicial       As Double
    Dim saldoDeudorFinal         As Double
    Dim sumaDigitos              As Double
    Dim blnFaltanCupones         As Boolean
    Dim tasaDiaria               As Double
    Dim Tasa                     As Double
    Dim TasaRecalculadaIgv       As Double
    Dim igvIntereses             As Double
    Dim saldoIgvIntereses        As Double
    Dim SDI_cupant               As Double
    Dim INT_cup                  As Double

    'Arreglos de valores individuales por desembolso, para el caso de desembolsos múltiples
    Dim valorDesembolso()        As Double
    Dim fechaDesemb()            As Double
    Dim SDI()                    As Double
    Dim SDF()                    As Double
    Dim SDIAnterior()            As Double
    Dim SDFAnterior()            As Double
    Dim Inter()                  As Double
    Dim IGVDesemb()              As Double
    Dim Amort()                  As Double
    Dim acumAmort()              As Double
    Dim tempSDI()                As Double
    Dim tempInt()                As Double
    Dim tempSDIAnterior()        As Double
    Dim tempIntAnterior()        As Double
    
    'condiciones financieras tomadas de la BD
    Dim cantCupones              As Integer
    Dim numDesembolsos           As Integer
    Dim cantTramos               As Integer
    Dim fechaEmision             As Date
    Dim fechavencimiento         As Date
    Dim fechaPrimerCorte         As Date
    Dim fechaAPartir             As Date
    Dim Principal                As Double
    Dim tasaIni                  As Double
    Dim strTipoTasa              As String
    Dim unidadesPeriodo          As Integer
    Dim indNumeroCuotas          As Boolean
    Dim indDesembolsosMultiples  As Integer
    Dim indCorteAFinPeriodo      As Boolean
    Dim indPeriodoPersonalizable As Integer
    Dim indCortePrimerCupon      As Integer
    Dim indFechaAPartir          As Integer
    Dim indexTipoAmortizacion    As Integer
    Dim indexDesplazamientoCorte As Integer
    Dim indexDesplazamientoPago  As Integer
    
    
    
    Dim strCodDesplazamientoCorte   As String
    Dim strCodDesplazamientoPago    As String
    
    Dim indexPeriodoTasa         As Integer
    Dim indexTipoCupon           As Integer
    Dim indexBaseCalculo         As Integer
    Dim indexPeriodoCupon        As Integer
    Dim indexUnidadPeriodo       As Integer
    Dim tramoAmortizacion        As Boolean              'tipotramo
    Dim adoRegistroTramo         As New ADODB.Recordset
    Dim adoDesembolsoTmp         As New ADODB.Recordset
    Dim adoRegistroDesembolso    As New ADODB.Recordset
    
    Dim comm                     As ADODB.Command
    Set comm = New ADODB.Command
    
    Dim i As Integer
    
    ReDim listaCupones(cantCupones)
    
    Dim adoRegistroCondicionesFinancieras As New ADODB.Recordset
        
    'query de la tabla instrumentoInversionCondicionesFinancieras
    adoComm.CommandText = "SELECT IndNumCuotas, NumCuotas, FechaEmision, FechaVencimiento, ValorNominal," & _
                        "Tasa, TipoTasa, PeriodoTasa, TipoCupon, BaseCalculo, TipoAmortizacion, PeriodoCupon, " & _
                        "IndPeriodoPersonalizable, CantUnidadesPeriodo, UnidadPeriodo, DesplazamientoCorte, DesplazamientoPago, IndCortePrimerCupon, " & _
                        "FechaPrimerCorte, IndFechaAPartir, FechaAPartir, IndDesembolsosMultiples, CantDesembolsos, CantTramos, " & _
                        "TipoTramo from InstrumentoInversionCondicionesFinancieras where CodTitulo = '" & CodTitulo & "'"
    Set adoRegistroCondicionesFinancieras = adoComm.Execute
    
    'Procesamiento de adoRegistroCondicionesFinancieras
    If adoRegistroCondicionesFinancieras.Fields.Item("IndNumCuotas") = "0" Then
        indNumeroCuotas = False
    Else
        indNumeroCuotas = True
    End If
    
    cantCupones = adoRegistroCondicionesFinancieras.Fields.Item("NumCuotas")
    fechaEmision = adoRegistroCondicionesFinancieras.Fields.Item("FechaEmision")
    fechavencimiento = adoRegistroCondicionesFinancieras.Fields.Item("FechaVencimiento")
    Principal = adoRegistroCondicionesFinancieras.Fields.Item("ValorNominal")
    tasaIni = adoRegistroCondicionesFinancieras.Fields.Item("Tasa")
    
    If adoRegistroCondicionesFinancieras.Fields.Item("TipoTasa") = "01" Then
        strTipoTasa = "Efectiva"
    ElseIf adoRegistroCondicionesFinancieras.Fields.Item("TipoTasa") = "02" Then
        strTipoTasa = "Nominal"
    End If
    
    indexPeriodoTasa = CInt(adoRegistroCondicionesFinancieras.Fields.Item("PeriodoTasa")) - 1
    indexTipoCupon = CInt(adoRegistroCondicionesFinancieras.Fields.Item("TipoCupon")) - 1
    indexBaseCalculo = CInt(adoRegistroCondicionesFinancieras.Fields.Item("BaseCalculo"))
    
    'ajuste de indexbasecalculo
    If indexBaseCalculo = 1 Then
        indexBaseCalculo = 4
    ElseIf indexBaseCalculo > 3 Then
        indexBaseCalculo = indexBaseCalculo - 4
    End If
    
    indexTipoAmortizacion = CInt(adoRegistroCondicionesFinancieras.Fields.Item("TipoAmortizacion")) - 1
    indexPeriodoCupon = CInt(adoRegistroCondicionesFinancieras.Fields.Item("PeriodoCupon")) - 1
    indPeriodoPersonalizable = CInt(adoRegistroCondicionesFinancieras.Fields.Item("IndPeriodoPersonalizable"))
    unidadesPeriodo = adoRegistroCondicionesFinancieras.Fields.Item("CantUnidadesPeriodo")
    indexUnidadPeriodo = adoRegistroCondicionesFinancieras.Fields.Item("UnidadPeriodo") - 1
    
    indexDesplazamientoCorte = CInt(adoRegistroCondicionesFinancieras.Fields.Item("DesplazamientoCorte"))
    strCodDesplazamientoCorte = adoRegistroCondicionesFinancieras.Fields.Item("DesplazamientoCorte")
    
    indexDesplazamientoPago = CInt(adoRegistroCondicionesFinancieras.Fields.Item("DesplazamientoPago"))
    strCodDesplazamientoPago = adoRegistroCondicionesFinancieras.Fields.Item("DesplazamientoPago")
    
    indCortePrimerCupon = adoRegistroCondicionesFinancieras.Fields.Item("IndCortePrimerCupon")
    fechaPrimerCorte = adoRegistroCondicionesFinancieras.Fields.Item("FechaPrimerCorte")
    indFechaAPartir = CInt(adoRegistroCondicionesFinancieras.Fields.Item("IndFechaAPartir"))
    fechaAPartir = adoRegistroCondicionesFinancieras.Fields.Item("FechaAPartir")
    indDesembolsosMultiples = adoRegistroCondicionesFinancieras.Fields.Item("IndDesembolsosMultiples")
    numDesembolsos = adoRegistroCondicionesFinancieras.Fields.Item("CantDesembolsos")
    cantTramos = adoRegistroCondicionesFinancieras.Fields.Item("CantTramos")
    
    If adoRegistroCondicionesFinancieras.Fields.Item("TipoTramo") = "cuota" Then
        tramoAmortizacion = False
    Else
        tramoAmortizacion = True
    End If
    
    'fin
    
    adoComm.CommandText = "SELECT  NumTramo, InicioTramo, FinTramo, Valor From InstrumentoInversionCalendarioTramo where CodTitulo =  '" & CodTitulo & "'"
    Set adoRegistroTramo = adoComm.Execute
    
    adoComm.CommandText = "SELECT NumDesembolso,ValorDesembolso,FechaDesembolso From InstrumentoInversionCalendarioDesembolso where CodTitulo =  '" & CodTitulo & "'"
    'Set adoDesembolsoTmp = adoComm.Execute
    Set adoRegistroDesembolso = adoComm.Execute
    
'    Set adoRegistroDesembolso = New ADODB.Recordset
'
'    With adoRegistroDesembolso
'        .CursorLocation = adUseClient
'        .Fields.Append "NumDesembolso", adInteger, 999
'        .Fields.Append "FechaDesembolso", adDate
'        .Fields.Append "ValorDesembolso", adDouble, 14
'    End With
'
'    adoRegistroDesembolso.Open
'
'    If Not adoDesembolsoTmp.EOF Then
'        adoDesembolsoTmp.MoveFirst
'    End If

    If Not adoRegistroDesembolso.EOF Then
        adoRegistroDesembolso.MoveFirst
    End If

    'Dim i As Integer
'    While Not adoDesembolsoTmp.EOF
'        adoRegistroDesembolso.AddNew Array("NumDesembolso", "FechaDesembolso", "ValorDesembolso"), Array(adoDesembolsoTmp.Fields.Item("NumDesembolso").Value, adoDesembolsoTmp.Fields.Item("FechaDesembolso").Value, adoDesembolsoTmp.Fields.Item("ValorDesembolso").Value)
'        adoDesembolsoTmp.MoveNext
'    Wend
        
'    If desembNuevo Then
'        numDesembolsos = numDesembolsos + 1
'        adoRegistroDesembolso.AddNew
'        adoRegistroDesembolso.Fields.Item("NumDesembolso") = adoRegistroDesembolso.RecordCount + 1
'        adoRegistroDesembolso.Fields.Item("ValorDesembolso") = valorDesembolsoNuevo
'        adoRegistroDesembolso.Fields.Item("FechaDesembolso") = fechaDesembolsoNuevo
'        adoRegistroDesembolso.MoveFirst
'        indDesembolsosMultiples = 1
'    End If
    
    numCupon = 0
    numTramo = 1
    interes = 0
    total = 0
    amortizacion = 0
    acumuladoAmortizacion = 0
    cuota = 0
    saldoDeudorInicial = 0
    saldoDeudorFinal = 0
    sumaDigitos = 0
    igvIntereses = 0
    SDI_cupant = 0 'JJCC
    INT_cup = 0 'JJCC
    
    blnFaltanCupones = True
    
'    If (numOrigen = 1) And (numDesembolsosOcurridos > 0) Then
'        numDesembolsos = numDesembolsosOcurridos
'    End If

    '    If numOrigen = 0 Then
    '        strSQL = "delete from InversionSolicitudCalendarioTmp"
    '        adoComm.CommandText = strSQL
    '        adoComm.Execute
    '    End If
    '
    If indexTipoAmortizacion < 3 Then
        Tasa = tasaIni / 100
        saldoDeudorInicial = Principal
        saldoDeudorFinal = Principal

        If indexTipoAmortizacion = 0 Then
            Tasa = recalcularTasa(tasaIni, strTipoTasa, indexPeriodoTasa, indexPeriodoCupon, indexBaseCalculo, unidadesPeriodo, fechaEmision)
            cuota = CalculoCuotaConstante(CodTitulo, SubDetalleFile, cti_igv, Principal, cantCupones, fechaEmision, fechavencimiento, 0.01)
        ElseIf indexTipoAmortizacion = 1 Then
            sumaDigitos = cantCupones * (cantCupones + 1) / 2
        End If
    End If

    'inicializacion de variables
    fechaInicio = fechaEmision
    fechaCorte = fechaInicio
    fechaCorteDesplazada = fechaInicio
    
    If indexTipoAmortizacion = 3 Then
        adoRegistroTramo.MoveFirst
    End If
    
    'Redimensionamiento de arreglos de variables por desembolso
    ReDim SDI(0 To numDesembolsos - 1)
    ReDim SDF(0 To numDesembolsos - 1)
    ReDim SDIAnterior(0 To numDesembolsos - 1)
    ReDim SDFAnterior(0 To numDesembolsos - 1)
    ReDim Inter(0 To numDesembolsos - 1)
    ReDim Amort(0 To numDesembolsos - 1)
    ReDim acumAmort(0 To numDesembolsos - 1)
    ReDim Amort(0 To numDesembolsos - 1)
    ReDim IGVDesemb(0 To numDesembolsos - 1)
    ReDim fechaDesemb(0 To numDesembolsos - 1)
    ReDim valorDesembolso(0 To numDesembolsos - 1)
    ReDim tempSDI(0 To numDesembolsos - 1)
    ReDim tempInt(0 To numDesembolsos - 1)
    ReDim tempSDIAnterior(0 To numDesembolsos - 1)
    ReDim tempIntAnterior(0 To numDesembolsos - 1)

    If indexTipoAmortizacion = 3 Then
        Dim difDias         As Double
        Dim cantDesembolsos As Integer
        Dim tasaTemp        As Double

        If indDesembolsosMultiples = 1 Then
            cantDesembolsos = numDesembolsos - 1
        Else
            cantDesembolsos = 0
        End If
    End If

    If indexTipoAmortizacion = 3 Then
        Tasa = tasaIni / 100
    End If
    
    While blnFaltanCupones
        saldoDeudorInicial = saldoDeudorFinal
        numCupon = numCupon + 1
        fechaCorteAnterior = fechaCorte
        fechaCorteDesplazadaAnt = fechaCorteDesplazada
        
        'calculo de la fecha de corte
        If (numCupon = 1) And (indCortePrimerCupon = 1) Then
            fechaCorte = fechaPrimerCorte
            fechaCorteAnterior = fechaInicio
            fechaCorteDesplazadaAnt = fechaInicio
            fechaPago = fechaCorte
        Else

            If indCorteAFinPeriodo = True Then
                Dim tempfechainicio As Date
                tempfechainicio = DateAdd("d", 1, fechaInicio)
                tempfechainicio = DateAdd("d", 1, tempfechainicio)

                If numCupon > 1 Then
                    fechaCorte = ultimaFechaPeriodo(tempfechainicio, indexPeriodoCupon)
                Else
                    fechaCorte = ultimaFechaPeriodo(fechaInicio, indexPeriodoCupon)
                End If

            Else
                'JAFR 14/04/2011 Encapsulación de calculo de fecha de corte.
                'fechaCorteAnterior = fechaInicio
                fechaCorte = CalculaFechaSiguienteCalendario(fechaCorteAnterior, indexBaseCalculo, indexPeriodoCupon, indexUnidadPeriodo, unidadesPeriodo)
            End If

            If ((fechaCorte >= fechavencimiento) And Not (indNumeroCuotas = True)) Or ((numCupon = cantCupones) And (indNumeroCuotas = True)) Then
                If ((fechaCorte >= fechavencimiento) And Not (indNumeroCuotas = True)) Then
                    fechaCorte = fechavencimiento
                End If
            End If
        End If

        'fin del cálculo de la fecha de corte
        
        'si no es el primer cupon con fecha de corte especifica....
        'se realiza el desplazamiento especificado tanto para fecha de corte como para fecha de pago
        'If Not ((numCupon = 1)) Then 'And (indCortePrimerCupon = 1)) Then
        fechaCorteDesplazada = desplazamientoDiaLaborable(fechaCorte, strCodDesplazamientoCorte)
        fechaPago = desplazamientoDiaLaborable(fechaCorte, strCodDesplazamientoPago)
            'fechaCorte = fechaCorteDesplazada
        'Else
        '    fechaCorteDesplazada = fechaCorte
        'End If
        
        'Teniendo la fecha de corte desplazada, se calcula tasa entre las dos fechas de corte pertinentes
        Tasa = recalcularTasaDifDias(tasaIni / 100, indexPeriodoTasa, indexBaseCalculo, fechaEmision, indexPeriodoCupon, strTipoTasa, DateDiff("d", fechaCorteDesplazadaAnt, fechaCorteDesplazada))
        
        If fechaCorteDesplazada >= fechavencimiento Then blnFaltanCupones = False
        'hasta aqui ya se tienen calculadas las fechas.
        
        'aqui inicia el cálculo del monto a pagar en el cupon
        Select Case indexTipoAmortizacion

            Case 0
                'Es tasa nominal-efectiva
                'interes = Round(saldoDeudorInicial * ((1 + (Tasa / 30)) ^ (DateDiff("d", fechaCorteAnterior, fechaCorte)) - 1), 2)
                'comentado ACR 28/02/2013
                interes = saldoDeudorInicial * Tasa
                
                If SubDetalleFile = "001" Then

                    'JJCC: Indica si el cálculo incluirá el igv o no
                    If cti_igv Then
                        igvIntereses = Round(interes * gdblTasaIgv, 2)
                    Else
                        igvIntereses = 0
                    End If

                Else
                    igvIntereses = 0
                End If
                
                amortizacion = cuota - (interes + igvIntereses)

            Case 1
                interes = saldoDeudorInicial * Tasa

                If SubDetalleFile = "001" Then
                    If cti_igv Then
                        igvIntereses = interes * gdblTasaIgv
                    Else
                        igvIntereses = 0
                    End If

                Else
                    igvIntereses = 0
                End If
                
                amortizacion = (numCupon / sumaDigitos) * Principal
                cuota = amortizacion + (interes + igvIntereses)

            Case 2
                
                'interes = Round(saldoDeudorInicial * ((1 + (Tasa / 30)) ^ (DateDiff("d", fechaCorteAnterior, fechaCorte)) - 1), 2)
                'comentado ACR 10/03/2013
                interes = saldoDeudorInicial * Tasa
                
                If SubDetalleFile = "001" Then
                    If cti_igv Then
                        igvIntereses = interes * gdblTasaIgv
                    Else
                        igvIntereses = 0
                    End If

                Else
                    igvIntereses = 0
                End If
                
                amortizacion = Principal / cantCupones
                cuota = amortizacion + (interes + igvIntereses)

            Case 3
                'caso de cuotas por tramos aqui
                
                'se redimensionan los arrays
                Dim sumainter     As Double
                Dim sumaigv       As Double
                Dim sumaintereses As Double
                sumainter = 0
                sumaigv = 0
                sumaintereses = 0
                Dim numDesembolso As Integer

                If indDesembolsosMultiples = 1 Then

                    For numDesembolso = 0 To cantDesembolsos '- 1
                        tempSDIAnterior(numDesembolso) = tempSDI(numDesembolso)
                        tempIntAnterior(numDesembolso) = tempInt(numDesembolso)
                        SDIAnterior(numDesembolso) = SDI(numDesembolso)
                        SDFAnterior(numDesembolso) = SDF(numDesembolso)
                    Next

                Else
                    tempSDIAnterior(0) = tempSDI(0)
                    tempIntAnterior(0) = tempInt(0)
                    SDIAnterior(0) = SDI(0)
                    SDFAnterior(0) = SDF(0)
                End If
                
                Dim fechadesembolso As Date
                
                'Se ubica el primer registro de la grilla de desembolsos
                If indDesembolsosMultiples = 1 Then
                    adoRegistroDesembolso.MoveFirst
                End If
                
                'se recorren los registros de desembolsos
                For numDesembolso = 0 To cantDesembolsos '- 1
                    
                    'se obtiene la fecha del desembolso
                    If indDesembolsosMultiples = 1 Then
                        fechadesembolso = adoRegistroDesembolso.Fields.Item("FechaDesembolso").Value
                        valorDesembolso(numDesembolso) = adoRegistroDesembolso.Fields.Item("ValorDesembolso").Value
                    Else
                        fechadesembolso = fechaEmision
                    End If

                    fechaDesemb(numDesembolso) = fechadesembolso
                    
                    If (fechaCorteDesplazada >= fechadesembolso) And (fechaCorteDesplazadaAnt <= fechadesembolso) And (tempSDIAnterior(numDesembolso) = 0) Then

                        'se obtiene el valor del desembolso
                        If indDesembolsosMultiples = 1 Then
                            tempSDI(numDesembolso) = adoRegistroDesembolso.Fields.Item("ValorDesembolso").Value
                        Else
                            tempSDI(numDesembolso) = Principal
                        End If

                        '---------------------------------------
                        valorDesembolso(numDesembolso) = tempSDI(numDesembolso)
                        '---------------------------------------
                    Else
                        tempSDI(numDesembolso) = 0
                    End If
                    
                    If (tempSDIAnterior(numDesembolso) > 0) Then
                        tempInt(numDesembolso) = 0
                    Else
                        difDias = Abs(DateDiff("d", fechadesembolso, fechaCorteDesplazada))
                        
                        Tasa = recalcularTasaDifDias(tasaIni / 100, indexPeriodoTasa, indexBaseCalculo, fechaEmision, indexPeriodoCupon, strTipoTasa, difDias)
                    
                        If (fechaCorteDesplazada >= fechadesembolso) And (fechaCorteDesplazadaAnt <= fechadesembolso) And (tempIntAnterior(numDesembolso) = 0) Then
                            
                            'se calcula el interes de los dias que pasaron desde el desembolso hasta la fecha de corte actual, si es que aplica
                            If indDesembolsosMultiples = 1 Then
                                tempInt(numDesembolso) = adoRegistroDesembolso.Fields.Item("ValorDesembolso").Value * Tasa
                            Else
                                tempInt(numDesembolso) = Principal * Tasa
                            End If

                        Else
                            tempInt(numDesembolso) = 0
                        End If
                    End If
                    
                    'se avanza al siguiente registro de desembolsos
                    If indDesembolsosMultiples = 1 Then
                        If Not adoRegistroDesembolso.EOF Then
                            adoRegistroDesembolso.MoveNext
                        End If
                    End If

                Next
                    
                Dim sumaTempInt As Double
                sumaTempInt = 0

                For i = 0 To cantDesembolsos '- 1
                    saldoDeudorInicial = saldoDeudorInicial + tempSDI(i)
                    sumaTempInt = sumaTempInt + tempInt(i)
                Next
                
                difDias = Abs(DateDiff("d", fechaCorteDesplazada, fechaCorteDesplazadaAnt))
                Tasa = recalcularTasaDifDias(tasaIni / 100, indexPeriodoTasa, indexBaseCalculo, fechaEmision, indexPeriodoCupon, strTipoTasa, difDias)
                
                If sumaTempInt = 0 Then
                    interes = saldoDeudorInicial * Tasa
                Else

                    If numCupon > 1 Then
                        interes = sumaTempInt + saldoDeudorFinal * Tasa
                    Else
                        interes = sumaTempInt
                    End If
                End If
                
                'Se toma el valor del interés del cupón actual para usarlo
                'luego en el cálculo detallado por desembolsos.
                INT_cup = interes
                
                'Se comprueba si el tramo actual corresponde al cupon actual
                If (numCupon > adoRegistroTramo.Fields.Item("FinTramo").Value) Then
                    numTramo = numTramo + 1

                    If Not adoRegistroTramo.EOF Then
                        adoRegistroTramo.MoveNext
                    End If
                End If

                Dim numCuponesDelTramo As Integer
                
                'se obtiene el numero de cupones del tramo actual
                numCuponesDelTramo = adoRegistroTramo.Fields.Item("FinTramo").Value - adoRegistroTramo.Fields.Item("InicioTramo").Value + 1
                
                If tramoAmortizacion Then

                    'caso de especificación de amortización constante por tramos
                    If numTramo < cantTramos Then
                        'se obtiene el valor de la cuota en el tramo
                        amortizacion = adoRegistroTramo.Fields.Item("Valor").Value
                    ElseIf numTramo = cantTramos Then
                        'cálculo de la amortización para el ÚLTIMO tramo
                        difDias = Abs(DateDiff("d", fechaCorteDesplazada, fechaCorteDesplazadaAnt))
                        Tasa = recalcularTasaDifDias(tasaIni / 100, indexPeriodoTasa, indexBaseCalculo, fechaEmision, indexPeriodoCupon, strTipoTasa, difDias)

                        tasaTemp = Tasa
                        amortizacion = saldoDeudorInicial / numCuponesDelTramo
                    End If
                    
                    If SubDetalleFile = "001" Then
                        If cti_igv Then
                            cuota = amortizacion + interes * (1 + gdblTasaIgv)
                        Else
                            cuota = amortizacion + interes
                        End If

                    Else
                        cuota = amortizacion + interes
                    End If
                    
                End If

                If Not tramoAmortizacion Then

                    'caso de especificacion de cuota constante por tramos
                    If numTramo < cantTramos Then
                        'se obtiene el valor de la cuota en el tramo
                        cuota = adoRegistroTramo.Fields.Item("Valor").Value

                        If SubDetalleFile = "001" Then
                            If cti_igv Then
                                amortizacion = cuota - (interes + (interes * gdblTasaIgv))
                            Else
                                amortizacion = cuota - interes
                            End If

                        Else
                            amortizacion = cuota - interes
                        End If

                    ElseIf numTramo = cantTramos Then
                        'cálculo de la cuota para el ÚLTIMO tramo
                        difDias = Abs(DateDiff("d", fechaCorteDesplazada, fechaCorteDesplazadaAnt))
                        Tasa = recalcularTasaDifDias(tasaIni / 100, indexPeriodoTasa, indexBaseCalculo, fechaEmision, indexPeriodoCupon, strTipoTasa, difDias)

                        amortizacion = saldoDeudorInicial
                        tasaTemp = Tasa

                        If tasaTemp = 0 Then
                            cuota = 0
                        Else
                            cuota = CalculoCuotaConstante(CodTitulo, SubDetalleFile, cti_igv, saldoDeudorInicial, numCuponesDelTramo, fechaCorteDesplazadaAnt, fechavencimiento, 0.01)
                        End If
                        
                        amortizacion = saldoDeudorInicial
                        cuota = amortizacion + interes + (interes * gdblTasaIgv)
                        
                        If SubDetalleFile = "001" Then
                            If cti_igv Then
                                cuota = amortizacion + interes + (interes * gdblTasaIgv)
                            Else
                                cuota = amortizacion + interes
                            End If

                        Else
                            cuota = amortizacion + interes
                        End If
                    End If
                End If
                
                If indDesembolsosMultiples = 1 Then
                    adoRegistroDesembolso.MoveFirst
                End If

                'Reconstrucción del For que toma los valores detallados por cupón y
                'número de desembolso.
                For i = 0 To cantDesembolsos '- 1
                    
                    'Se especifican las condiciones para cada caso para poder obtener
                    'los valores que se van a coger dependiendo de los valores asignados
                    'anteriormente.
                    If numCupon = 1 Then
                        If i = 0 Then
                            acumAmort(i) = 0
                            tempSDIAnterior(i) = 0
                            tempIntAnterior(i) = 0
                            SDIAnterior(i) = 0
                            SDFAnterior(i) = 0
                        ElseIf i > 0 Then
                            tempSDIAnterior(i) = tempSDI(i)
                            tempIntAnterior(i) = tempInt(i)
                            SDIAnterior(i) = 0
                            SDFAnterior(i) = 0
                        End If

                    ElseIf numCupon > 1 Then

                        If i = 0 Then
                            tempSDIAnterior(i) = tempSDI(i)
                            tempIntAnterior(i) = tempInt(i)
                            SDIAnterior(i) = SDI(i)
                            SDFAnterior(i) = SDF(i)
                        ElseIf i > 0 Then
                            tempSDIAnterior(i) = tempSDI(i)
                            tempIntAnterior(i) = tempInt(i)
                            SDIAnterior(i) = SDI(i)
                            SDFAnterior(i) = SDF(i)
                        End If
                    End If
                                        
                    'Se obtiene el saldo deudor inicial dependiendo si la fecha de los desembolsos
                    'coincide o no con la fecha de los cupones.
                    If (fechaCorteDesplazada >= fechaDesemb(i)) And (fechaCorteDesplazadaAnt <= fechaDesemb(i)) Then
                        SDIAnterior(i) = 0
                        SDFAnterior(i) = 0

                        If SDFAnterior(i) = 0 Then
                            If indDesembolsosMultiples = 1 Then
                                SDI(i) = adoRegistroDesembolso.Fields.Item("ValorDesembolso").Value
                            Else
                                SDI(i) = Principal
                            End If

                        Else
                            SDI(i) = SDFAnterior(i)
                        End If

                    Else
                        SDI(i) = SDFAnterior(i)
                    End If
                    
                    'Se obtiene el interés dependiendo si la tasa de interés se tiene que recalcular
                    'o si seguirá siendo la misma.
                    If fechaCorteDesplazada > fechaDesemb(i) Then
                        If fechaCorteDesplazadaAnt < fechaDesemb(i) Then
                            'JAFR: Cálculo de la tasa dependiendo de la diferencia de días
                            difDias = Abs(DateDiff("d", fechaCorteDesplazada, fechaDesemb(i)))
                        Else
                            difDias = Abs(DateDiff("d", fechaCorteDesplazada, fechaCorteDesplazadaAnt))
                        End If
                        
                        Tasa = recalcularTasaDifDias(tasaIni / 100, indexPeriodoTasa, indexBaseCalculo, fechaEmision, indexPeriodoCupon, strTipoTasa, difDias)
                        'Fin JAFR.
                        
                        tasaTemp = Tasa
                        Inter(i) = SDI(i) * tasaTemp
                        sumainter = INT_cup
                                                                        
                        If Not tramoAmortizacion Then
                            sumaigv = sumainter * gdblTasaIgv
                            sumaintereses = sumainter + sumaigv
                            
                            If numTramo < cantTramos Then
                                If SubDetalleFile = "001" Then
                                    If cti_igv Then
                                        Amort(i) = (cuota - sumaintereses) * (Inter(i) + (Inter(i) * gdblTasaIgv)) / sumaintereses
                                    Else
                                        Amort(i) = (cuota - sumainter) * Inter(i) / sumainter
                                    End If

                                Else
                                    Amort(i) = (cuota - sumainter) * Inter(i) / sumainter
                                End If

                            ElseIf numTramo = cantTramos Then
                                Amort(i) = SDI(i)
                            End If
                            
                        End If
                        
                        If tramoAmortizacion Then
                            Amort(i) = (cuota - sumainter) * Inter(i) / sumainter
                        End If

                    Else
                        'En caso no se halla dado aún el desembolso, el interés y la amortización
                        'siguen siendo 0
                        Inter(i) = 0
                        Amort(i) = 0
                    End If
                    
                    acumAmort(i) = acumAmort(i) + Amort(i)
                    SDF(i) = SDI(i) - Amort(i)

                    'Se avanza al siguiente registro de desembolsos
                    If indDesembolsosMultiples = 1 Then
                        If Not adoRegistroDesembolso.EOF Then
                            adoRegistroDesembolso.MoveNext
                        End If
                    End If
                                        
                Next

                'Fin JJCC.
        End Select

        saldoDeudorFinal = saldoDeudorInicial - amortizacion
        
        'ajuste de la ultima cuota; no contempla cuotas por tramos
        If cantCupones = numCupon And indexTipoAmortizacion <> 3 Then
            If saldoDeudorFinal <> 0 Then
                amortizacion = amortizacion + saldoDeudorFinal
                cuota = amortizacion + (interes + igvIntereses)
                saldoDeudorFinal = 0
            End If
        End If
        
        SDI_cupant = saldoDeudorInicial 'JJCC: Se toma el valor del SDI del CUPON ANTERIOR.
        
        acumuladoAmortizacion = acumuladoAmortizacion + amortizacion
        
        tasaDiaria = convertirTasa((tasaIni / 100), strTipoTasa, indexPeriodoTasa, 6, indexBaseCalculo, 0, fechaEmision)
        
        Dim cuotaTmp As Double
        cuotaTmp = cuota

        If SubDetalleFile = "001" Then
            If cti_igv Then
                igvIntereses = interes * gdblTasaIgv
                saldoIgvIntereses = interes * gdblTasaIgv
            Else
                igvIntereses = 0
                saldoIgvIntereses = 0
            End If

        Else
            igvIntereses = 0
            saldoIgvIntereses = 0
        End If
        
        Dim intRegistro As Integer

        'almacenamiento de datos en la lista de cupones
        ' If numOrigen = 0 Then
        adoComm.CommandText = "INSERT INTO InversionOperacionCalendarioCuota VALUES ('" & strCodFondo & "','" & codAdministradora & _
                                "','" & codSolicitud & "','','','" & CodTitulo & "','" & CodFile_Descuento_Flujos_Dinerarios & "','" & codAnalitica & _
                                "','" & Format$(numCupon, "000") & "','001',0,'01','" & Convertyyyymmdd(fechaCorteDesplazadaAnt) & "','" & Convertyyyymmdd(fechaPago) & _
                                "','" & Convertyyyymmdd(fechaCorteDesplazada) & "','19000101','19000101'," & DateDiff("d", fechaCorteDesplazadaAnt, fechaCorteDesplazada) & _
                                ",cast(" & saldoDeudorInicial & " as decimal(19,2)),cast(" & amortizacion & " as decimal(19,2)),cast(" & amortizacion & _
                                " as decimal(19,2)),0,0,cast(" & interes & " as decimal(19,2)),0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,'01'," & _
                                gdblTasaIgv * 100 & ",cast(" & igvIntereses & _
                                " as decimal(19,2)),0,0," & gdblTasaIgv * 100 & ",0,0,0,0,0,0,'" & gstrFechaActual & "','" & gstrLogin & "')"
        
        adoConn.Execute adoComm.CommandText, intRegistro
        
        If (indexTipoAmortizacion = 3) And (indDesembolsosMultiples = 1) Then
            Dim fechaIni As Date
            
            For i = 0 To cantDesembolsos

                If fechaDesemb(i) > fechaCorteDesplazadaAnt Then
                    fechaIni = fechaDesemb(i)
                Else
                    fechaIni = fechaCorteDesplazadaAnt
                End If
                
                If SubDetalleFile = "001" Then
                    If cti_igv Then
                        IGVDesemb(i) = Inter(i) * gdblTasaIgv
                    Else
                        IGVDesemb(i) = 0
                    End If

                Else
                    IGVDesemb(i) = 0
                End If
                   
                If numOrigen = 0 Then
                    adoComm.CommandText = "INSERT INTO InversionOperacionCalendarioCuota VALUES ('" & strCodFondo & "','" & codAdministradora & "','" & codSolicitud & _
                                            "','','','" & CodTitulo & "','" & CodFile_Descuento_Flujos_Dinerarios & "','" & codAnalitica & "','" & Format$(numCupon, "000") & _
                                            "','001'," & (i + 1) & ",'01','" & fechaIni & "','" & fechaCorteDesplazada & "','" & fechaPago & "','19000101','19000101'," & _
                                            DateDiff("d", fechaIni, fechaCorteDesplazada) & ",cast(" & SDI(i) & " as decimal(19,2)),cast(" & Amort(i) & _
                                            " as decimal(19,2)),cast(" & Amort(i) & " as decimal(19,2)),0,0,cast(" & Inter(i) & _
                                            " as decimal(19,2)),0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,'01'," & gdblTasaIgv * 100 & _
                                            ",cast(" & IGVDesemb(i) & " as decimal(19,2)),0,0," & gdblTasaIgv * 100 & ",0,0,0,0,0,0,'" & gstrFechaActual & _
                                            "','" & gstrLogin & "')"
                ElseIf numOrigen = 1 Then

                    If numDesembolsos = 1 Then
                        adoComm.CommandText = "INSERT INTO InversionOperacionCalendarioCuota VALUES ('" & strCodFondo & "','" & codAdministradora & "','" & codSolicitud & "','','','" & CodTitulo & "','" & CodFile_Descuento_Flujos_Dinerarios & "','" & codAnalitica & "','" & Format$(numCupon, "000") & "','001'," & (i + 1) & ",'01','" & fechaIni & "','" & fechaCorteDesplazada & "','" & fechaPago & "','19000101'," & DateDiff("d", fechaIni, fechaCorteDesplazada) & ",cast(" & SDI(i) & " as decimal(19,2)),cast(" & Amort(i) & " as decimal(19,2)),cast(" & Amort(i) & " as decimal(19,2)),0,cast(" & interes & " as decimal(19,2)),cast(" & Inter(i) & " as decimal(19,2)),cast(" & Inter(i) & " as decimal(19,2)),0,0,0,0,0,0,0,0,0,0,'01'," & gdblTasaIgv * 100 & ",cast(" & IGVDesemb(i) & " as decimal(19,2)),0," & gdblTasaIgv * 100 & ",0,0,0,0,0,0,'" & gstrFechaActual & "','" & gstrLogin & "')"
                                            
                    ElseIf (numDesembolsos > 1) And (fechaDesembolsoNuevo < fechaCorteDesplazada) Then

                        If i < cantDesembolsos Then
                            adoComm.CommandText = "UPDATE InversionOperacionCalendarioCuota SET " & "MontoPrincipalAdeudado = cast(" & SDI(i) & " as decimal(19,2)), MontoPrincipalCuota = cast(" & Amort(i) & " as decimal(19,2)), MontoPrincipalSecuencial = cast(" & Amort(i) & " as decimal(19,2)), MontoInteresCuota = cast(" & Inter(i) & " as decimal(19,2)), MontoInteresSecuencial = cast(" & Inter(i) & " as decimal(19,2)), MontoImpuestoInteres = cast(" & IGVDesemb(i) & " as decimal(19,2)), InteresAdicAdeudado = 0, InteresAdicionalCuota = 0, InteresAdicionalSecuencial = 0, " & " MontoImpuestoInteresAdic = 0 " & "WHERE CodFondo = '" & strCodFondo & "' and CodAdministradora = '" & codAdministradora & "' and NumOperacionOrig = '" & codSolicitud & "' and NumCuota = '" & Format$(numCupon, "000") & "' and NumDesembolso = " & (i + 1)
                                                    
                        ElseIf i = cantDesembolsos Then
                            adoComm.CommandText = "INSERT INTO InversionOperacionCalendarioCuota VALUES ('" & strCodFondo & "','" & codAdministradora & "','" & codSolicitud & "','','','" & CodTitulo & "','" & CodFile_Descuento_Flujos_Dinerarios & "','" & codAnalitica & "','" & Format$(numCupon, "000") & "','001'," & (i + 1) & ",'01','" & fechaIni & "','" & fechaCorteDesplazada & "','" & fechaPago & "','19000101'," & DateDiff("d", fechaIni, fechaCorteDesplazada) & ",cast(" & SDI(i) & " as decimal(19,2)),cast(" & Amort(i) & " as decimal(19,2)),cast(" & Amort(i) & " as decimal(19,2)),0,cast(" & interes & " as decimal(19,2)),cast(" & Inter(i) & " as decimal(19,2)),cast(" & Inter(i) & " as decimal(19,2)),0,0,0,0,0,0,0,0,0,0,'01'," & gdblTasaIgv * 100 & ",cast(" & IGVDesemb(i) & " as decimal(19,2)),0," & gdblTasaIgv * 100 & ",0,0,0,0,0,0,'" & gstrFechaActual & "','" & gstrLogin & "')"
                        End If
                    End If
                End If
                
                adoConn.Execute adoComm.CommandText, intRegistro
                
            Next

        End If
        
        If numCupon = cantCupones Then
            acumuladoAmortizacion = Principal
            saldoDeudorFinal = 0
        End If
        
        'fechaInicio = fechaCorte
        cuota = cuotaTmp
    Wend
    cantCupones = numCupon
    
    MsgBox "El cronograma de pagos se generó satisfactoriamente.", vbInformation, "Generar cronograma"
    
End Sub


Public Sub GeneraOrdenEvento(lngNroAcuerdo As Long, strCodTitu As String, strCodAnalitica As String, strTipoAcuerdo As String, dblValor As Double, strFchCort As String, blnGenerar As Boolean)

    Dim adoresult           As ADODB.Recordset, adoDatos    As ADODB.Recordset
    Dim strFechaMovimiento  As String
    Dim lngSecEnt           As Long, intRes                 As Integer
        
    '*** FALTA ***
    strFechaMovimiento = Convertyyyymmdd(gdatFechaActual)
        
    With adoComm
        Set adoDatos = New ADODB.Recordset
        
        .CommandText = "SELECT CodFondo FROM Fondo WHERE CodAdministradora='" & gstrCodAdministradora & "' AND TipoCartera='" & Codigo_Valor_RentaVariable & "'"
        Set adoDatos = .Execute
            
        Do While Not adoDatos.EOF
            lngSecEnt = 0
            'lngSecEnt = FObtSec(adoDatos("COD_FOND"), "ENT")
        
            If blnGenerar = True Then
                .CommandText = "SELECT ValorNominal CantAcciones FROM tblCarteraInv WHERE CodTitulo='" & strCodTitu & "' AND " & _
                    "FechaCartera='" & strFchCort & "' AND CodFondo='" & adoDatos("CodFondo") & "' AND " & _
                    "CodAdministradora='" & gstrCodAdministradora & "'"
            Else
                .CommandText = "SELECT SaldoFinal CantAcciones FROM InversionKardex WHERE CodAnalitica='" & strCodAnalitica & "' AND " & _
                    "CodFile='004' AND CodFondo='" & adoDatos("CodFondo") & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                    "IndUltimoMovimiento='X' AND IndAnulado=''"
            End If
            Set adoresult = .Execute
            
            If Not adoresult.EOF Then
                Select Case strTipoAcuerdo
                    Case Codigo_Evento_Liberacion  '*** Ver liberadas ***
                        If dblValor > 0 Then
                            .CommandText = "INSERT INTO EventoCorporativoOrden VALUES ('"
                            .CommandText = .CommandText & adoDatos("CodFondo") & "','" & strCodTitu & ","
                            .CommandText = .CommandText & lngSecEnt & "," & lngNroAcuerdo & ",'"
                            .CommandText = .CommandText & strFechaMovimiento & "','04',"
                            .CommandText = .CommandText & strCodAnalitica & "',"
                            .CommandText = .CommandText & CInt(adoresult("CantAcciones")) & ","
                            .CommandText = .CommandText & (dblValor * 0.01 * CInt(adoresult("CantAcciones"))) & ","
                            .CommandText = .CommandText & "0,0,0,0,0,'"
                            .CommandText = .CommandText & strTipoAcuerdo & "','','"
                            .CommandText = .CommandText & strFechaMovimiento & "','" & gstrLogin & "','','')"
                            adoConn.Execute .CommandText
                    
                            'LGenSec adoDatos("CodFondo"), "ENT"
                        End If
                    Case Codigo_Evento_Dividendo  '*** Ver dividendos ***
                        If (dblValor * CDec(adoresult("CantAcciones"))) > 0 Then
                            .CommandText = "INSERT INTO EventoCorporativoOrden VALUES ('"
                            .CommandText = .CommandText & adoDatos("CodFondo") & "','" & strCodTitu & ","
                            .CommandText = .CommandText & lngSecEnt & "," & lngNroAcuerdo & ",'"
                            .CommandText = .CommandText & strFechaMovimiento & "','04',"
                            .CommandText = .CommandText & strCodAnalitica & "',"
                            .CommandText = .CommandText & CInt(adoresult("CantAcciones")) & ",0,"
                            .CommandText = .CommandText & (dblValor * CInt(adoresult("CantAcciones"))) & ","
                            .CommandText = .CommandText & "0,0,0,0,'"
                            .CommandText = .CommandText & strTipoAcuerdo & "','','"
                            .CommandText = .CommandText & strFechaMovimiento & "','" & gstrLogin & "','','')"
                            adoConn.Execute .CommandText
                    
                            'LGenSec adoDatos("CodFondo"), "ENT"
                        End If
                End Select
            End If
            adoresult.Close: Set adoresult = Nothing

            adoDatos.MoveNext
        Loop
        adoDatos.Close: Set adoDatos = Nothing
    End With
     
End Sub

Public Sub GenerarAsientoDividendo(ByVal strpCodFile As String, _
                                   ByVal strpCodAnalitica As String, _
                                   ByVal strpCodFondo As String, _
                                   ByVal strpCodAdministradora, _
                                   ByVal strpCodDetalleFile As String, _
                                   ByVal strpCodDinamica As String, _
                                   ByVal curpMontoSubtotal As Currency, _
                                   ByVal curpMontoComision As Currency, _
                                   ByVal strpCodMoneda As String, _
                                   ByVal strpDescripAsiento As String, _
                                   ByVal strpCodModulo As String, _
                                   ByVal strpCodTitulo As String, _
                                   ByVal datpFechaEntrega As Date, _
                                   ByVal strpTipoContraparte As String, _
                                   ByVal strpCodContraparte As String, _
                                   ByVal strpCodEmisor As String)
    'GenerarAsientoDividendo(strCodFile                     , strCodAnalitica                , strCodFondo,                 gstrCodAdministradora,       strCodDetalleFile,                  Codigo_Dinamica_Dividendos,      CCur(txtValor.Text),                 CCur(txtValorComision.Text),         strCodMoneda,                  Trim(tdgConsulta.Columns(5).Value), frmMainMdi.Tag,                Trim(tdgConsulta.Columns(7).Value),  dtpFechaEntrega.Value,    Codigo_Tipo_Persona_Agente,          strCodAgente,                      strCodEmisor)
    Dim adoRegistro              As ADODB.Recordset
    Dim curMontoMovimientoMN     As Currency, curMontoMovimientoME       As Currency
    Dim curMontoContable         As Currency
    Dim strIndDebeHaber          As String, strDescripMovimiento         As String
    Dim strNumAsiento            As String, strFechaGrabar               As String
    Dim strFechaCierre           As String, strFechaSiguiente            As String
    Dim strSQLOrdenCaja          As String, strSQLOrdenCajaDetalle       As String
    Dim strSQLOperacion          As String
    Dim intCantRegistros         As Integer
    Dim strIndUltimoMovimiento   As String
    Dim curMontoTotal            As Currency
    Dim strFechaEntrega          As String
    Dim strCodAgente             As String
    Dim dblpTipoCambio           As Double
    Dim dblValorTipoCambio       As Double, strTipoDocumento             As String
    Dim strNumDocumento          As String, strTipoPersonaContraparte    As String
    Dim strCodPersonaContraparte As String
    Dim strIndContracuenta       As String, strCodContracuenta           As String
    Dim strCodFileContracuenta   As String, strCodAnaliticaContracuenta  As String
    strCodAgente = strpCodContraparte 'Se asume que la contraparte es agente
    curMontoTotal = curpMontoSubtotal - curpMontoComision

    With adoComm
        '*** Obtener Secuenciales ***
        strNumCaja = ObtenerSecuencialInversionOperacion(strpCodFondo, Valor_NumOrdenCaja)
        strNumAsiento = ObtenerSecuencialInversionOperacion(strpCodFondo, Valor_NumComprobante)
        strNumOperacion = ObtenerSecuencialInversionOperacion(strpCodFondo, Valor_NumOperacion)
        strFechaEntrega = Convertyyyymmdd(datpFechaEntrega)
        strFechaCierre = Convertyyyymmdd(gdatFechaActual)
        strFechaSiguiente = Convertyyyymmdd(DateAdd("d", 1, gdatFechaActual))
        strFechaGrabar = strFechaCierre & Space(1) & Format$(Time, "hh:mm")
        .CommandText = "SELECT dbo.uf_ACObtenerTipoCambioMoneda1('" & gstrCodClaseTipoCambioOperacionFondo & "','" & Codigo_Valor_TipoCambioCompra & "','" & strFechaCierre & "','" & strpCodMoneda & "','" & Codigo_Moneda_Local & "',5) AS 'ValorTipoCambio'"
        Set adoRegistro = .Execute

        If Not adoRegistro.EOF Then
            dblpTipoCambio = adoRegistro("ValorTipoCambio")
        Else
            dblpTipoCambio = 1

        End If

        adoRegistro.Close
        .CommandText = "SELECT COUNT(*) NumRegistros FROM DinamicaContable " & "WHERE TipoOperacion='" & strpCodDinamica & "' AND CodFile='" & strpCodFile & "' AND (CodDetalleFile='" & strpCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & strpCodAdministradora & "' AND CodMoneda = '" & IIf(strpCodMoneda = Codigo_Moneda_Local, Codigo_Moneda_Local, Codigo_Moneda_Extranjero) & "'"
        Set adoRegistro = .Execute

        If Not adoRegistro.EOF Then
            If CInt(adoRegistro("NumRegistros")) > 0 Then
                intCantRegistros = CInt(adoRegistro("NumRegistros"))

            End If

        End If

        adoRegistro.Close
        '*** Cabecera ***
        .CommandText = "{ call up_ACAdicAsientoContable('" & strpCodFondo & "','" & strpCodAdministradora & "','" & strNumAsiento & "','" & strFechaGrabar & "','" & gstrPeriodoActual & "','" & gstrMesActual & "','','" & strpDescripAsiento & "','" & strpCodMoneda & "','" & Codigo_Moneda_Local & "','',''," & CDec(curMontoTotal) & ",'" & Estado_Activo & "'," & intCantRegistros & ",'" & strFechaGrabar & "','" & strpCodModulo & "',''," & CDec(gdblTipoCambio) & ",'','','" & strpDescripAsiento & "','','X','') }"
        adoConn.Execute .CommandText
        '*** Detalle Contable ***
        .CommandText = "SELECT NumSecuencial,TipoCuentaInversion,IndDebeHaber,DescripDinamica,CodCuenta FROM DinamicaContable " & "WHERE TipoOperacion='" & strpCodDinamica & "' AND CodFile='" & strpCodFile & "' AND (CodDetalleFile='" & strpCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "' AND CodMoneda = '" & IIf(strpCodMoneda = Codigo_Moneda_Local, Codigo_Moneda_Local, Codigo_Moneda_Extranjero) & "' ORDER BY NumSecuencial"
        Set adoRegistro = .Execute

        Do While Not adoRegistro.EOF

            Select Case Trim(adoRegistro("TipoCuentaInversion"))

            Case Codigo_CtaIngresoOperacional
                curMontoMovimientoMN = curpMontoSubtotal

            Case Codigo_CtaXCobrarDividendos
                curMontoMovimientoMN = curMontoTotal
                strCtaXCobrar = Trim(adoRegistro("CodCuenta"))

            Case Codigo_CtaComision
                curMontoMovimientoMN = curpMontoComision

            End Select

            strIndDebeHaber = Trim(adoRegistro("IndDebeHaber"))

            If strIndDebeHaber = "H" Then
                curMontoMovimientoMN = curMontoMovimientoMN * -1

                If curMontoMovimientoMN > 0 Then strIndDebeHaber = "D"
            ElseIf strIndDebeHaber = "D" Then

                If curMontoMovimientoMN < 0 Then strIndDebeHaber = "H"

            End If

            If strIndDebeHaber = "T" Then
                If curMontoMovimientoMN > 0 Then
                    strIndDebeHaber = "D"
                Else
                    strIndDebeHaber = "H"

                End If

            End If

            strDescripMovimiento = Trim(adoRegistro("DescripDinamica"))
            curMontoMovimientoME = 0
            curMontoContable = curMontoMovimientoMN

            If strpCodMoneda <> Codigo_Moneda_Local Then
                curMontoContable = Round(curMontoMovimientoMN * dblpTipoCambio, 2)
                curMontoMovimientoME = curMontoMovimientoMN
                curMontoMovimientoMN = 0

            End If

            If CInt(adoRegistro("NumSecuencial")) = 3 Then
                strIndUltimoMovimiento = "X"
            Else
                strIndUltimoMovimiento = ""

            End If

            dblValorTipoCambio = 1
            strTipoDocumento = ""
            strNumDocumento = ""
            strTipoPersonaContraparte = Codigo_Tipo_Persona_Agente
            strCodPersonaContraparte = strpCodContraparte
            strIndContracuenta = ""
            strCodContracuenta = ""
            strCodFileContracuenta = ""
            strCodAnaliticaContracuenta = ""
            '*** Movimiento ***
            .CommandText = "{ call up_ACAdicAsientoContableDetalle('" & strNumAsiento & "','" & strpCodFondo & "','" & gstrCodAdministradora & "'," & CInt(adoRegistro("NumSecuencial")) & ",'" & strFechaGrabar & "','" & gstrPeriodoActual & "','" & gstrMesActual & "','" & strDescripMovimiento & "','" & strIndDebeHaber & "','" & Trim(adoRegistro("CodCuenta")) & "','" & strpCodMoneda & "'," & CDec(curMontoMovimientoMN) & "," & CDec(curMontoMovimientoME) & "," & CDec(curMontoContable) & "," & dblValorTipoCambio & ",'" & strpCodFile & "','" & strpCodAnalitica & "','" & strTipoDocumento & "','" & strNumDocumento & "','" & strTipoPersonaContraparte & "','" & strCodPersonaContraparte & "','" & strIndContracuenta & "','" & strCodContracuenta & "','" & strCodFileContracuenta & "','" & strCodAnaliticaContracuenta & "','" & strIndUltimoMovimiento & "') }"
            adoConn.Execute .CommandText

            '*** Validar valor de cuenta contable ***
            If Trim(adoRegistro("CodCuenta")) = Valor_Caracter Then
                MsgBox "Registro Nro. " & CStr(adoRegistro("NumSecuencial")) & " de Asiento Contable no tiene cuenta asignada", vbCritical, "Asiento Impuesto"
                gblnRollBack = True
                Exit Sub

            End If

            adoRegistro.MoveNext
        Loop
        adoRegistro.Close: Set adoRegistro = Nothing
        '*** Actualizar el n?mero del parámetro **
        .CommandText = "{ call up_ACActUltNumero('" & strpCodFondo & "','" & strpCodAdministradora & "','" & Valor_NumComprobante & "','" & strNumAsiento & "') }"
        adoConn.Execute .CommandText
        'strFechaEntrega
        '*** Orden de Cobro/Pago ***
        '*** Cabecera ***
        strSQLOrdenCaja = "{ call up_ACAdicMovimientoFondo('" & strpCodFondo & "','" & gstrCodAdministradora & "','" & strNumCaja & "','" & strFechaGrabar & "','" & strpCodModulo & "','','" & strFechaEntrega & "','" & "','','','','S','" & strCtaXCobrar & "'," & CDec(curMontoTotal) & ",'" & strpCodFile & "','" & strpCodAnalitica & "','" & strpCodMoneda & "','" & strpDescripAsiento & "','" & Codigo_Caja_Dividendos & "','','" & Estado_Caja_NoConfirmado & "','','','','','','','','','',0,'" & strpCodContraparte & "','" & strpTipoContraparte & "','" & gstrLogin & "') }"
        '*** Detalle ***
        strSQLOrdenCajaDetalle = "{ call up_ACAdicMovimientoFondoDetalle('" & strpCodFondo & "','" & gstrCodAdministradora & "','" & strNumCaja & "','" & strFechaGrabar & "',1,'" & strpCodModulo & "','" & strpDescripAsiento & "','" & "H','" & strCtaXCobrar & "'," & CDec(curMontoTotal) * -1 & ",'" & strpCodFile & "','" & strpCodAnalitica & "','','','" & strpCodMoneda & "','') }"
        '*** Guardar orden de Cobro/Pago ***
        adoConn.Execute strSQLOrdenCaja
        adoConn.Execute strSQLOrdenCajaDetalle
        '****** Prepara informacion para completar en la liqjuidacion asiento de reclasificacion de dividendos devengados a percibidos
        '*** Detalle Contable ***
        .CommandText = "SELECT NumSecuencial,IndDebeHaber,DescripDinamica,CodCuenta FROM DinamicaContable " & "WHERE TipoOperacion='" & Codigo_Dinamica_Dividendos_Percibidos & "' AND CodFile='" & strpCodFile & "' AND (CodDetalleFile='" & strpCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "' AND CodMoneda = '" & IIf(strpCodMoneda = Codigo_Moneda_Local, Codigo_Moneda_Local, Codigo_Moneda_Extranjero) & "' ORDER BY NumSecuencial"
        Set adoRegistro = .Execute

        Do While Not adoRegistro.EOF
            '*** Detalle ***
            strSQLOrdenCajaDetalle = "{ call up_ACAdicMovimientoFondoDetalle('" & strpCodFondo & "','" & gstrCodAdministradora & "','" & strNumCaja & "','" & strFechaGrabar & "'," & CInt(adoRegistro("NumSecuencial")) + 1 & ",'" & strpCodModulo & "','" & Trim(adoRegistro("DescripDinamica")) & "','" & Trim(adoRegistro("IndDebeHaber")) & "','" & Trim(adoRegistro("CodCuenta")) & "'," & IIf(Trim(adoRegistro("IndDebeHaber")) = "D", CDec(curpMontoSubtotal), CDec(curpMontoSubtotal) * -1) & ",'" & strpCodFile & "','" & strpCodAnalitica & "','','','" & strpCodMoneda & "','') }"
            adoConn.Execute strSQLOrdenCajaDetalle
            adoRegistro.MoveNext
        Loop
        adoRegistro.Close: Set adoRegistro = Nothing
        '*** Actualizar el número del parámetro ***
        .CommandText = "{ call up_ACActUltNumero('" & strpCodFondo & "','" & gstrCodAdministradora & "','" & Valor_NumOrdenCaja & "','" & strNumCaja & "') }"
        adoConn.Execute .CommandText
        '*** Operación de Inversión ***
        strSQLOperacion = "{ call up_IVAdicInversionOperacion('" & strpCodFondo & "','" & gstrCodAdministradora & "','" & strNumOperacion & "','" & strFechaGrabar & "','" & strpCodTitulo & "','" & gstrPeriodoActual & "','" & gstrMesActual & "','','" & Estado_Activo & "','" & strpCodAnalitica & "','" & strpCodFile & "','" & strpCodAnalitica & "','" & strpCodDetalleFile & "','','" & Codigo_Caja_Dividendos & "','" & Codigo_Operacion_Contado & "','','" & strpDescripAsiento & "','" & strpCodEmisor & "','" & strCodAgente & "','','" & strFechaGrabar & "','" & strFechaGrabar & "','" & Convertyyyymmdd(Valor_Fecha) & "','" & strpCodMoneda & "','" & strpCodMoneda & "','" & strpCodMoneda & "'," & "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0," & CDec(curMontoTotal) & "," & CDec(curMontoTotal) & ",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,'" & "X','" & strNumAsiento & "','','','','','','','',0,'','','','','',0,0,0,'','','','" & gstrLogin & "') }"
        '*** Guardar operación ***
        adoConn.Execute strSQLOperacion
        '*** Actualizar el número del parámetro ***
        .CommandText = "{ call up_ACActUltNumero('" & strpCodFondo & "','" & gstrCodAdministradora & "','" & Valor_NumOperacion & "','" & strNumOperacion & "') }"
        adoConn.Execute adoComm.CommandText

    End With

End Sub

Public Function obtenerValorParametro(ByVal codparametro As String, ByVal nemotecnico As String) As Integer
    Dim strSQL As String
    strSQL = "{ call Sp_GNAyudaBusqueda(1,'" & codparametro & "','" & nemotecnico & "') }"
    Dim adoBusqueda As ADODB.Recordset
    Dim valorResult As Long
    Set adoBusqueda = New ADODB.Recordset
    adoComm.CommandText = strSQL
    Set adoBusqueda = adoComm.Execute
    Do Until adoBusqueda.EOF
        valorResult = adoBusqueda("ValorParametro")
        adoBusqueda.MoveNext
    Loop
    adoBusqueda.Close: Set adoBusqueda = Nothing
    obtenerValorParametro = valorResult
End Function

Public Function potencia(ByVal base As Double, ByVal exponente As Double) As Double
    Dim result As Double
    result = Math.Exp(exponente * Math.Log(base))
    potencia = result
End Function

Public Function recalcularTasa(ByVal tasaIni As Double, ByVal strTipoTasa As String, ByVal indexPeriodoTasa As Integer, ByVal indexPeriodoCupon As Integer, ByVal indexBaseCalculo As Integer, ByVal unidadesPeriodo As Integer, ByVal fechaEmision As Date) As Double
    Dim rTasa As Double
    Dim valorParametro As Integer
    Dim vbase As Integer
        
    valorParametro = obtenerValorParametro(Format$(indexPeriodoTasa + 1, "00"), "TIPPDC")
    vbase = calcularValorDias(DevolverValorBase(indexBaseCalculo, fechaEmision), valorParametro)
    
    If indexPeriodoCupon < 7 Then
        rTasa = convertirTasa((tasaIni / 100), strTipoTasa, indexPeriodoTasa, indexPeriodoCupon, indexBaseCalculo, 0, fechaEmision)
    Else
        rTasa = convertirTasa((tasaIni / 100), strTipoTasa, indexPeriodoTasa, indexPeriodoCupon, indexBaseCalculo, unidadesPeriodo, fechaEmision)
    End If
    recalcularTasa = rTasa
End Function

Public Function recalcularTasaDifDias(ByVal Tasa As Double, ByVal indexPeriodoTasa As Integer, ByVal indexBaseCalculo As Integer, ByVal fechaEmision As Date, ByVal periodoInicial As Integer, ByVal tipo As String, ByVal difDias As Integer) As Double
    Dim result As Double
    Dim vbase As Double
    If tipo = "Efectiva" Then
        Dim valorParametro As Integer
        valorParametro = obtenerValorParametro(Format$(indexPeriodoTasa + 1, "00"), "TIPPDC")
        vbase = calcularValorDias(DevolverValorBase(indexBaseCalculo, fechaEmision), valorParametro)
        result = potencia(1 + Tasa, difDias / vbase) - 1
    ElseIf tipo = "Nominal" Then
        Dim valorPeriodoInicial As Integer
        Dim valorPeriodoFinal As Integer
        valorPeriodoInicial = obtenerValorParametro("0" & Trim(Str(indexPeriodoTasa + 1)), "TIPPDC")
        result = potencia(1 + Tasa / valorPeriodoInicial, difDias) - 1
    End If
    recalcularTasaDifDias = result
End Function

Public Function ultimaFechaPeriodo(ByVal fecha As Date, ByVal periodo As Integer) As Date
    Dim result As Date
    result = fecha
    If periodo < 6 Then
        Dim dia As Integer
        Dim mes As Integer
        Dim anho As Integer
        Dim tmp As Integer
        
        dia = Day(fecha)
        mes = Month(fecha)
        anho = Year(fecha)
        
        If dia > 27 Then
            result = DateAdd("d", -1, result)
        End If
        tmp = 0
        Select Case periodo
            Case 0
                tmp = 12 - (mes Mod 12)
            Case 1
                tmp = 6 - (mes Mod 6)
            Case 2
                tmp = 3 - (mes Mod 3)
            Case 3
                tmp = 2 - (mes Mod 2)
        End Select
        result = DateAdd("m", tmp, result)
        
        If periodo = 5 Then
            If dia < 15 Then
                dia = 15
                result = Convertddmmyyyy(Format$(anho, "0000") & Format$(mes, "00") & Format$(dia, "00"))
            Else
                result = UltimaFechaMes(Month(result), Year(result))
            End If
        Else
            result = UltimaFechaMes(Month(result), Year(result))
        End If
    End If
    ultimaFechaPeriodo = result
End Function
Public Function ValorActual(curMontoVencimiento As Currency, dblTasaRendimiento As Double, intBaseCalculo As Integer, intNumDias As Integer) As Currency

    '*** Parámetros : Valor al vencimiento                          ***
    '***              Tasa de rendimiento esperada                  ***
    '***              Base anual de la TRE                          ***
    '***              Plazo (Fecha Vencimiento - Fecha Operación)   ***
    Dim curValorActual As Currency
    
    curValorActual = 0
    If curMontoVencimiento > 0 And dblTasaRendimiento > 0 And intNumDias > 0 And intBaseCalculo > 0 Then
        curValorActual = curMontoVencimiento / ((1 + dblTasaRendimiento * 0.01) ^ (intNumDias / intBaseCalculo))
    End If
    
    ValorActual = curValorActual

End Function

Public Function ValorTasa(curMontoVencimiento As Currency, curMontoLimpio As Currency, intBaseCalculo As Integer, intNumDias As Integer) As Double

    '*** Parámetros : Valor al vencimiento                          ***
    '***              Valor Limpio                                  ***
    '***              Base anual de la TRE                          ***
    '***              Plazo (Fecha Vencimiento - Fecha Operación)   ***
    Dim dblTasaDescuento As Double
    
    dblTasaDescuento = 0
    If curMontoVencimiento > 0 And curMontoLimpio > 0 And intNumDias > 0 And intBaseCalculo > 0 Then
        dblTasaDescuento = ((curMontoVencimiento / curMontoLimpio) ^ (intBaseCalculo / intNumDias)) - 1
    End If
    
    ValorTasa = dblTasaDescuento

End Function

Public Function ValorVencimiento(curpMontoNominal As Currency, dblpTasaInteres As Double, intpBaseCalculo As Integer, intpNumDias As Integer, intpNumDias30 As Integer, strpTipoTasa As String, strpBaseCalculo As String) As Currency

    '*** Parámetros : Monto Nominal                                             ***
    '***              Tasa de interés expresada en términos anuales             ***
    '***              Base anual (360 - 365)                                    ***
    '***              Plazo (Fecha Vencimiento - Fecha Emisión)                 ***
    '***              Plazo Días 30 (Fecha Vencimiento - Fecha Emisión)         ***
    '***              Tipo de Tasa (Efectiva - Nominal)                         ***
    '***              Base de Cálculo (Act/360,Act/365,30/360,30/365,Act/Act)   ***
    
    Dim curValorVencimiento As Currency, dblFactorCalculo   As Double
    
    curValorVencimiento = 0
    If curpMontoNominal > 0 And dblpTasaInteres > 0 And intpNumDias > 0 And intpBaseCalculo > 0 Then
        If strpTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
            If strpBaseCalculo = Codigo_Base_30_360 Or strpBaseCalculo = Codigo_Base_30_365 Then
                dblFactorCalculo = ((1 + dblpTasaInteres * 0.01) ^ (intpNumDias30 / intpBaseCalculo))
            Else
                dblFactorCalculo = ((1 + dblpTasaInteres * 0.01) ^ (intpNumDias / intpBaseCalculo))
            End If
        Else
            If strpBaseCalculo = Codigo_Base_30_360 Or strpBaseCalculo = Codigo_Base_30_365 Then
                dblFactorCalculo = 1 + (((dblpTasaInteres * 0.01) / intpBaseCalculo) * intpNumDias30)
            Else
                dblFactorCalculo = 1 + (((dblpTasaInteres * 0.01) / intpBaseCalculo) * intpNumDias)
            End If
        End If
        curValorVencimiento = Round(curpMontoNominal * dblFactorCalculo, 2)
    End If
    
    ValorVencimiento = curValorVencimiento

End Function



