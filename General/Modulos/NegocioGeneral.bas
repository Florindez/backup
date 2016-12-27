Attribute VB_Name = "NegocioGeneral"
Option Explicit


Private Function CalculaCapitalAjusteVAC(strpFechaLiquidacion As String, strpFechaEmision As String, dblpValorNominal As Double, strpFechaIniCupon As String, strpFechaFinCupon As String, strpTipoVac As String, strpTipoCalculo As String, intpDiasCupon As Integer, intpDiasCorridos As Integer, dblpTasaCupon As Double, intpBaseAnual As Integer) As Double

    '*** Cálculo del Valor Nominal Reajustado para el Cálculo del ***
    '*** Interés Corrido y el VAC Corrido para Bonos VAC          ***
    Dim adoRecTasa              As ADODB.Recordset
    Dim dblVacEmision           As Double, dblVacLiquidacion        As Double
    Dim dblVacIniCupon          As Double, dblVacFinCupon           As Double
    Dim dblCapitalVac           As Double, dblFactor                As Double
    Dim strFechaEmisionMas1     As String, strFechaLiquidacionMas1  As String
    Dim strFechaInicuponMas1    As String, strFechaFinCuponMas1     As String

    CalculaCapitalAjusteVAC = 0
    
    dblVacEmision = 0: dblVacLiquidacion = 0: dblVacIniCupon = 0: dblVacFinCupon = 0
    dblCapitalVac = 0
    
    strFechaEmisionMas1 = DateAdd("d", 1, Convertddmmyyyy(strpFechaEmision))
    strFechaLiquidacionMas1 = DateAdd("d", 1, Convertddmmyyyy(strpFechaLiquidacion))
    strFechaInicuponMas1 = DateAdd("d", 1, Convertddmmyyyy(strpFechaIniCupon))
    strFechaFinCuponMas1 = DateAdd("d", 1, Convertddmmyyyy(strpFechaFinCupon))
    
    With adoComm
        Set adoRecTasa = New ADODB.Recordset

        '*** Obtener las Tasas VAC: Emisión, Liquidación, Cupón Inicial y Cupón Final ***
        .CommandText = "SELECT FechaRegistro,ValorTasa FROM InversionTasa " & _
            "WHERE CodTasa='" & Codigo_Tipo_Ajuste_Vac & "' AND " & _
            "((FechaRegistro>='" & strpFechaEmision & "' AND FechaRegistro<'" & strFechaEmisionMas1 & "') OR " & _
            "(FechaRegistro>='" & strpFechaLiquidacion & "' AND FechaRegistro<'" & strFechaLiquidacionMas1 & "') OR " & _
            "(FechaRegistro>='" & strpFechaIniCupon & "' AND FechaRegistro<'" & strFechaInicuponMas1 & "') OR " & _
            "(FechaRegistro>='" & strpFechaFinCupon & "'AND FechaRegistro<'" & strFechaFinCuponMas1 & "'))"
        Set adoRecTasa = .Execute
        
        Do While Not adoRecTasa.EOF
            Select Case Convertyyyymmdd(adoRecTasa("FechaRegistro"))
                Case strpFechaEmision                  '*** A la fecha de emisión          ***
                    dblVacEmision = adoRecTasa("ValorTasa")
                    If strpFechaEmision = strpFechaLiquidacion Then
                        dblVacLiquidacion = adoRecTasa("ValorTasa")
                    End If
                Case strpFechaLiquidacion                    '*** A la fecha de liquidación      ***
                    dblVacLiquidacion = adoRecTasa("ValorTasa")
                    If strpFechaIniCupon = strpFechaLiquidacion Then
                        dblVacIniCupon = adoRecTasa("ValorTasa")
                    End If
                    If strpFechaFinCupon = strpFechaLiquidacion Then
                        dblVacFinCupon = adoRecTasa("ValorTasa")
                    End If
                Case strpFechaIniCupon                   '*** A la fecha de inicio del cupón ***
                    dblVacIniCupon = adoRecTasa("ValorTasa")
                Case strpFechaFinCupon                   '*** A la fecha de corte del cupón  ***
                    dblVacFinCupon = adoRecTasa("ValorTasa")
            End Select
            adoRecTasa.MoveNext
        Loop
        adoRecTasa.Close: Set adoRecTasa = Nothing
    End With

    '*** REVISAR ***
    '*** Si es VAC Periodico proyectar Tasa VAC a la fecha de liquidación ***
    If strpTipoVac = Codigo_Tipo_Ajuste_Vac Then
        If strpTipoCalculo = Codigo_Vac_Emision Then '*** A partir del cupón anterior ***
            If dblVacIniCupon > 0 And intpDiasCupon > 0 Then
                dblFactor = (dblVacFinCupon / dblVacIniCupon) ^ (intpDiasCorridos / intpDiasCupon) - 1
                dblVacLiquidacion = dblFactor
            Else
                MsgBox "La Tasa VAC a la Fecha de Liquidación No Existe, la Operación no se puede realizar.", vbCritical, "Aviso"
                dblCapitalVac = 0: Exit Function
            End If
        Else                                    '*** A partir del cupón vigente ***
            If dblVacLiquidacion = 0 Then
                MsgBox "La Tasa VAC a la Fecha de Liquidación No Existe, la Operación no se puede realizar.", vbCritical, "Aviso"
                dblCapitalVac = 0: Exit Function
            Else
                If dblVacEmision > 0 Then
                    dblVacLiquidacion = (dblVacFinCupon / dblVacEmision)
                Else
                    MsgBox "La Tasa VAC a la Fecha de EMISION No Existe, la Operación no se puede realizar.", vbCritical, "Aviso"
                    dblCapitalVac = 0: Exit Function
                End If
            End If
        End If
    End If

    '*** REVISAR ***
    '*** Cálculo del Valor Nominal Reajustado a la fecha de liquidación ***
    If dblVacEmision > 0 And dblVacLiquidacion > 0 Then
        If strpTipoVac = Codigo_Tipo_Ajuste_Vac Then           '*** VAC Periodico      ***
            dblCapitalVac = Round(dblpValorNominal * dblVacLiquidacion, 2)
        Else                           '*** VAC Al Vencimiento ***
            dblCapitalVac = Round(dblpValorNominal * (dblVacLiquidacion / dblVacEmision), 2)
        End If
    Else
        dblCapitalVac = 0
        MsgBox "Falta Registrar el VAC de la LIQUIDACION DE LA ORDEN o el VAC de EMISION del Título", vbCritical, "Aviso"
    End If

    CalculaCapitalAjusteVAC = dblCapitalVac
    
End Function

Public Function CalculoAjusteVAC(strpCodTitulo As String, dblpCantidad As Double, datpFechaEmision As Date, datpFechaLiquidacion As Date, strpCuponCalculo As String, strpTipoTasa As String, strpPeriodoPago As String, strpTipoVac As String, intpBase As Integer, intpDiasDeRenta As Integer) As Double

    CalculoAjusteVAC = 0
    '*** Cálculo de los Intereses Corridos a la Fecha de la Liquidación                       ***
    '*** Si NumCupon='001' AND FechaLiquidacion = FechaInicio Cupón Vigente Días Corridos = 0 ***
    '*** SiNo  Días Corridos = (FechaLiquidacion) - (FechaInicio Cupón Vigente) + 1           ***
    '*** Los Intereses Corridos en Bonos se calculan sobre el Valor Nominal                   ***
    '*** y el las Letras Hipotecarias sobre el Saldo por Amortizar.                           ***
    '*** Para el caso de Bonos VAC se calcula sobre el Valor Nominal Ajustado                 ***
        
    Dim adoRegistroBono     As ADODB.Recordset, adoRegistroTmp  As ADODB.Recordset
    Dim strFechaLiquidacion As String
    Dim intDiaTranscurridos As Integer, intDiasCUP              As Integer
    Dim intDiasCorridos     As Integer, intRes                  As Integer
    Dim intDiasPeriodo      As Integer
    Dim dblIntCorr          As Double, dblCapitalRea            As Double
    Dim dblTasDia           As Double, dblVacCorrido            As Double
    Dim strFechaInicioCupon As String, strFechaFinCupon         As String
    Dim datFechaInicioCupon As Date, datFechaFinCupon           As Date
        
    strFechaLiquidacion = Convertyyyymmdd(datpFechaLiquidacion)
    
    With adoComm
        Set adoRegistroBono = New ADODB.Recordset
        .CommandType = adCmdText
        
        '*** Obtener datos del cupón vigente ***
        .CommandText = "SELECT FactorDiario,FechaInicio,FechaVencimiento,NumCupon,CantDiasPeriodo " & _
            "FROM InstrumentoInversionCalendario WHERE CodTitulo='" & strpCodTitulo & "' AND IndVigente='X'"
        Set adoRegistroBono = .Execute
        
        If Not adoRegistroBono.EOF Then
            '*** Fecha de inicio del cupón ***
            strFechaInicioCupon = Convertyyyymmdd(adoRegistroBono("FechaInicio"))
            datFechaInicioCupon = adoRegistroBono("FechaInicio")
            datFechaFinCupon = adoRegistroBono("FechaVencimiento")
            intDiasPeriodo = adoRegistroBono("CantDiasPeriodo")
                   
            '*** Días corridos entre el inicio del cupón y la fecha de liquidación ***
            If (strFechaInicioCupon = strFechaLiquidacion) And (adoRegistroBono("NumCupon") = "001") Then
                intDiasCorridos = 0
                intDiaTranscurridos = 0
            Else
                If adoRegistroBono("NumCupon") = "001" Then
                   intDiasCorridos = DateDiff("d", datFechaInicioCupon, datpFechaLiquidacion)
                   intDiaTranscurridos = DateDiff("d", datFechaInicioCupon, datpFechaLiquidacion)
                Else
                   intDiasCorridos = DateDiff("d", datFechaInicioCupon, datpFechaLiquidacion) + 1
                   intDiaTranscurridos = DateDiff("d", datFechaInicioCupon, datpFechaLiquidacion) + 1
                End If
            End If
        
            '*** Obtención de parametros para Bonos VAC ***
                '*** REVISAR ***
                If strpCuponCalculo = Codigo_Tipo_Ajuste_Vac Then   '*** Bonos VAC Periodicos: Cálculo a partir del cupón anterior ***
                    Set adoRegistroTmp = New ADODB.Recordset
                    
                    If CInt(adoRegistroBono("NumCupon")) = 1 Then
                        '*** Primer cupón construir las fechas y días del cupón ***
                        .CommandText = "SELECT ValorParametro FROM AuxiliarParametro WHERE CodParametro='" & strpPeriodoPago & "' AND CodTipoParametro='TIPFRE'"
                        Set adoRegistroTmp = .Execute
                        
                        If Not adoRegistroTmp.EOF Then
                            intDiasPeriodo = CInt(adoRegistroTmp("ValorParametro"))
                        End If
                        adoRegistroTmp.Close: Set adoRegistroTmp = Nothing
    
                        datFechaFinCupon = DateAdd("d", -1, Convertddmmyyyy(datpFechaEmision))
                        datFechaInicioCupon = DateAdd("m", Int(intDiasPeriodo / 30) * -1, datFechaFinCupon)
                        intDiasCUP = DateDiff("d", datFechaInicioCupon, datFechaFinCupon) + 1
                    Else
                        '*** Cualquier otro cupón: extraer los datos del cupón anterior ***
                        .CommandText = "SELECT FechaInicio,FechaVencimiento,CantDiasPeriodo " & _
                            "FROM InstrumentoInversionCalendario WHERE CodTitulo='" & strpCodTitulo & "' AND NumCupon='" & Format(CInt(adoRegistroBono("NumCupon")) - 1, "000") & "'"
                        Set adoRegistroTmp = .Execute
                     
                        If Not adoRegistroTmp.EOF Then
                            datFechaInicioCupon = DateAdd("d", -1, adoRegistroTmp("FechaInicio"))
                            datFechaFinCupon = adoRegistroTmp("FechaVencimiento")
                            intDiasCUP = CInt(adoRegistroTmp("CantDiasPeriodo"))
                        End If
                        adoRegistroTmp.Close: Set adoRegistroTmp = Nothing
                    End If
                Else
                    strFechaFinCupon = Convertyyyymmdd(adoRegistroBono("FechaVencimiento"))
                End If
    
                strFechaInicioCupon = Convertyyyymmdd(datFechaInicioCupon)
                strFechaFinCupon = Convertyyyymmdd(datFechaFinCupon)
        End If
        adoRegistroBono.Close
                        
        .CommandText = "SELECT CodFile,NumCupon,FechaInicio,FactorDiario,ValorAmortizacion,SaldoAmortizacion,TasaInteres,FactorInteres1,CantDiasPeriodo,FechaVencimiento " & _
            "FROM InstrumentoInversionCalendario WHERE CodTitulo='" & strpCodTitulo & "' AND IndVigente='X'"
        Set adoRegistroBono = .Execute
        
        If Not adoRegistroBono.EOF Then
            If adoRegistroBono("CodFile") = "005" Then  '*** Bonos ***
                Dim curIntCapRea    As Currency, curDifReaCap   As Currency
                
                '*** Cálculo automático de Intereses Corridos ***
                    '*** REVISAR ***
                    If strpTipoVac = Codigo_Tipo_Ajuste_Vac Then  '*** Factor diario Bonos VAC Periodicos ***
                        If (adoRegistroBono("FactorInteres1") = 0) Or (adoRegistroBono("FactorInteres1") = Null) Then
                            MsgBox "Cupón Vigente no tiene factor del periodo sin VAC, VERIFIQUE.", vbCritical, "Aviso"
                            dblTasDia = 0: dblIntCorr = 0
                            Exit Function
                        End If
                        dblTasDia = ((1 + adoRegistroBono("FactorInteres1")) ^ (1 / adoRegistroBono("CantDiasPeriodo"))) - 1
                    Else '*** Factor Diario Bonos VAC Al Vcto. ***
                        dblTasDia = adoRegistroBono("FactorDiario")
                    End If
                 
                    '*** Cálculo del Capital Nominal Reajustado para todos los Bonos ***
                    dblCapitalRea = CalculaCapitalAjusteVAC(strFechaLiquidacion, Convertyyyymmdd(datpFechaEmision), dblpCantidad, strFechaInicioCupon, strFechaFinCupon, strpTipoVac, strpCuponCalculo, intDiasPeriodo, intpDiasDeRenta, adoRegistroBono("TasaInteres") * 0.01, intpBase)
                 
                    '*** Cálculo de la diferencia del Capital Reajustado ***
                    If dblCapitalRea > 0 Then
                        curDifReaCap = dblCapitalRea - dblpCantidad
                    End If
                 
                    '*** REVISAR ***
                    If strpTipoVac = Codigo_Tipo_Ajuste_Vac Then
                        '*** VAC Corrido Adelantado ***
                        dblVacCorrido = 0
                        '*** Interés Corrido del Capital Reajustado para todos los Bonos ***
                        curIntCapRea = dblCapitalRea * ((1 + dblTasDia) ^ intDiaTranscurridos - 1)
                        dblIntCorr = dblpCantidad * ((1 + dblTasDia) ^ intDiaTranscurridos - 1)
                        curIntCapRea = curIntCapRea - dblIntCorr
                        '*** Vac Corrido ***
'                        dblIntCorr = Round(dblIntCorr + curIntCapRea + curDifReaCap, 2)
                        dblVacCorrido = curDifReaCap - dblIntCorr
                    ElseIf strpTipoVac = "V" Then
                        '*** VAC Corrido Adelantado ***
                        If IsNumeric(dblpCantidad) Then
                           dblVacCorrido = Format(curDifReaCap, "0.00")
                        Else
                           dblVacCorrido = 0
                        End If
                        '*** Interés Corrido del Capital Reajustado ***
                        curIntCapRea = dblCapitalRea * ((1 + dblTasDia) ^ intDiaTranscurridos - 1)
                        dblIntCorr = dblpCantidad * ((1 + dblTasDia) ^ intDiaTranscurridos - 1)
                        curIntCapRea = curIntCapRea - dblIntCorr
                        '*** Interés Corrido ***
                        dblIntCorr = Round(dblIntCorr + curIntCapRea, 2)
                    End If
            Else '*** Letras Hipotecarias ***
                Dim curSaldoXAmor   As Currency
                '*** VAC Corrido Adelantado ***
                dblVacCorrido = 0
                '*** Interés Corrido ***
                dblIntCorr = Format((adoRegistroBono("ValorAmortizacion") + adoRegistroBono("SaldoAmortizacion")) * ((1 + adoRegistroBono("FactorDiario")) ^ intDiaTranscurridos - 1), "0.000000")
                curSaldoXAmor = adoRegistroBono("ValorAmortizacion") + adoRegistroBono("SaldoAmortizacion")
                
            End If
        Else
            dblIntCorr = 0: dblVacCorrido = 0
        End If
        adoRegistroBono.Close: Set adoRegistroBono = Nothing
    End With
                    
    CalculoAjusteVAC = dblCapitalRea

End Function



Public Sub GenerarRestriccionFondo(strCodFondo As String)

    Dim adoRegistro As ADODB.Recordset
       
    Set adoRegistro = New ADODB.Recordset
   
    adoComm.CommandText = "SELECT CodFile FROM InversionFile WHERE IndInstrumento='X'"
    Set adoRegistro = adoComm.Execute
   
    Do While Not adoRegistro.EOF
    
        adoComm.CommandText = "INSERT INTO RestriccionFondo (CodFondo,CodAdministradora,CodFile,PorcenRestriccion) VALUES ('" & strCodFondo & "','" & gstrCodAdministradora & "','" & adoRegistro("CodFile") & "',0)"
        adoConn.Execute adoComm.CommandText
        
        adoRegistro.MoveNext
        
    Loop

End Sub

Public Sub GenerarPeriodosPagoFondo(strpCodFondo As String, intpNumPeriodos As Integer)

    Dim intContador As Integer, dblPorcenPago   As Double
    Dim datFecha    As Date
       
    If intpNumPeriodos = 1 Then Exit Sub
    
    dblPorcenPago = Round(100 / intpNumPeriodos, 2)
    datFecha = gdatFechaActual
    
    For intContador = 1 To intpNumPeriodos
    
        adoComm.CommandText = "INSERT INTO FondoPagoSuscripcion VALUES ('" & _
            strpCodFondo & "','" & gstrCodAdministradora & "'," & intContador & ",'" & _
            Convertyyyymmdd(datFecha) & "','" & Convertyyyymmdd(datFecha) & "'," & _
            "0,0)"
        adoConn.Execute adoComm.CommandText
    
        datFecha = DateAdd("d", 1, datFecha)
    Next
    
End Sub
Public Function VerificarOperacionRetencion(strFondo As String, strFecha As String, strfechaMas1Dia As String) As Boolean

    Dim adoRegistro As ADODB.Recordset
    
    VerificarOperacionRetencion = False
    
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
'        .CommandText = "SELECT FMMOVCTA.NRO_MCTA, FMMOVCTA.VAL_MOVI from FMMOVCTA, FMOPERAC, FMPERSON, FMPARTIC WHERE "
'        .CommandText = .CommandText & "FMMOVCTA.COD_FOND = FMOPERAC.COD_FOND AND FMMOVCTA.NRO_OPER  = FMOPERAC.NRO_OPER AND "
'        .CommandText = .CommandText & "FMMOVCTA.NRO_FOLI = FMOPERAC.NRO_FOLI AND FMMOVCTA.COD_BAND *= FMPERSON.COD_PERS AND "
'        .CommandText = .CommandText & "FMMOVCTA.COD_PART = FMOPERAC.COD_PART AND FMMOVCTA.COD_PART  = FMPARTIC.COD_PART AND "
'        .CommandText = .CommandText & "FMOPERAC.COD_PART = FMPARTIC.COD_PART AND "
'        .CommandText = .CommandText & "(TIP_PAGO='C' OR TIP_PAGO='T') AND (FMMOVCTA.TIP_OPER='SD' OR FMMOVCTA.TIP_OPER='SC') AND "
'        .CommandText = .CommandText & "(FMMOVCTA.FLG_CONF='' OR FMMOVCTA.FLG_CONF=NULL) AND "
'        .CommandText = .CommandText & "FMMOVCTA.COD_FOND='" & strFondo & "' AND "
'        .CommandText = .CommandText & "FCH_FRET='" & strFeccie & "'"
'        Set adoRegistro = .Execute
'        If Not adoRegistro.EOF Then
'            Exit Function
'        End If
'        adoRegistro.Close: Set adoRegistro = Nothing
    End With

    VerificarOperacionRetencion = True
    
End Function

Public Function VerificarOrdenInversion(strFondo As String, strFecha As String, strfechaMas1Dia As String) As Boolean

    Dim adoRegistro As ADODB.Recordset
    
    VerificarOrdenInversion = False
    
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
'        .CommandText = "SELECT COD_FOND FROM FMORDTIT WHERE FCH_ORDE='" & strFeccie & "' AND STA_ORDE='V' AND COD_FOND='" & strCodFondo & "'"
'        Set adoRegistro = .Execute
'        If Not adoRegistro.EOF Then
'            adoRegistro.Close: Set adoRegistro = Nothing
'            Exit Function
'        End If
'        adoRegistro.Close: Set adoRegistro = Nothing
    End With

    VerificarOrdenInversion = True
    
End Function

Public Function VerificarPeriodoContable(strFondo As String, strFecha As String, strfechaMas1Dia As String) As Boolean

    Dim adoRegistro As ADODB.Recordset
    
    VerificarPeriodoContable = False
    
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
'        .CommandText = "SELECT FCH_FINA FROM FMPRDCON WHERE COD_FOND='" & strCodFondo & "' AND MES_CONT='99' AND FLG_CIER='X'"
'        Set adoRegistro = .Execute
'        If Not adoRegistro.EOF Then
'            If adoRegistro("FCH_FINA") = strFeccie Then
'                adoRegistro.Close
'                .CommandText = "SELECT FCH_FINA FROM FMPRDCON WHERE COD_FOND='" & strCodFondo & "' AND MES_CONT='00' AND FCH_FINA='" & strFechaSiguiente & "'"
'                Set adoRegistro = .Execute
'                If adoRegistro.EOF Then
'                    MsgBox "Por favor genere el nuevo periodo contable.", vbOKOnly + vbInformation, "Cierre Diario"
'                    adoRegistro.Close: Set adoRegistro = Nothing
'                    Exit Function
'                End If
'            End If
'        End If
'        adoRegistro.Close: Set adoRegistro = Nothing
    End With

    VerificarPeriodoContable = True
    
End Function

