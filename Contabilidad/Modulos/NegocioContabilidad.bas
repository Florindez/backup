Attribute VB_Name = "NegocioContabilidad"
Public gstrCodCuentaBusquedaPlanContable        As String
Public gstrDescripCuentaBusquedaPlanContable    As String

Option Explicit

Public Sub GenerarAsientoImpuesto(ByVal strpCodFile As String, ByVal strpCodAnalitica As String, ByVal strpCodFondo As String, ByVal strpCodAdministradora, ByVal strpCodDetalleFile As String, ByVal strpCodDinamica As String, ByVal curpMontoAsiento As Currency, ByVal curpMontoNoGravado As Currency, ByVal dblpTipoCambio As Double, ByVal strpCodMoneda As String, ByVal strpDescripAsiento As String, ByVal strpCodModulo As String, Optional ByVal strpCodTipoComprobante As String, Optional ByVal strpNumComprobante As String, Optional ByVal strpTipoAuxiliar As String, Optional ByVal strpCodAuxiliar As String)

    Dim adoRegistro                 As ADODB.Recordset
    Dim curMontoMovimientoMN        As Currency, curMontoMovimientoME       As Currency
    Dim curMontoContable            As Currency
    Dim strIndDebeHaber             As String, strDescripMovimiento         As String
    Dim strNumAsiento               As String, strFechaGrabar               As String
    Dim strFechaCierre              As String, strFechaSiguiente            As String
    Dim intCantRegistros            As Integer
    
    With adoComm
        strNumAsiento = ObtenerSecuencialInversionOperacion(strpCodFondo, Valor_NumComprobante)
        strFechaCierre = Convertyyyymmdd(gdatFechaActual)
        strFechaSiguiente = Convertyyyymmdd(DateAdd("d", 1, gdatFechaActual))
        strFechaGrabar = strFechaCierre & Space(1) & Format(Time, "hh:mm")
                
        If IsNull(strpCodTipoComprobante) Then strpCodTipoComprobante = ""
        If IsNull(strpNumComprobante) Then strpNumComprobante = ""
                
        If IsNull(strpTipoAuxiliar) Then strpTipoAuxiliar = ""
        If IsNull(strpCodAuxiliar) Then strpCodAuxiliar = ""
                
        .CommandText = "SELECT COUNT(*) NumRegistros FROM DinamicaContable " & _
            "WHERE TipoOperacion='" & strpCodDinamica & "' AND CodFile='" & strpCodFile & "' AND (CodDetalleFile='" & _
            strpCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & strpCodAdministradora & "' AND CodMoneda = '" & _
            IIf(strpCodMoneda = Codigo_Moneda_Local, Codigo_Moneda_Local, Codigo_Moneda_Extranjero) & "'"
        Set adoRegistro = .Execute
    
        If Not adoRegistro.EOF Then
            If CInt(adoRegistro("NumRegistros")) > 0 Then
                intCantRegistros = CInt(adoRegistro("NumRegistros"))
            End If
        End If
        adoRegistro.Close
        
        '*** Cabecera ***
        .CommandText = "{ call up_ACAdicAsientoContable('" & _
            strpCodFondo & "','" & strpCodAdministradora & "','" & strNumAsiento & "','" & _
            strFechaGrabar & "','" & _
            gstrPeriodoActual & "','" & gstrMesActual & "','','" & _
            "FACTURACIÓN - " & strpDescripAsiento & "','" & strpCodMoneda & "','" & _
            Codigo_Moneda_Local & "','" & strpCodTipoComprobante & "','" & strpNumComprobante & "'," & _
            (CDec(curpMontoAsiento * (1 + gdblTasaIgv)) + curpMontoNoGravado) & ",'" & Estado_Activo & "'," & _
            intCantRegistros & ",'" & _
            strFechaGrabar & "','" & _
            strpCodModulo & "',''," & _
            CDec(dblpTipoCambio) & ",'','','" & _
            strpDescripAsiento & "','','X','') }"
        adoConn.Execute .CommandText
        
        '*** Detalle Contable ***
        .CommandText = "SELECT NumSecuencial,TipoCuentaInversion,IndDebeHaber,DescripDinamica,CodCuenta FROM DinamicaContable " & _
            "WHERE TipoOperacion='" & strpCodDinamica & "' AND CodFile='" & strpCodFile & "' AND (CodDetalleFile='" & _
            strpCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "' AND CodMoneda = '" & _
            IIf(strpCodMoneda = Codigo_Moneda_Local, Codigo_Moneda_Local, Codigo_Moneda_Extranjero) & "' ORDER BY NumSecuencial"
            
        Set adoRegistro = .Execute
    
        Do While Not adoRegistro.EOF
            Select Case Trim(adoRegistro("TipoCuentaInversion"))
            
                Case Codigo_CtaProvGasto
                    curMontoMovimientoMN = (curpMontoAsiento + curpMontoNoGravado)
                
                Case Codigo_CtaImpuesto
                    curMontoMovimientoMN = Round(curpMontoAsiento * (gdblTasaIgv), 2)
                
                Case Codigo_CtaXPagarEmitida
                    curMontoMovimientoMN = Round(curpMontoAsiento * (1 + gdblTasaIgv), 2) + curpMontoNoGravado
            
                Case Codigo_CtaInversion
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaProvInteres
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaInteres
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaCosto
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaIngresoOperacional
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaInteresVencido
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaVacCorrido
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaXPagar
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaXCobrar
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaInteresCorrido
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaProvReajusteK
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaReajusteK
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaProvFlucMercado
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaFlucMercado
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaProvInteresVac
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaInteresVac
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaIntCorridoK
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaProvFlucK
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaFlucK
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaInversionTransito
                    curMontoMovimientoMN = curpMontoAsiento
                
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
            curMontoContable = Round(curMontoMovimientoMN, 2)
    
            If strpCodMoneda <> Codigo_Moneda_Local Then
                curMontoContable = Round(curMontoMovimientoMN * dblpTipoCambio, 2)
                curMontoMovimientoME = curMontoMovimientoMN
                curMontoMovimientoMN = 0
            End If
                        
            '*** Movimiento ***
            .CommandText = "{ call up_ACAdicAsientoContableDetalle('" & _
                strNumAsiento & "','" & strpCodFondo & "','" & _
                gstrCodAdministradora & "'," & _
                CInt(adoRegistro("NumSecuencial")) & ",'" & _
                strFechaGrabar & "','" & _
                gstrPeriodoActual & "','" & _
                gstrMesActual & "','" & _
                strDescripMovimiento & "','" & _
                strIndDebeHaber & "','" & _
                Trim(adoRegistro("CodCuenta")) & "','" & _
                strpCodMoneda & "'," & _
                CDec(curMontoMovimientoMN) & "," & _
                CDec(curMontoMovimientoME) & "," & _
                CDec(curMontoContable) & ",'" & _
                strpCodFile & "','" & _
                strpCodAnalitica & "','" & _
                strpTipoAuxiliar & "','" & _
                strpCodAuxiliar & "') }"
            adoConn.Execute .CommandText
        
            '*** Saldos ***
            .CommandText = "{ call up_ACGenPartidaContableSaldos('" & _
                strpCodFondo & "','" & gstrCodAdministradora & "','" & _
                gstrPeriodoActual & "','" & gstrMesActual & "','" & _
                Trim(adoRegistro("CodCuenta")) & "','" & _
                strpCodFile & "','" & _
                strpCodAnalitica & "','" & _
                strFechaCierre & "','" & _
                strFechaSiguiente & "'," & _
                CDec(curMontoMovimientoMN) & "," & _
                CDec(curMontoMovimientoME) & "," & _
                CDec(curMontoContable) & ",'" & _
                strIndDebeHaber & "','" & _
                strpCodMoneda & "') }"
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
                
        '-- Verifica y ajusta posibles descuadres
        .CommandText = "{ call up_ACProcAsientoContableAjuste('" & _
                strpCodFondo & "','" & _
                strpCodAdministradora & "','" & _
                strNumAsiento & "') }"
        adoConn.Execute .CommandText
        
        '*** Actualizar el número del parámetro **
        adoComm.CommandText = "{ call up_ACActUltNumero('" & _
                    strpCodFondo & "','" & _
                    strpCodAdministradora & "','" & _
                    Valor_NumComprobante & "','" & _
                    strNumAsiento & "') }"
        adoConn.Execute .CommandText
    End With
    
End Sub
Public Sub GenerarAsientoGasto(ByVal strpCodFile As String, ByVal strpCodAnalitica As String, ByVal strpCodFondo As String, ByVal strpCodAdministradora, ByVal strpCodDetalleFile As String, ByVal strpCodDinamica As String, ByVal curpMontoAsiento As Currency, ByVal curpMontoNoGravado As Currency, ByVal dblpTipoCambio As Double, ByVal strpCodMoneda As String, ByVal strpDescripAsiento As String, ByVal strpCodModulo As String, Optional ByVal strpCodTipoComprobante As String, Optional ByVal strpNumComprobante As String, Optional ByVal strpCodAfectacion As String, Optional ByVal strpTipoPersona As String, Optional ByVal strpCodPersona As String)

    Dim adoRegistro                 As ADODB.Recordset
    Dim curMontoMovimientoMN        As Currency, curMontoMovimientoME       As Currency
    Dim curMontoContable            As Currency
    Dim strIndDebeHaber             As String, strDescripMovimiento         As String
    Dim strNumAsiento               As String, strFechaGrabar               As String
    Dim strFechaCierre              As String, strFechaSiguiente            As String
    Dim intCantRegistros            As Integer, curMontoAsiento             As Currency
    
    Dim dblValorTipoCambio          As Double, strTipoDocumento             As String
    Dim strNumDocumento             As String, strTipoPersonaContraparte    As String
    Dim strCodPersonaContraparte    As String, strIndContracuenta           As String
    Dim strCodContracuenta          As String, strCodFileContracuenta        As String
    Dim strCodAnaliticaContracuenta As String, strIndUltimoMovimiento       As String
    
    With adoComm
        strNumAsiento = ObtenerSecuencialInversionOperacion(strpCodFondo, Valor_NumComprobante)
        strFechaCierre = Convertyyyymmdd(gdatFechaActual)
        strFechaSiguiente = Convertyyyymmdd(DateAdd("d", 1, gdatFechaActual))
        strFechaGrabar = strFechaCierre & Space(1) & Format(Time, "hh:mm")
                
        If IsNull(strpCodTipoComprobante) Then strpCodTipoComprobante = ""
        If IsNull(strpNumComprobante) Then strpNumComprobante = ""
                
        If IsNull(strpTipoPersona) Then strpTipoPersona = ""
        If IsNull(strpCodPersona) Then strpCodPersona = ""
                
        If strpCodDinamica = Codigo_Dinamica_Gasto Then 'no emitida
            curMontoAsiento = curpMontoAsiento + curpMontoNoGravado
        End If
        
        If strpCodDinamica = Codigo_Dinamica_Gasto_Emitida Then 'emitida
            If strpCodAfectacion = Codigo_Afecto Then
                curMontoAsiento = Round((curpMontoAsiento * (1 + gdblTasaIgv)) + curpMontoNoGravado, 2)
            Else
                curMontoAsiento = curpMontoAsiento + curpMontoNoGravado
            End If
        End If
        
                
        .CommandText = "SELECT COUNT(*) NumRegistros FROM DinamicaContable " & _
            "WHERE TipoOperacion='" & strpCodDinamica & "' AND CodFile='" & strpCodFile & "' AND (CodDetalleFile='" & _
            strpCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & strpCodAdministradora & "' AND CodMoneda = '" & _
            IIf(strpCodMoneda = Codigo_Moneda_Local, Codigo_Moneda_Local, Codigo_Moneda_Extranjero) & "'"
        Set adoRegistro = .Execute
    
        If Not adoRegistro.EOF Then
            If CInt(adoRegistro("NumRegistros")) > 0 Then
                intCantRegistros = CInt(adoRegistro("NumRegistros"))
            End If
        End If
        adoRegistro.Close
        
        '*** Cabecera ***
        .CommandText = "{ call up_ACAdicAsientoContable('" & _
            strpCodFondo & "','" & strpCodAdministradora & "','" & strNumAsiento & "','" & _
            strFechaGrabar & "','" & _
            gstrPeriodoActual & "','" & gstrMesActual & "','','" & _
            "FACTURACIÓN - " & strpDescripAsiento & "','" & strpCodMoneda & "','" & _
            Codigo_Moneda_Local & "','" & strpCodTipoComprobante & "','" & strpNumComprobante & "'," & _
            CDec(curMontoAsiento) & ",'" & Estado_Activo & "'," & _
            intCantRegistros & ",'" & _
            strFechaGrabar & "','" & _
            strpCodModulo & "',''," & _
            CDec(dblpTipoCambio) & ",'','','" & _
            strpDescripAsiento & "','','X','') }"
        adoConn.Execute .CommandText
        
        '*** Detalle Contable ***
        .CommandText = "SELECT NumSecuencial,TipoCuentaInversion,IndDebeHaber,DescripDinamica,CodCuenta FROM DinamicaContable " & _
            "WHERE TipoOperacion='" & strpCodDinamica & "' AND CodFile='" & strpCodFile & "' AND (CodDetalleFile='" & _
            strpCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "' AND CodMoneda = '" & _
            IIf(strpCodMoneda = Codigo_Moneda_Local, Codigo_Moneda_Local, Codigo_Moneda_Extranjero) & "' ORDER BY NumSecuencial"
            
        Set adoRegistro = .Execute
    
        Do While Not adoRegistro.EOF
            
            curMontoMovimientoMN = 0
            
            Select Case Trim(adoRegistro("TipoCuentaInversion"))
            
                Case Codigo_CtaProvGasto
                    curMontoMovimientoMN = (curpMontoAsiento + curpMontoNoGravado)
                
                Case Codigo_CtaImpuesto
                    If strpCodDinamica = Codigo_Dinamica_Gasto_Emitida Then 'emitida
                        If strpCodAfectacion = Codigo_Afecto Then
                            curMontoMovimientoMN = Round(curpMontoAsiento * (gdblTasaIgv), 2)
                        Else
                            curMontoMovimientoMN = 0
                        End If
                    Else
                        curMontoMovimientoMN = 0
                    End If
                        
                Case Codigo_CtaXPagarEmitida
                    If strpCodDinamica = Codigo_Dinamica_Gasto_Emitida Then 'emitida
                        If strpCodAfectacion = Codigo_Afecto Then
                            curMontoMovimientoMN = Round(curpMontoAsiento * (1 + gdblTasaIgv), 2) + curpMontoNoGravado
                        Else
                            curMontoMovimientoMN = curpMontoAsiento + curpMontoNoGravado
                        End If
                    Else
                        curMontoMovimientoMN = 0
                    End If
            
                Case Codigo_CtaInversion
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaProvInteres
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaInteres
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaCosto
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaIngresoOperacional
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaInteresVencido
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaVacCorrido
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaXPagar
                    curMontoMovimientoMN = curpMontoAsiento + curpMontoNoGravado
                    
                Case Codigo_CtaXCobrar
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaInteresCorrido
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaProvReajusteK
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaReajusteK
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaProvFlucMercado
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaFlucMercado
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaProvInteresVac
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaInteresVac
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaIntCorridoK
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaProvFlucK
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaFlucK
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaInversionTransito
                    curMontoMovimientoMN = curpMontoAsiento
                
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
            curMontoContable = Round(curMontoMovimientoMN, 2)
    
            If strpCodMoneda <> Codigo_Moneda_Local Then
                curMontoContable = Round(curMontoMovimientoMN * dblpTipoCambio, 2)
                curMontoMovimientoME = curMontoMovimientoMN
                curMontoMovimientoMN = 0
                dblValorTipoCambio = dblpTipoCambio
            Else
                dblValorTipoCambio = 1
            End If
            
            If curMontoContable <> 0 Then
                        
                strTipoDocumento = strpCodTipoComprobante
                strNumDocumento = strpNumComprobante
                
                strTipoPersonaContraparte = Valor_Caracter
                strCodPersonaContraparte = Valor_Caracter
                
                strIndContracuenta = Valor_Caracter
                strCodContracuenta = Valor_Caracter
                strCodFileContracuenta = Valor_Caracter
                strCodAnaliticaContracuenta = Valor_Caracter
                strIndUltimoMovimiento = Valor_Caracter
                        
                '*** Movimiento ***
                .CommandText = "{ call up_ACAdicAsientoContableDetalle('" & strNumAsiento & "','" & strpCodFondo & "','" & _
                    gstrCodAdministradora & "'," & _
                    CInt(adoRegistro("NumSecuencial")) & ",'" & _
                    strFechaGrabar & "','" & _
                    gstrPeriodoActual & "','" & _
                    gstrMesActual & "','" & _
                    strDescripMovimiento & "','" & _
                    strIndDebeHaber & "','" & _
                    Trim(adoRegistro("CodCuenta")) & "','" & _
                    strpCodMoneda & "'," & _
                    CDec(curMontoMovimientoMN) & "," & _
                    CDec(curMontoMovimientoME) & "," & _
                    CDec(curMontoContable) & "," & _
                    dblValorTipoCambio & ",'" & _
                    strpCodFile & "','" & _
                    strpCodAnalitica & "','" & _
                    strTipoDocumento & "','" & _
                    strNumDocumento & "','" & _
                    strTipoPersonaContraparte & "','" & _
                    strCodPersonaContraparte & "','" & _
                    strIndContracuenta & "','" & _
                    strCodContracuenta & "','" & _
                    strCodFileContracuenta & "','" & _
                    strCodAnaliticaContracuenta & "','" & _
                    strIndUltimoMovimiento & "') }"
                adoConn.Execute .CommandText
                                
                                
                '*** Validar valor de cuenta contable ***
                If Trim(adoRegistro("CodCuenta")) = Valor_Caracter Then
                    MsgBox "Registro Nro. " & CStr(adoRegistro("NumSecuencial")) & " de Asiento Contable no tiene cuenta asignada", vbCritical, "Asiento Impuesto"
                    gblnRollBack = True
                    Exit Sub
                End If
            
            End If
            
            adoRegistro.MoveNext
        Loop
        adoRegistro.Close: Set adoRegistro = Nothing
                
        '-- Verifica y ajusta posibles descuadres
        .CommandText = "{ call up_ACProcAsientoContableAjuste('" & _
                strpCodFondo & "','" & _
                strpCodAdministradora & "','" & _
                strNumAsiento & "') }"
        adoConn.Execute .CommandText
        
        '*** Actualizar el número del parámetro **
        adoComm.CommandText = "{ call up_ACActUltNumero('" & _
                    strpCodFondo & "','" & _
                    strpCodAdministradora & "','" & _
                    Valor_NumComprobante & "','" & _
                    strNumAsiento & "') }"
        adoConn.Execute .CommandText
    End With
    
End Sub


Public Sub GenerarAsientoMismoDia(ByVal strpCodFile As String, ByVal strpCodAnalitica As String, ByVal strpCodFondo As String, ByVal strpCodAdministradora, ByVal strpCodDetalleFile As String, ByVal strpCodDinamica As String, ByVal curpMontoAsiento As Currency, ByVal dblpTipoCambio As Double, ByVal strpCodMoneda As String, ByVal strpDescripAsiento As String, ByVal strpCodModulo As String)

    Dim adoRegistro                 As ADODB.Recordset
    Dim curMontoMovimientoMN        As Currency, curMontoMovimientoME       As Currency
    Dim curMontoContable            As Currency
    Dim strIndDebeHaber             As String, strDescripMovimiento         As String
    Dim strNumAsiento               As String, strFechaGrabar               As String
    Dim strFechaCierre              As String, strFechaSiguiente            As String
    Dim intCantRegistros            As Integer
    
    With adoComm
        strNumAsiento = ObtenerSecuencialInversionOperacion(strpCodFondo, Valor_NumComprobante)
        strFechaCierre = Convertyyyymmdd(gdatFechaActual)
        strFechaSiguiente = Convertyyyymmdd(DateAdd("d", 1, gdatFechaActual))
        strFechaGrabar = strFechaCierre & Space(1) & Format(Time, "hh:mm")
                
        .CommandText = "SELECT COUNT(*) NumRegistros FROM DinamicaContable " & _
            "WHERE TipoOperacion='" & strpCodDinamica & "' AND CodFile='" & strpCodFile & "' AND (CodDetalleFile='" & _
            strpCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & strpCodAdministradora & "' AND CodMoneda = '" & _
            IIf(strpCodMoneda = Codigo_Moneda_Local, Codigo_Moneda_Local, Codigo_Moneda_Extranjero) & "'"
        Set adoRegistro = .Execute
    
        If Not adoRegistro.EOF Then
            If CInt(adoRegistro("NumRegistros")) > 0 Then
                intCantRegistros = CInt(adoRegistro("NumRegistros"))
            End If
        End If
        adoRegistro.Close
        
        '*** Cabecera ***
        .CommandText = "{ call up_ACAdicAsientoContable('" & _
            strpCodFondo & "','" & strpCodAdministradora & "','" & strNumAsiento & "','" & _
            strFechaGrabar & "','" & _
            gstrPeriodoActual & "','" & gstrMesActual & "','','" & _
            strpDescripAsiento & "','" & strpCodMoneda & "','" & _
            Codigo_Moneda_Local & "','',''," & _
            CDec(curpMontoAsiento) & ",'" & Estado_Activo & "'," & _
            intCantRegistros & ",'" & _
            strFechaGrabar & "','" & _
            strpCodModulo & "',''," & _
            CDec(dblpTipoCambio) & ",'','','" & _
            strpDescripAsiento & "','','X','') }"
        adoConn.Execute .CommandText
        
        '*** Detalle Contable ***
        .CommandText = "SELECT NumSecuencial,TipoCuentaInversion,IndDebeHaber,DescripDinamica,CodCuenta FROM DinamicaContable " & _
            "WHERE TipoOperacion='" & strpCodDinamica & "' AND CodFile='" & strpCodFile & "' AND (CodDetalleFile='" & _
            strpCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "' AND CodMoneda = '" & _
            IIf(strpCodMoneda = Codigo_Moneda_Local, Codigo_Moneda_Local, Codigo_Moneda_Extranjero) & "' ORDER BY NumSecuencial"
        Set adoRegistro = .Execute
    
        Do While Not adoRegistro.EOF
            Select Case Trim(adoRegistro("TipoCuentaInversion"))
                Case Codigo_CtaInversion
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaProvInteres
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaInteres
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaCosto
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaIngresoOperacional
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaInteresVencido
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaVacCorrido
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaXPagar
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaXCobrar
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaInteresCorrido
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaProvReajusteK
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaReajusteK
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaProvFlucMercado
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaFlucMercado
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaProvInteresVac
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaInteresVac
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaIntCorridoK
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaProvFlucK
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaFlucK
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaInversionTransito
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaProvGasto
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaImpuesto
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaCostoGastosBancarios  'ACR: 09/09/08
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaCostoComisionEspecial 'ACR: 09/09/08
                    curMontoMovimientoMN = curpMontoAsiento
                    
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
                        
            '*** Movimiento ***
            .CommandText = "{ call up_ACAdicAsientoContableDetalle('" & _
                strNumAsiento & "','" & strpCodFondo & "','" & _
                gstrCodAdministradora & "'," & _
                CInt(adoRegistro("NumSecuencial")) & ",'" & _
                strFechaGrabar & "','" & _
                gstrPeriodoActual & "','" & _
                gstrMesActual & "','" & _
                strDescripMovimiento & "','" & _
                strIndDebeHaber & "','" & _
                Trim(adoRegistro("CodCuenta")) & "','" & _
                strpCodMoneda & "'," & _
                CDec(curMontoMovimientoMN) & "," & _
                CDec(curMontoMovimientoME) & "," & _
                CDec(curMontoContable) & ",'" & _
                strpCodFile & "','" & _
                strpCodAnalitica & "') }"
            adoConn.Execute .CommandText
        
            '*** Saldos ***
            .CommandText = "{ call up_ACGenPartidaContableSaldos('" & _
                strpCodFondo & "','" & gstrCodAdministradora & "','" & _
                gstrPeriodoActual & "','" & gstrMesActual & "','" & _
                Trim(adoRegistro("CodCuenta")) & "','" & _
                strpCodFile & "','" & _
                strpCodAnalitica & "','" & _
                strFechaCierre & "','" & _
                strFechaSiguiente & "'," & _
                CDec(curMontoMovimientoMN) & "," & _
                CDec(curMontoMovimientoME) & "," & _
                CDec(curMontoContable) & ",'" & _
                strIndDebeHaber & "','" & _
                strpCodMoneda & "') }"
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
                
        '-- Verifica y ajusta posibles descuadres
        .CommandText = "{ call up_ACProcAsientoContableAjuste('" & _
                strpCodFondo & "','" & _
                strpCodAdministradora & "','" & _
                strNumAsiento & "') }"
        adoConn.Execute .CommandText
        
        '*** Actualizar el número del parámetro **
        adoComm.CommandText = "{ call up_ACActUltNumero('" & _
                    strpCodFondo & "','" & _
                    strpCodAdministradora & "','" & _
                    Valor_NumComprobante & "','" & _
                    strNumAsiento & "') }"
        adoConn.Execute .CommandText
        
        
        
        
    End With
    
End Sub
Public Sub GenerarAsientoContable(ByVal strpCodFile As String, ByVal strpCodAnalitica As String, ByVal strpCodFondo As String, ByVal strpCodAdministradora, ByVal strpCodDetalleFile As String, ByVal strpCodDinamica As String, ByVal curpMontoAsiento As Currency, ByVal dblpTipoCambio As Double, ByVal strpCodMoneda As String, ByVal strpDescripAsiento As String, ByVal strpCodModulo As String, Optional ByVal strpCodTipoComprobante As String, Optional ByVal strpNumComprobante As String, Optional ByVal strpTipoAuxiliar As String, Optional ByVal strpCodAuxiliar As String)

    Dim adoRegistro                 As ADODB.Recordset
    Dim curMontoMovimientoMN        As Currency, curMontoMovimientoME       As Currency
    Dim curMontoContable            As Currency
    Dim strIndDebeHaber             As String, strDescripMovimiento         As String
    Dim strNumAsiento               As String, strFechaGrabar               As String
    Dim strFechaCierre              As String, strFechaSiguiente            As String
    Dim intCantRegistros            As Integer, strCodMonedaRegistro        As String
    
    With adoComm
        strNumAsiento = ObtenerSecuencialInversionOperacion(strpCodFondo, Valor_NumComprobante)
        strFechaCierre = Convertyyyymmdd(gdatFechaActual)
        strFechaSiguiente = Convertyyyymmdd(DateAdd("d", 1, gdatFechaActual))
        strFechaGrabar = strFechaCierre & Space(1) & Format(Time, "hh:mm")
                
        If IsNull(strpCodTipoComprobante) Then strpCodTipoComprobante = ""
        If IsNull(strpNumComprobante) Then strpNumComprobante = ""
        
        If IsNull(strpTipoAuxiliar) Then strpTipoAuxiliar = ""
        If IsNull(strpCodAuxiliar) Then strpCodAuxiliar = ""
                
        .CommandText = "SELECT COUNT(*) NumRegistros FROM DinamicaContable " & _
            "WHERE TipoOperacion='" & strpCodDinamica & "' AND CodFile='" & strpCodFile & "' AND (CodDetalleFile='" & _
            strpCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & strpCodAdministradora & "' AND CodMoneda = '" & _
            IIf(strpCodMoneda = Codigo_Moneda_Local, Codigo_Moneda_Local, Codigo_Moneda_Extranjero) & "'"
        Set adoRegistro = .Execute
    
        If Not adoRegistro.EOF Then
            If CInt(adoRegistro("NumRegistros")) > 0 Then
                intCantRegistros = CInt(adoRegistro("NumRegistros"))
            End If
        End If
        adoRegistro.Close
        
        '*** Cabecera ***
        .CommandText = "{ call up_ACAdicAsientoContable('" & _
            strpCodFondo & "','" & strpCodAdministradora & "','" & strNumAsiento & "','" & _
            strFechaGrabar & "','" & _
            gstrPeriodoActual & "','" & gstrMesActual & "','','" & _
            strpDescripAsiento & "','" & strpCodMoneda & "','" & _
            Codigo_Moneda_Local & "','" & strpCodTipoComprobante & "','" & strpNumComprobante & "'," & _
            CDec(curpMontoAsiento) & ",'" & Estado_Activo & "'," & _
            intCantRegistros & ",'" & _
            strFechaGrabar & "','" & _
            strpCodModulo & "',''," & _
            CDec(dblpTipoCambio) & ",'','','" & _
            strpDescripAsiento & "','','X','') }"
        adoConn.Execute .CommandText
        
        '*** Detalle Contable ***
        .CommandText = "SELECT NumSecuencial,TipoCuentaInversion,IndDebeHaber,DescripDinamica,CodCuenta FROM DinamicaContable " & _
            "WHERE TipoOperacion='" & strpCodDinamica & "' AND CodFile='" & strpCodFile & "' AND (CodDetalleFile='" & _
            strpCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "' AND CodMoneda = '" & _
            IIf(strpCodMoneda = Codigo_Moneda_Local, Codigo_Moneda_Local, Codigo_Moneda_Extranjero) & "' ORDER BY NumSecuencial"
        Set adoRegistro = .Execute
    
        Do While Not adoRegistro.EOF
        
            strCodMonedaRegistro = strpCodMoneda
            
            Select Case Trim(adoRegistro("TipoCuentaInversion"))
                Case Codigo_CtaInversion
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaProvInteres
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaInteres
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaCosto
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaIngresoOperacional
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaInteresVencido
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaVacCorrido
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaXPagar
                    curMontoMovimientoMN = curpMontoAsiento
                
                Case Codigo_CtaXPagarEmitida
                    curMontoMovimientoMN = curpMontoAsiento
                
                Case Codigo_CtaXCobrar
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaInteresCorrido
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaProvReajusteK
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaReajusteK
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaProvFlucMercado
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaFlucMercado
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaProvInteresVac
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaInteresVac
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaIntCorridoK
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaProvFlucK
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaFlucK
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaInversionTransito
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaProvGasto
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaImpuesto
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaCostoGastosBancarios     'ACR: 09/09/08
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaCostoComisionEspecial    'ACR: 09/09/08
                    curMontoMovimientoMN = curpMontoAsiento
                
                Case Codigo_CtaDetraccion               'ACR: 28/11/08
                    If strpCodMoneda <> Codigo_Moneda_Local Then
                        curMontoMovimientoMN = Round(curpMontoAsiento * dblpTipoCambio, 2)
                        strCodMonedaRegistro = Codigo_Moneda_Local
                    Else
                        curMontoMovimientoMN = curpMontoAsiento
                    End If
                    
                Case Codigo_CtaComision
                    curMontoMovimientoMN = curpMontoAsiento
                
                Case Codigo_CtaRetencion
                    If strpCodMoneda <> Codigo_Moneda_Local Then
                        curMontoMovimientoMN = Round(curpMontoAsiento * dblpTipoCambio, 2)
                        strCodMonedaRegistro = Codigo_Moneda_Local
                    Else
                        curMontoMovimientoMN = curpMontoAsiento
                    End If
                    
                Case Codigo_CtaIngresoOperacional_AjusteRedondeo
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_Perdida_AjusteRedondeo
                    curMontoMovimientoMN = curpMontoAsiento
                    
                    
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
    
            If strCodMonedaRegistro <> Codigo_Moneda_Local Then
                curMontoContable = Round(curMontoMovimientoMN * dblpTipoCambio, 2)
                curMontoMovimientoME = curMontoMovimientoMN
                curMontoMovimientoMN = 0
            End If
                        
            '*** Movimiento ***
            .CommandText = "{ call up_ACAdicAsientoContableDetalle('" & _
                strNumAsiento & "','" & strpCodFondo & "','" & _
                gstrCodAdministradora & "'," & _
                CInt(adoRegistro("NumSecuencial")) & ",'" & _
                strFechaGrabar & "','" & _
                gstrPeriodoActual & "','" & _
                gstrMesActual & "','" & _
                strDescripMovimiento & "','" & _
                strIndDebeHaber & "','" & _
                Trim(adoRegistro("CodCuenta")) & "','" & _
                strCodMonedaRegistro & "'," & _
                CDec(curMontoMovimientoMN) & "," & _
                CDec(curMontoMovimientoME) & "," & _
                CDec(curMontoContable) & ",'" & _
                strpCodFile & "','" & _
                strpCodAnalitica & "','" & _
                strpTipoAuxiliar & "','" & _
                strpCodAuxiliar & "') }"
            adoConn.Execute .CommandText
        
            '*** Saldos ***
            .CommandText = "{ call up_ACGenPartidaContableSaldos('" & _
                strpCodFondo & "','" & gstrCodAdministradora & "','" & _
                gstrPeriodoActual & "','" & gstrMesActual & "','" & _
                Trim(adoRegistro("CodCuenta")) & "','" & _
                strpCodFile & "','" & _
                strpCodAnalitica & "','" & _
                strFechaCierre & "','" & _
                strFechaSiguiente & "'," & _
                CDec(curMontoMovimientoMN) & "," & _
                CDec(curMontoMovimientoME) & "," & _
                CDec(curMontoContable) & ",'" & _
                strIndDebeHaber & "','" & _
                strCodMonedaRegistro & "') }"
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
        
        '*** Actualizar el número del parámetro **
        adoComm.CommandText = "{ call up_ACActUltNumero('" & _
                    strpCodFondo & "','" & _
                    strpCodAdministradora & "','" & _
                    Valor_NumComprobante & "','" & _
                    strNumAsiento & "') }"
        adoConn.Execute .CommandText
        
    End With
    
End Sub


Public Sub GenerarOrdenGastosFondo(ByVal strpCodDetalleGasto As String, ByVal strpCodGasto As String, ByVal strpCodFondo As String, ByVal intpNumGasto As Integer, ByVal strpCodProveedor As String, ByVal numpRegistro As Long)

    Dim adoRegistro             As ADODB.Recordset
    Dim adoConsulta             As ADODB.Recordset
    Dim adoAuxiliar             As ADODB.Recordset
    Dim strNumCaja              As String, strCodFile                   As String
    Dim strCodDetalleFile       As String, strDescripGasto              As String
    Dim strSQLOrdenCaja         As String, strSQLOrdenCajaDetalle       As String
    Dim strSQLOrdenCajaMN       As String, strSQLOrdenCajaDetalleMN     As String
    Dim strSQLOrdenCajaDetalleMN2   As String
    Dim strSQLOrdenCajaDetalleI As String, strSQLOrdenCajaDetalleMNI    As String
    Dim strIndDetraccion        As String, strIndImpuesto               As String
    Dim strIndRetencion         As String, strCodCreditoFiscal          As String
    Dim strCodAnalitica         As String, strCodMonedaGasto            As String
    Dim curSaldoProvision       As Currency, curValorImpuesto           As Currency
    Dim dblTipoCambioGasto              As Double
    Dim blnVenceGasto                   As Boolean
    Dim strTipoAuxiliar                 As String
    Dim strCodAuxiliar                  As String
    Dim dblMontoDetraccion              As Double
    Dim dblMontoDetraccionMN            As Double, dblMontoDetraccionME As Double
    Dim dblAjusteDetraccionMN      As Double
    

    frmMainMdi.stbMdi.Panels(3).Text = "Generando Orden de Pago de Gastos..."
    
    strTipoAuxiliar = "02" 'Gastos - Proveedores
    
    Set adoRegistro = New ADODB.Recordset
    With adoComm
        .CommandText = "SELECT * FROM FondoGasto " & _
            "WHERE CodCuenta='" & strpCodGasto & "' AND " & _
            "NumGasto=" & intpNumGasto & " AND CodFondo='" & strpCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND IndConfirma=''"
        Set adoRegistro = .Execute
        
        Do While Not adoRegistro.EOF
            blnVenceGasto = False
            strCodCreditoFiscal = Trim(adoRegistro("CodCreditoFiscal"))
        
            '*** Obtener Secuenciales ***
            strNumCaja = ObtenerSecuencialInversionOperacion(strpCodFondo, Valor_NumOrdenCaja)
           
            strCodAuxiliar = adoRegistro("TipoProveedor") & adoRegistro("CodProveedor")
           
            blnVenceGasto = True 'ACR
            
            '*** Si vence la provisión del Gasto ***
            If blnVenceGasto Then
                strCodFile = Trim(adoRegistro("CodFile"))
                strCodAnalitica = Trim(adoRegistro("CodAnalitica"))
                
                strCodDetalleFile = strpCodDetalleGasto
                
                strCodMonedaGasto = Trim(adoRegistro("CodMoneda"))
                        
                Set adoConsulta = New ADODB.Recordset
                
                '*** Obtener Descripción del Gasto ***
                .CommandText = "SELECT DescripCuenta FROM PlanContable WHERE CodCuenta='" & strpCodGasto & "'"
                Set adoConsulta = .Execute
                
                If Not adoConsulta.EOF Then
                    strDescripGasto = Trim(adoConsulta("DescripCuenta"))
                End If
                adoConsulta.Close
                            
                strDescripGasto = Trim(adoRegistro("DescripGasto"))
                
                '*** Obtener las cuentas de inversión ***
                Call ObtenerCuentasInversion(strCodFile, strCodDetalleFile, Trim((adoRegistro("CodMoneda"))))
                
                '*** Obtener Tipo de Cambio ***
                '*** Ojo corregir el tipo (Codigo_TipoCambio_SBS)por sunat
                                                                                                                                                
                .CommandText = "SELECT CodDetraccionSiNo,CodFormaPagoDetraccion,CodMonedaDetraccion,CodFileDetraccion,CodAnaliticaDetraccion,MontoDetraccion,TipoCambioPago,MontoPago,MontoTotal,FechaPago,Importe,ValorImpuesto,ValorNoGravado,ValorTotal,CodTipoComprobante,NumComprobante,FechaComprobante " & _
                    "FROM RegistroCompra WHERE NumGasto=" & CInt(adoRegistro("NumGasto")) & " AND CodFondo='" & strpCodFondo & "' AND " & _
                    "CodAdministradora='" & gstrCodAdministradora & "' and NumRegistro = " & numpRegistro
                Set adoConsulta = .Execute
                
                If Not adoConsulta.EOF Then
                    
                    '*** Obtener Tipo de Cambio ***
                    'Es tipo de cambio SUNAT de la fecha de emision del documento si el documento es factura
                    If strCodMonedaGasto <> Codigo_Moneda_Local Then
                        'If adoConsulta("CodTipoComprobante") = Codigo_Comprobante_Factura Then 'Factura
                            dblTipoCambioGasto = ObtenerTipoCambioMoneda(Codigo_TipoCambio_SBS, Codigo_Valor_TipoCambioVenta, adoConsulta("FechaComprobante"), Codigo_Moneda_Local, strCodMonedaGasto)
                        'Else
                            'dblTipoCambioGasto = ObtenerTipoCambioMoneda(Codigo_TipoCambio_Sunat, Codigo_Valor_TipoCambioVenta, adoConsulta("FechaComprobante"), Codigo_Moneda_Local, strCodMonedaGasto)
                            'dblTipoCambioGasto = ObtenerTipoCambioMoneda(Codigo_TipoCambio_Sunat, Codigo_Valor_TipoCambioVenta, gdatFechaActual, Codigo_Moneda_Local, strCodMonedaGasto)
                            If dblTipoCambioGasto = 0 Then
                                dblTipoCambioGasto = ObtenerTipoCambioMoneda(Codigo_TipoCambio_SBS, Codigo_Valor_TipoCambioVenta, DateAdd("d", -1, gdatFechaActual), Codigo_Moneda_Local, strCodMonedaGasto)
                            End If
                        'End If
                    Else
                        dblTipoCambioGasto = 1
                    End If
                    
                    If adoRegistro("CodAplicacionDevengo") = Codigo_Aplica_Devengo_Inmediata Then
                        'Actualiza la provision; es igual al gasto
                        .CommandText = "UPDATE FondoGasto SET MontoDevengo = MontoGasto " & _
                        "WHERE CodCuenta='" & strpCodGasto & "' AND " & _
                        "NumGasto=" & CInt(adoRegistro("NumGasto")) & " AND CodFondo='" & strpCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"

                        adoConn.Execute .CommandText
                    End If
                
                    Set adoAuxiliar = New ADODB.Recordset
                    
                    .CommandText = "SELECT IndImpuesto,IndRetencion " & _
                        "FROM TipoComprobantePago WHERE CodTipoComprobantePago='" & adoConsulta("CodTipoComprobante") & "'"
                    Set adoAuxiliar = .Execute
            
                    If Not adoAuxiliar.EOF Then
                        strIndImpuesto = Trim(adoAuxiliar("IndImpuesto"))
                        strIndRetencion = Trim(adoAuxiliar("IndRetencion"))
                    End If
                    adoAuxiliar.Close: Set adoAuxiliar = Nothing
                    
'                    Set adoAuxiliar = New ADODB.Recordset
'                        .CommandText = "SELECT * FROM FondoGastoPeriodo " & _
'                            "WHERE " & _
'                            "NumGasto=" & intpNumGasto & " AND CodFondo='" & strpCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
'                    Set adoAuxiliar = .Execute
                
                    dblMontoDetraccion = 0
                    dblMontoDetraccionMN = 0
                    
                    If adoConsulta("CodDetraccionSiNo") = Codigo_Respuesta_Si Then 'ESTE CAMPO INDICA SI TIENE DETRACCION O RETENCION PERO NO ESPECIFICA EXACTAMENTE CUAL DE LOS 2 ES....
                        
                        dblMontoDetraccion = CDec(adoConsulta("MontoDetraccion"))
                        
                        If strIndImpuesto = Valor_Indicador Then
                            strIndDetraccion = Valor_Indicador
                            
                            'If strCodMonedaGasto <> Codigo_Moneda_Local Then
                                dblMontoDetraccionMN = dblMontoDetraccion
                            'End If
                            
'                            Call GenerarAsientoContable(strCodFile, strCodAnalitica, strpCodFondo, gstrCodAdministradora, strCodDetalleFile, Codigo_Dinamica_Detraccion, dblMontoDetraccion, dblTipoCambioGasto, strCodMonedaGasto, strDescripGasto, frmMainMdi.Tag, Trim(adoConsulta("CodTipoComprobante")), Trim(adoConsulta("NumComprobante")), strTipoAuxiliar, strCodAuxiliar)
                            
                            'Control del redondeo y actualiza detraccion redondeada
                            'dblAjusteDetraccionMN = dblMontoDetraccionMN - Round(dblMontoDetraccionMN, 0)
                            dblAjusteDetraccionMN = Round(Round(dblMontoDetraccionMN) - dblMontoDetraccionMN, 2)
                            
'                            If dblAjusteDetraccionMN < 0 Then
'                                Call GenerarAsientoContable(strCodFile, strCodAnalitica, strpCodFondo, gstrCodAdministradora, strCodDetalleFile, Codigo_Dinamica_Detraccion_Ajuste_Redondeo_Ganancia, Abs(dblAjusteDetraccionMN), dblTipoCambioGasto, Codigo_Moneda_Local, "Ajuste por Redondeo " & strDescripGasto, frmMainMdi.Tag, Trim(adoConsulta("CodTipoComprobante")), Trim(adoConsulta("NumComprobante")), strTipoAuxiliar, strCodAuxiliar)
'                            ElseIf dblAjusteDetraccionMN > 0 Then
'                                Call GenerarAsientoContable(strCodFile, strCodAnalitica, strpCodFondo, gstrCodAdministradora, strCodDetalleFile, Codigo_Dinamica_Detraccion_Ajuste_Redondeo_Perdida, Abs(dblAjusteDetraccionMN), dblTipoCambioGasto, Codigo_Moneda_Local, "Ajuste por Redondeo " & strDescripGasto, frmMainMdi.Tag, Trim(adoConsulta("CodTipoComprobante")), Trim(adoConsulta("NumComprobante")), strTipoAuxiliar, strCodAuxiliar)
'                            End If
                        
                        ElseIf strIndRetencion = Valor_Indicador Then
                            Call GenerarAsientoContable(strCodFile, strCodAnalitica, strpCodFondo, gstrCodAdministradora, strCodDetalleFile, Codigo_Dinamica_Retencion, CDec(adoRegistro("MontoGasto") - adoConsulta("MontoPago")), dblTipoCambioGasto, strCodMonedaGasto, strDescripGasto, frmMainMdi.Tag, Trim(adoConsulta("CodTipoComprobante")), Trim(adoConsulta("NumComprobante")), strTipoAuxiliar, strCodAuxiliar)
                        End If
                        
                        'strTipoAuxiliar, strCodAuxiliar
                        'CodTipoComprobante,NumComprobante
                        
                        '*** Orden de Cobro/Pago ***
                        strSQLOrdenCaja = "{ call up_ACAdicMovimientoFondo('" & strpCodFondo & "','" & gstrCodAdministradora & "','" & _
                            strNumCaja & "','" & Convertyyyymmdd(gdatFechaActual) & "','" & Trim(frmMainMdi.Tag) & "','','" & Convertyyyymmdd(adoConsulta("FechaPago")) & "'," & _
                            "'','" & adoConsulta("CodTipoComprobante") & "','" & adoConsulta("NumComprobante") & "','','S','" & strCtaXPagarEmitida & "'," & CDec(adoConsulta("MontoPago")) * -1 & ",'" & _
                            strCodFile & "','" & strCodAnalitica & "','" & strCodMonedaGasto & "','" & _
                            strDescripGasto & "','" & Codigo_Caja_Gasto & "','','" & Estado_Caja_NoConfirmado & "','','','','','','','','','',0,'" & strpCodProveedor & "','" & Codigo_Tipo_Persona_Proveedor & "','" & gstrLogin & "') }"
                                        
                        '*** Orden de Cobro/Pago Detalle Impuesto ***
                        curValorImpuesto = Round(CCur(adoConsulta("ValorImpuesto") * (1 - gdblTasaDetraccion)), 2)
                                                
                        If strIndRetencion = Valor_Indicador Then
                                
                            '*** Orden de Cobro/Pago Detalle ***
                            strSQLOrdenCajaDetalle = "{ call up_ACAdicMovimientoFondoDetalle('" & strpCodFondo & "','" & gstrCodAdministradora & "','" & _
                                strNumCaja & "','" & Convertyyyymmdd(gdatFechaActual) & "',1,'" & Trim(frmMainMdi.Tag) & "','" & strDescripGasto & "','" & _
                                "D','" & strCtaXPagarEmitida & "'," & CDec(adoConsulta("MontoPago")) & ",'" & _
                                strCodFile & "','" & strCodAnalitica & "','" & strTipoAuxiliar & "','" & strCodAuxiliar & "','" & strCodMonedaGasto & "','') }"
                        Else
                            If strCodCreditoFiscal = Codigo_Tipo_Credito_RentaNoGravada Then
                                    
                                '*** Orden de Cobro/Pago Detalle ***
                                strSQLOrdenCajaDetalle = "{ call up_ACAdicMovimientoFondoDetalle('" & strpCodFondo & "','" & gstrCodAdministradora & "','" & _
                                    strNumCaja & "','" & Convertyyyymmdd(gdatFechaActual) & "',1,'" & Trim(frmMainMdi.Tag) & "','" & strDescripGasto & "','" & _
                                    "D','" & strCtaXPagarEmitida & "'," & CDec(adoConsulta("MontoPago")) & ",'" & _
                                    strCodFile & "','" & strCodAnalitica & "','" & strTipoAuxiliar & "','" & strCodAuxiliar & "','" & strCodMonedaGasto & "','') }"
                            Else
                                '*** Orden de Cobro/Pago Detalle ***
                                strSQLOrdenCajaDetalle = "{ call up_ACAdicMovimientoFondoDetalle('" & strpCodFondo & "','" & gstrCodAdministradora & "','" & _
                                    strNumCaja & "','" & Convertyyyymmdd(gdatFechaActual) & "',1,'" & Trim(frmMainMdi.Tag) & "','" & strDescripGasto & "','" & _
                                    "D','" & strCtaXPagarEmitida & "'," & CDec(adoConsulta("MontoPago")) & ",'" & _
                                    strCodFile & "','" & strCodAnalitica & "','" & strTipoAuxiliar & "','" & strCodAuxiliar & "','" & strCodMonedaGasto & "','') }"
                            End If
                        End If
                    Else
                        '*** Orden de Cobro/Pago ***
                        strSQLOrdenCaja = "{ call up_ACAdicMovimientoFondo('" & strpCodFondo & "','" & gstrCodAdministradora & "','" & _
                            strNumCaja & "','" & Convertyyyymmdd(gdatFechaActual) & "','" & Trim(frmMainMdi.Tag) & "','','" & Convertyyyymmdd(adoConsulta("FechaPago")) & "'," & _
                            "'','" & adoConsulta("CodTipoComprobante") & "','" & adoConsulta("NumComprobante") & "','','S','" & strCtaXPagarEmitida & "'," & CDec(adoConsulta("ValorTotal") * -1) & ",'" & _
                            strCodFile & "','" & strCodAnalitica & "','" & strCodMonedaGasto & "','" & _
                            strDescripGasto & "','" & Codigo_Caja_Gasto & "','','" & Estado_Caja_NoConfirmado & "','','','','','','','','','',0,'" & strpCodProveedor & "','" & Codigo_Tipo_Persona_Proveedor & "','" & gstrLogin & "') }"
                
'                        If strCodCreditoFiscal = Codigo_Tipo_Credito_RentaNoGravada Then
                            '*** Orden de Cobro/Pago Detalle ***
                            strSQLOrdenCajaDetalle = "{ call up_ACAdicMovimientoFondoDetalle('" & strpCodFondo & "','" & gstrCodAdministradora & "','" & _
                                strNumCaja & "','" & Convertyyyymmdd(gdatFechaActual) & "',1,'" & Trim(frmMainMdi.Tag) & "','" & strDescripGasto & "','" & _
                                "D','" & strCtaXPagarEmitida & "'," & CDec(adoConsulta("ValorTotal")) & ",'" & _
                                strCodFile & "','" & strCodAnalitica & "','" & strTipoAuxiliar & "','" & strCodAuxiliar & "','" & strCodMonedaGasto & "','') }"
                        
                    End If
                                  
                    'On Error GoTo Ctrl_Error
                
                    '*** Orden de Cobro ***
                    adoConn.Execute strSQLOrdenCaja
                    adoConn.Execute strSQLOrdenCajaDetalle
                    
                    'Actualiza orden recien creada con el codigo de gasto correspondiente 'ACR
                    .CommandText = "UPDATE MovimientoFondo SET NumGasto=" & adoRegistro("NumGasto") & " " & _
                    "WHERE NumOrdenCobroPago='" & strNumCaja & "' AND CodFondo='" & strpCodFondo & "' AND " & _
                    "CodAdministradora='" & gstrCodAdministradora & "'"
                    
                    adoConn.Execute .CommandText
                    
                    If adoConsulta("CodDetraccionSiNo") = Codigo_Respuesta_Si Then
                        
                        '*** Actualizar el número del parámetro **
                        .CommandText = "{ call up_ACActUltNumero('" & strpCodFondo & "','" & _
                            gstrCodAdministradora & "','" & Valor_NumOrdenCaja & "','" & strNumCaja & "') }"
                        adoConn.Execute .CommandText
                    
                        '*** Obtener Secuenciales ***
                        strNumCaja = ObtenerSecuencialInversionOperacion(strpCodFondo, Valor_NumOrdenCaja)
                                                
                        'cambia adoConsulta("MontoDetraccion") por dblMontoDetraccionMN
                                                
                        If strIndDetraccion = Valor_Indicador Then
                            '*** Orden de Cobro/Pago Detracción ***
                            strSQLOrdenCajaMN = "{ call up_ACAdicMovimientoFondo('" & strpCodFondo & "','" & gstrCodAdministradora & "','" & _
                                strNumCaja & "','" & Convertyyyymmdd(gdatFechaActual) & "','" & Trim(frmMainMdi.Tag) & "','','" & Convertyyyymmdd(adoConsulta("FechaPago")) & "'," & _
                                "'','" & adoConsulta("CodTipoComprobante") & "','" & adoConsulta("NumComprobante") & "','','S','" & strCtaDetraccion & "'," & CDec(Round(dblMontoDetraccionMN) * -1) & ",'" & _
                                strCodFile & "','" & strCodAnalitica & "','" & adoConsulta("CodMonedaDetraccion") & "','" & _
                                "DETRACCIÓN FAC." & Trim(adoConsulta("NumComprobante")) & " - " & strDescripGasto & "','" & Codigo_Caja_Gasto & "','','" & Estado_Caja_NoConfirmado & "','','','','','','','','',''," & adoRegistro("NumGasto") & ",'" & strpCodProveedor & "','" & Codigo_Tipo_Persona_Proveedor & "','" & gstrLogin & "') }"
                        ElseIf strIndRetencion = Valor_Indicador Then
                            strSQLOrdenCajaMN = "{ call up_ACAdicMovimientoFondo('" & strpCodFondo & "','" & gstrCodAdministradora & "','" & _
                                strNumCaja & "','" & Convertyyyymmdd(gdatFechaActual) & "','" & Trim(frmMainMdi.Tag) & "','','" & Convertyyyymmdd(adoConsulta("FechaPago")) & "'," & _
                                "'','" & adoConsulta("CodTipoComprobante") & "','" & adoConsulta("NumComprobante") & "','','S','" & strCtaRetencion & "'," & CDec(dblMontoDetraccionMN * -1) & ",'" & _
                                strCodFile & "','" & strCodAnalitica & "','" & adoConsulta("CodMonedaDetraccion") & "','" & _
                                strDescripGasto & "','" & Codigo_Caja_Gasto & "','','" & Estado_Caja_NoConfirmado & "','','','','','','','','','',0,'" & strpCodProveedor & "','" & Codigo_Tipo_Persona_Proveedor & "','" & gstrLogin & "') }"
                        End If
                                        
                        '*** Orden de Cobro/Pago Detalle Detracción 'Impuesto ***
                        dblMontoDetraccionME = CCur(adoConsulta("MontoTotal") - adoConsulta("MontoPago"))
          
                        If strIndDetraccion = Valor_Indicador Then
                                If CDec(dblAjusteDetraccionMN) = 0 Then
                                    strSQLOrdenCajaDetalleMN = "{ call up_ACAdicMovimientoFondoDetalle('" & strpCodFondo & "','" & gstrCodAdministradora & "','" & _
                                        strNumCaja & "','" & Convertyyyymmdd(gdatFechaActual) & "',1,'" & Trim(frmMainMdi.Tag) & "','" & "FACT. " & Trim(adoConsulta("NumComprobante")) & "','" & _
                                        "D','" & strCtaDetraccion & "'," & CDec(dblMontoDetraccionME) & ",'" & _
                                        strCodFile & "','" & strCodAnalitica & "','" & strTipoAuxiliar & "','" & strCodAuxiliar & "','" & Codigo_Moneda_Local & "','') }"
                                ElseIf CDec(dblAjusteDetraccionMN) > 0 Then
                                    strSQLOrdenCajaDetalleMN = "{ call up_ACAdicMovimientoFondoDetalle('" & strpCodFondo & "','" & gstrCodAdministradora & "','" & _
                                        strNumCaja & "','" & Convertyyyymmdd(gdatFechaActual) & "',1,'" & Trim(frmMainMdi.Tag) & "','" & "FACT. " & Trim(adoConsulta("NumComprobante")) & "','" & _
                                        "D','" & strCtaDetraccion & "'," & CDec(dblMontoDetraccionME) & ",'" & _
                                        strCodFile & "','" & strCodAnalitica & "','" & strTipoAuxiliar & "','" & strCodAuxiliar & "','" & strCodMonedaGasto & "','') }"
                                    '***Pérdida por Redondeo***
                                    strSQLOrdenCajaDetalleMN2 = "{ call up_ACAdicMovimientoFondoDetalle('" & strpCodFondo & "','" & gstrCodAdministradora & "','" & _
                                        strNumCaja & "','" & Convertyyyymmdd(gdatFechaActual) & "',2,'" & Trim(frmMainMdi.Tag) & "','" & "PÉRDIDA POR REDONDEO" & "','" & _
                                        "D','" & strCtaPerdidaRedondeo & "'," & CDec(dblAjusteDetraccionMN) & ",'" & _
                                        strCodFile & "','" & strCodAnalitica & "','" & strTipoAuxiliar & "','" & strCodAuxiliar & "','" & Codigo_Moneda_Local & "','') }"
                                ElseIf CDec(dblAjusteDetraccionMN) < 0 Then
                                    strSQLOrdenCajaDetalleMN = "{ call up_ACAdicMovimientoFondoDetalle('" & strpCodFondo & "','" & gstrCodAdministradora & "','" & _
                                        strNumCaja & "','" & Convertyyyymmdd(gdatFechaActual) & "',1,'" & Trim(frmMainMdi.Tag) & "','" & "FACT. " & Trim(adoConsulta("NumComprobante")) & "','" & _
                                        "D','" & strCtaDetraccion & "'," & CDec(dblMontoDetraccionME) & ",'" & _
                                        strCodFile & "','" & strCodAnalitica & "','" & strTipoAuxiliar & "','" & strCodAuxiliar & "','" & strCodMonedaGasto & "','') }"
                                    '***Ganancia por Redondeo***
                                    strSQLOrdenCajaDetalleMN2 = "{ call up_ACAdicMovimientoFondoDetalle('" & strpCodFondo & "','" & gstrCodAdministradora & "','" & _
                                        strNumCaja & "','" & Convertyyyymmdd(gdatFechaActual) & "',2,'" & Trim(frmMainMdi.Tag) & "','" & "GANANCIA POR REDONDEO" & "','" & _
                                        "H','" & strCtaGananciaRedondeo & "'," & CDec(dblAjusteDetraccionMN) & ",'" & _
                                        strCodFile & "','" & strCodAnalitica & "','" & strTipoAuxiliar & "','" & strCodAuxiliar & "','" & Codigo_Moneda_Local & "','') }"
                                End If
                        ElseIf strIndRetencion = Valor_Indicador Then
                            strSQLOrdenCajaDetalleMN = "{ call up_ACAdicMovimientoFondoDetalle('" & strpCodFondo & "','" & gstrCodAdministradora & "','" & _
                                strNumCaja & "','" & Convertyyyymmdd(gdatFechaActual) & "',1,'" & Trim(frmMainMdi.Tag) & "','" & strDescripGasto & "','" & _
                                "D','" & strCtaRetencion & "'," & CDec(dblMontoDetraccionMN) & ",'" & _
                                strCodFile & "','" & strCodAnalitica & "','" & strTipoAuxiliar & "','" & strCodAuxiliar & "','" & Codigo_Moneda_Local & " ','') }"
                        End If
                        
                        adoConn.Execute strSQLOrdenCajaMN
                        adoConn.Execute strSQLOrdenCajaDetalleMN
                        
                        If strSQLOrdenCajaDetalleMN2 <> "" Then
                            adoConn.Execute strSQLOrdenCajaDetalleMN2
                        End If
                        
                        'Actualiza orden recien creada con el codigo de gasto correspondiente 'ACR
                        .CommandText = "UPDATE MovimientoFondo SET NumGasto=" & adoRegistro("NumGasto") & " " & _
                        "WHERE NumOrdenCobroPago='" & strNumCaja & "' AND CodFondo='" & strpCodFondo & "' AND " & _
                        "CodAdministradora='" & gstrCodAdministradora & "'"
                        adoConn.Execute .CommandText
                                                
                        'Actualiza orden recien creada con el codigo de gasto correspondiente 'ACR
                        If strIndRetencion <> Valor_Indicador Then
                            .CommandText = "UPDATE MovimientoFondo SET NumGasto=" & adoRegistro("NumGasto") & ",ValorTipoCambio=" & adoConsulta("TipoCambioPago") & " " & _
                                "WHERE NumOrdenCobroPago='" & strNumCaja & "' AND CodFondo='" & strpCodFondo & "' AND " & _
                                "CodAdministradora='" & gstrCodAdministradora & "'"
                            adoConn.Execute .CommandText
                        End If
                    
                    End If
                
                    '*** Actualizar el número del parámetro **
                    .CommandText = "{ call up_ACActUltNumero('" & strpCodFondo & "','" & _
                        gstrCodAdministradora & "','" & Valor_NumOrdenCaja & "','" & strNumCaja & "') }"
                    adoConn.Execute .CommandText
                    
                End If
                adoConsulta.Close: Set adoConsulta = Nothing


                .CommandText = "UPDATE FondoGasto SET IndConfirma = 'X', FechaConfirma='" & Convertyyyymmdd(gdatFechaActual) & "' " & _
                    "WHERE CodCuenta = '" & strpCodGasto & "' AND " & _
                    "NumGasto=" & intpNumGasto & " AND CodFondo='" & strpCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
                adoConn.Execute .CommandText


'                '*** Actualizar Registro del Gasto ***
'                If gdatFechaActual = adoAuxiliar("FechaVencimiento") Then
'                    .CommandText = "UPDATE FondoGasto SET IndConfirma = 'X', FechaConfirma='" & Convertyyyymmdd(gdatFechaActual) & "' " & _
'                        "WHERE CodCuenta = '" & strpCodGasto & "' AND " & _
'                        "NumGasto=" & intpNumGasto & " AND CodFondo='" & strpCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
'                    adoConn.Execute .CommandText
'                ElseIf Trim(adoRegistro("CodCuenta")) <> Codigo_Cuenta_Comision_Fija Then
'                .CommandText = "UPDATE FondoGasto SET IndConfirma = 'X', FechaConfirma='" & Convertyyyymmdd(gdatFechaActual) & "' " & _
'                        "WHERE CodCuenta = '" & strpCodGasto & "' AND " & _
'                        "NumGasto=" & intpNumGasto & " AND CodFondo='" & strpCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
'                    adoConn.Execute .CommandText
'                End If
                
                
            End If
            adoRegistro.MoveNext
        Loop
        
        adoRegistro.Close: Set adoRegistro = Nothing
        
    End With
    
    Exit Sub
  
Ctrl_Error:
    MsgBox Mensaje_Proceso_NoExitoso, vbCritical
    Screen.MousePointer = vbDefault
    
End Sub

Public Sub GenerarAsientoFacturacion(ByVal strpCodFile As String, ByVal strpCodAnalitica As String, ByVal strpCodFondo As String, ByVal strpCodAdministradora, ByVal strpCodDetalleFile As String, ByVal strpCodDinamica As String, ByVal curpMontoAsiento As Currency, ByVal curpMontoNoGravado As Currency, ByVal dblpTipoCambio As Double, ByVal strpCodMoneda As String, ByVal strpDescripAsiento As String, ByVal strpCodModulo As String, Optional ByVal strpCodTipoComprobante As String, Optional ByVal strpNumComprobante As String, Optional ByVal strpTipoAuxiliar As String, Optional ByVal strpCodAuxiliar As String)

    Dim adoRegistro                 As ADODB.Recordset
    Dim curMontoMovimientoMN        As Currency, curMontoMovimientoME       As Currency
    Dim curMontoContable            As Currency
    Dim strIndDebeHaber             As String, strDescripMovimiento         As String
    Dim strNumAsiento               As String, strFechaGrabar               As String
    Dim strFechaCierre              As String, strFechaSiguiente            As String
    Dim intCantRegistros            As Integer
    
    With adoComm
        strNumAsiento = ObtenerSecuencialInversionOperacion(strpCodFondo, Valor_NumComprobante)
        strFechaCierre = Convertyyyymmdd(gdatFechaActual)
        strFechaSiguiente = Convertyyyymmdd(DateAdd("d", 1, gdatFechaActual))
        strFechaGrabar = strFechaCierre & Space(1) & Format(Time, "hh:mm")
                
        If IsNull(strpCodTipoComprobante) Then strpCodTipoComprobante = ""
        If IsNull(strpNumComprobante) Then strpNumComprobante = ""
                
        If IsNull(strpTipoAuxiliar) Then strpTipoAuxiliar = ""
        If IsNull(strpCodAuxiliar) Then strpCodAuxiliar = ""
                
        .CommandText = "SELECT COUNT(*) NumRegistros FROM DinamicaContable " & _
            "WHERE TipoOperacion='" & strpCodDinamica & "' AND CodFile='" & strpCodFile & "' AND (CodDetalleFile='" & _
            strpCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & strpCodAdministradora & "' AND CodMoneda = '" & _
            IIf(strpCodMoneda = Codigo_Moneda_Local, Codigo_Moneda_Local, Codigo_Moneda_Extranjero) & "'"
        Set adoRegistro = .Execute
    
        If Not adoRegistro.EOF Then
            If CInt(adoRegistro("NumRegistros")) > 0 Then
                intCantRegistros = CInt(adoRegistro("NumRegistros"))
            End If
        End If
        adoRegistro.Close
        
        '*** Cabecera ***
        .CommandText = "{ call up_ACAdicAsientoContable('" & _
            strpCodFondo & "','" & strpCodAdministradora & "','" & strNumAsiento & "','" & _
            strFechaGrabar & "','" & _
            gstrPeriodoActual & "','" & gstrMesActual & "','','" & _
            strpDescripAsiento & "','" & strpCodMoneda & "','" & _
            Codigo_Moneda_Local & "','" & strpCodTipoComprobante & "','" & strpNumComprobante & "'," & _
            (CDec(curpMontoAsiento * (1 + gdblTasaIgv)) + curpMontoNoGravado) & ",'" & Estado_Activo & "'," & _
            intCantRegistros & ",'" & _
            strFechaGrabar & "','" & _
            strpCodModulo & "',''," & _
            CDec(gdblTipoCambio) & ",'','','" & _
            strpDescripAsiento & "','','X','') }"
        adoConn.Execute .CommandText
        
        '*** Detalle Contable ***
        .CommandText = "SELECT NumSecuencial,TipoCuentaInversion,IndDebeHaber,DescripDinamica,CodCuenta FROM DinamicaContable " & _
            "WHERE TipoOperacion='" & strpCodDinamica & "' AND CodFile='" & strpCodFile & "' AND (CodDetalleFile='" & _
            strpCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & gstrCodAdministradora & "' AND CodMoneda = '" & _
            IIf(strpCodMoneda = Codigo_Moneda_Local, Codigo_Moneda_Local, Codigo_Moneda_Extranjero) & "' ORDER BY NumSecuencial"
            
        Set adoRegistro = .Execute

        Do While Not adoRegistro.EOF
            Select Case Trim(adoRegistro("TipoCuentaInversion"))
            
                Case Codigo_CtaXPagar
                    curMontoMovimientoMN = curpMontoAsiento + curpMontoNoGravado
                
                Case Codigo_CtaImpuesto
                    curMontoMovimientoMN = curpMontoAsiento * (gdblTasaIgv)
                
                Case Codigo_CtaXPagarEmitida
                    curMontoMovimientoMN = Round(curpMontoAsiento * (1 + gdblTasaIgv), 2) + curpMontoNoGravado
                
                Case Codigo_CtaProvGasto
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaInversion
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaProvInteres
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaInteres
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaCosto
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaIngresoOperacional
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaInteresVencido
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaVacCorrido
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaXCobrar
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaInteresCorrido
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaProvReajusteK
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaReajusteK
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaProvFlucMercado
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaFlucMercado
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaProvInteresVac
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaInteresVac
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaIntCorridoK
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaProvFlucK
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaFlucK
                    curMontoMovimientoMN = curpMontoAsiento
                    
                Case Codigo_CtaInversionTransito
                    curMontoMovimientoMN = curpMontoAsiento
                
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
                        
            '*** Movimiento ***
            .CommandText = "{ call up_ACAdicAsientoContableDetalle('" & _
                strNumAsiento & "','" & strpCodFondo & "','" & _
                gstrCodAdministradora & "'," & _
                CInt(adoRegistro("NumSecuencial")) & ",'" & _
                strFechaGrabar & "','" & _
                gstrPeriodoActual & "','" & _
                gstrMesActual & "','" & _
                strDescripMovimiento & "','" & _
                strIndDebeHaber & "','" & _
                Trim(adoRegistro("CodCuenta")) & "','" & _
                strpCodMoneda & "'," & _
                CDec(curMontoMovimientoMN) & "," & _
                CDec(curMontoMovimientoME) & "," & _
                CDec(curMontoContable) & ",'" & _
                strpCodFile & "','" & _
                strpCodAnalitica & "','" & _
                strpTipoAuxiliar & "','" & _
                strpCodAuxiliar & "') }"
            adoConn.Execute .CommandText
        
            '*** Saldos ***
            .CommandText = "{ call up_ACGenPartidaContableSaldos('" & _
                strpCodFondo & "','" & gstrCodAdministradora & "','" & _
                gstrPeriodoActual & "','" & gstrMesActual & "','" & _
                Trim(adoRegistro("CodCuenta")) & "','" & _
                strpCodFile & "','" & _
                strpCodAnalitica & "','" & _
                strFechaCierre & "','" & _
                strFechaSiguiente & "'," & _
                CDec(curMontoMovimientoMN) & "," & _
                CDec(curMontoMovimientoME) & "," & _
                CDec(curMontoContable) & ",'" & _
                strIndDebeHaber & "','" & _
                strpCodMoneda & "') }"
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
        
        '-- Verifica y ajusta posibles descuadres
        .CommandText = "{ call up_ACProcAsientoContableAjuste('" & _
                strpCodFondo & "','" & _
                strpCodAdministradora & "','" & _
                strNumAsiento & "') }"
        adoConn.Execute .CommandText
        
        '*** Actualizar el número del parámetro **
        adoComm.CommandText = "{ call up_ACActUltNumero('" & _
                    strpCodFondo & "','" & _
                    strpCodAdministradora & "','" & _
                    Valor_NumComprobante & "','" & _
                    strNumAsiento & "') }"
        adoConn.Execute .CommandText
        
        .CommandText = "SELECT COUNT(*) NumRegistros FROM DinamicaContable " & _
            "WHERE TipoOperacion='" & strpCodDinamica & "' AND CodFile='" & strpCodFile & "' AND (CodDetalleFile='" & _
            strpCodDetalleFile & "' OR CodDetalleFile='000') AND CodAdministradora='" & strpCodAdministradora & "' AND CodMoneda = '" & _
            IIf(strpCodMoneda = Codigo_Moneda_Local, Codigo_Moneda_Local, Codigo_Moneda_Extranjero) & "'"
        Set adoRegistro = .Execute
    
        If Not adoRegistro.EOF Then
            If CInt(adoRegistro("NumRegistros")) > 0 Then
                intCantRegistros = CInt(adoRegistro("NumRegistros"))
            End If
        End If
        adoRegistro.Close
        
    End With
    
End Sub


Public Function ObtenerPlanContableVersion() As Integer

    Dim adoRegistro As ADODB.Recordset
    
    Dim intVersion As Integer

    Set adoRegistro = New ADODB.Recordset

    ObtenerPlanContableVersion = -1
    
    With adoComm
    
        .CommandText = "SELECT dbo.uf_CNObtenerPlanContableVigente('" & gstrCodAdministradora & "') AS NumVersion"
        
        Set adoRegistro = .Execute
    
    End With
    
    If Not adoRegistro.EOF Then
    
        Do Until adoRegistro.EOF
        
            intVersion = CInt(adoRegistro("NumVersion").Value)
            
            adoRegistro.MoveNext
        
        Loop
        
        ObtenerPlanContableVersion = intVersion
    
    End If
 
End Function

Public Sub ContabilizarRegistroCompra(ByVal strpCodFondo As String, ByVal strpCodAdministradora, ByVal strpFechaProceso As String, ByVal strpNroRegCompra As String, ByRef strMsgError As String)

    Dim strFechaGrabar               As String
    Dim strFechaSiguiente            As String
    
    frmMainMdi.stbMdi.Panels(3).Text = "Contabilizando Registro de Compras..."
    
    strFechaGrabar = Convertyyyymmdd(gdatFechaActual)
    strFechaSiguiente = Convertyyyymmdd(DateAdd("d", 1, gdatFechaActual))
        
    'CONTABILIZACIÓN DEL REGISTRO DE COMPRAS
    adoComm.CommandText = "{ call up_CNContabilizarRegistroCompra('" & strpCodFondo & "','" & gstrCodAdministradora & "','" & strFechaGrabar & "','" & strpNroRegCompra & " ') }"
    adoConn.Execute adoComm.CommandText

    If strMsgError = "" Then strMsgError = err.Description
    Screen.MousePointer = vbDefault
    
End Sub


