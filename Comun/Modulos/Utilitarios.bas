Attribute VB_Name = "Utilitarios"
Option Explicit

Public Sub ColorControlHabilitado(ByRef ctrlObjeto As Control)

    With ctrlObjeto
        .BackColor = &H80000005
    End With
    
End Sub

Public Sub ColorControlFechaHabilitado(ByRef ctrlObjeto As DTPicker)

    With ctrlObjeto
        .CalendarBackColor = &H80000005
    End With
    
End Sub
Public Sub ColorControlDeshabilitado(ByRef ctrlObjeto As Control)

    With ctrlObjeto
        .BackColor = &H80000016
    End With
    
End Sub

Public Sub ColorControlFechaDeshabilitado(ByRef ctrlObjeto As DTPicker)

    With ctrlObjeto
        .CalendarBackColor = &H80000016
    End With
    
End Sub
Public Sub FormatoCajaTexto(ByRef txtControl As TextBox, ByVal intDecimales As Integer)

    Dim strFormato As String
    Dim intLongitud As Integer
    
    If intDecimales = 0 Then
        strFormato = "###,###,###,###,###,##0"
    Else
        strFormato = "###,###,###,###,###,##0." & String(intDecimales, "0")
    End If
    
    With txtControl
        intLongitud = Len(txtControl)
        If intLongitud = 0 Then .Text = "0"
        .Text = Format(.Text, strFormato)
    End With
    
    intLongitud = Len(txtControl)
    If txtControl.SelStart < (intLongitud - (intDecimales + 1)) Then
        txtControl.SelStart = (intLongitud - (intDecimales + 1))
    End If
    
End Sub

Public Sub DarFormatoValor(ByRef ctrlValor As Variant, ByVal intDecimales As Integer)

    Dim strFormato As String
    Dim intLongitud As Integer
    
    If intDecimales = 0 Then
        strFormato = "###,###,###,###,###,##0"
    Else
        strFormato = "###,###,###,###,###,##0." & String(intDecimales, "0")
    End If
    
    ctrlValor = Format(ctrlValor, strFormato)
            
End Sub

Public Function NumAleatorio(ByVal intNumDigitos As Integer) As String
    
    Dim intValor    As Integer, intCont As Integer
    Dim strFormato  As String, strValor As String
    
    strValor = Valor_Caracter
    If intNumDigitos > 0 Then
    
        intCont = 0
        
        Do While intCont < 3000
        
            Randomize
        
            Select Case intNumDigitos
                Case 1
                    intValor = Int((9 * Rnd) + 1)
                    strValor = CStr(intValor)
            
                Case 2
                    intValor = Int((9 * Rnd) + 1)
                    strValor = CStr(intValor)
            
                    intValor = Int((9 * Rnd) + 1)
                    strValor = Trim(strValor) & CStr(intValor)
                    
                Case 3
                    intValor = Int((9 * Rnd) + 1)
                    strValor = CStr(intValor)
            
                    intValor = Int((9 * Rnd) + 1)
                    strValor = Trim(strValor) & CStr(intValor)
                    
                    intValor = Int((9 * Rnd) + 1)
                    strValor = Trim(strValor) & CStr(intValor)
                    
                Case 4
                    intValor = Int((9 * Rnd) + 1)
                    strValor = CStr(intValor)
            
                    intValor = Int((9 * Rnd) + 1)
                    strValor = Trim(strValor) & CStr(intValor)
                    
                    intValor = Int((9 * Rnd) + 1)
                    strValor = Trim(strValor) & CStr(intValor)
                    
                    intValor = Int((9 * Rnd) + 1)
                    strValor = Trim(strValor) & CStr(intValor)
                    
                Case 5
                    intValor = Int((9 * Rnd) + 1)
                    strValor = CStr(intValor)
                    
                    intValor = Int((9 * Rnd) + 1)
                    strValor = Trim(strValor) & CStr(intValor)
                    
                    intValor = Int((9 * Rnd) + 1)
                    strValor = Trim(strValor) & CStr(intValor)
                    
                    intValor = Int((9 * Rnd) + 1)
                    strValor = Trim(strValor) & CStr(intValor)
                    
                    intValor = Int((9 * Rnd) + 1)
                    strValor = Trim(strValor) & CStr(intValor)
                    
                Case 6
                    intValor = Int((9 * Rnd) + 1)
                    strValor = CStr(intValor)
                    
                    intValor = Int((9 * Rnd) + 1)
                    strValor = Trim(strValor) & CStr(intValor)
                    
                    intValor = Int((9 * Rnd) + 1)
                    strValor = Trim(strValor) & CStr(intValor)
                    
                    intValor = Int((9 * Rnd) + 1)
                    strValor = Trim(strValor) & CStr(intValor)
                    
                    intValor = Int((9 * Rnd) + 1)
                    strValor = Trim(strValor) & CStr(intValor)
                    
                    intValor = Int((9 * Rnd) + 1)
                    strValor = Trim(strValor) & CStr(intValor)
                
                Case 7
                    intValor = Int((9 * Rnd) + 1)
                    strValor = CStr(intValor)
                    
                    intValor = Int((9 * Rnd) + 1)
                    strValor = Trim(strValor) & CStr(intValor)
                    
                    intValor = Int((9 * Rnd) + 1)
                    strValor = Trim(strValor) & CStr(intValor)
                    
                    intValor = Int((9 * Rnd) + 1)
                    strValor = Trim(strValor) & CStr(intValor)
                    
                    intValor = Int((9 * Rnd) + 1)
                    strValor = Trim(strValor) & CStr(intValor)
                    
                    intValor = Int((9 * Rnd) + 1)
                    strValor = Trim(strValor) & CStr(intValor)
                    
                    intValor = Int((9 * Rnd) + 1)
                    strValor = Trim(strValor) & CStr(intValor)
                                    
                Case 8
                    intValor = Int((9 * Rnd) + 1)
                    strValor = CStr(intValor)
                    
                    intValor = Int((9 * Rnd) + 1)
                    strValor = Trim(strValor) & CStr(intValor)
                    
                    intValor = Int((9 * Rnd) + 1)
                    strValor = Trim(strValor) & CStr(intValor)
                    
                    intValor = Int((9 * Rnd) + 1)
                    strValor = Trim(strValor) & CStr(intValor)
                    
                    intValor = Int((9 * Rnd) + 1)
                    strValor = Trim(strValor) & CStr(intValor)
                    
                    intValor = Int((9 * Rnd) + 1)
                    strValor = Trim(strValor) & CStr(intValor)
                    
                    intValor = Int((9 * Rnd) + 1)
                    strValor = Trim(strValor) & CStr(intValor)
                    
                    intValor = Int((9 * Rnd) + 1)
                    strValor = Trim(strValor) & CStr(intValor)
                
                Case Else
                    intValor = Int((9 * Rnd) + 1)
                    strValor = CStr(intValor)
                    
                    intValor = Int((9 * Rnd) + 1)
                    strValor = Trim(strValor) & CStr(intValor)
                    
                    intValor = Int((9 * Rnd) + 1)
                    strValor = Trim(strValor) & CStr(intValor)
                    
                    intValor = Int((9 * Rnd) + 1)
                    strValor = Trim(strValor) & CStr(intValor)
                    
                    intValor = Int((9 * Rnd) + 1)
                    strValor = Trim(strValor) & CStr(intValor)
                    
                    intValor = Int((9 * Rnd) + 1)
                    strValor = Trim(strValor) & CStr(intValor)
                    
                    intValor = Int((9 * Rnd) + 1)
                    strValor = Trim(strValor) & CStr(intValor)
                    
                    intValor = Int((9 * Rnd) + 1)
                    strValor = Trim(strValor) & CStr(intValor)
                    
                    intValor = Int((9 * Rnd) + 1)
                    strValor = Trim(strValor) & CStr(intValor)
                    
                    intValor = Int((9 * Rnd) + 1)
                    strValor = Trim(strValor) & CStr(intValor)
            End Select
            
            intCont = intCont + 1
        Loop
    
        strFormato = String(intNumDigitos, "0")
        
        strValor = Format(strValor, strFormato)
        strValor = Left(strValor, intNumDigitos)
        
    End If
                
    NumAleatorio = strValor
            
End Function
Public Sub FormatoMillarEtiqueta(ByRef lblControl As Label, ByVal intDecimales As Integer)

    Dim strFormato As String
    Dim intLongitud As Integer
    
    If intDecimales = 0 Then
        strFormato = "###,###,###,###,###,##0"
    Else
        strFormato = "###,###,###,###,###,##0." & String(intDecimales, "0")
    End If
    
    With lblControl
        .Caption = Format(.Caption, strFormato)
    End With
            
End Sub

Public Sub FormatoEtiqueta(ByRef lblControl As Label, ByVal strAlineacion As AlignmentConstants)
On Error GoTo CtrlError
    With lblControl
        .Font = "MS Sans Serif"
        .FontBold = True
        .ForeColor = &H800000
        .Alignment = strAlineacion
    End With
CtrlError:
End Sub

Public Sub FormatoMarco(ByRef fraControl As Frame)

    With fraControl
        .Font = "Arial"
        .FontBold = True
        .ForeColor = &H80000012
    End With
    
End Sub
Public Sub ValidaCajaTexto(ByRef intTeclaPresionada As Integer, ByVal strTipo As String, Optional ByVal txtControl As TextBox, Optional ByVal intDecimales As Integer)
            
    Dim intTmpTecla     As Integer
    
    If intTeclaPresionada = vbKeyReturn Then Exit Sub
                
            If intTeclaPresionada = Asc(UCase(Chr(intTeclaPresionada))) Then        'HMC
                intTmpTecla = intTeclaPresionada                                    'HMC
            ElseIf intTeclaPresionada = Asc(LCase(Chr(intTeclaPresionada))) Then    'HMC
                intTmpTecla = intTeclaPresionada                                    'HMC
            End If                                                                  'HMC
    
    Select Case strTipo
        Case "N" 'Solo Números
            If intTeclaPresionada < Asc("0") Or intTeclaPresionada > Asc("9") Then
                If intTeclaPresionada <> 8 Then
                    intTeclaPresionada = 0
                    Beep
                End If
            End If
        
        Case "L" 'Solo Letras
                        
            '*** Mayúsculas ***
            intTeclaPresionada = Asc(UCase(Chr(intTeclaPresionada)))
            If intTeclaPresionada < Asc("A") Or intTeclaPresionada > Asc("Z") Then
                If intTeclaPresionada <> 8 And intTeclaPresionada <> 32 Then
                    intTeclaPresionada = 0
                    Beep
                End If
            End If
            
            '*** Minúsculas ***
            intTeclaPresionada = Asc(LCase(Chr(intTeclaPresionada)))
            If intTeclaPresionada < Asc("a") Or intTeclaPresionada > Asc("z") Then
                If intTeclaPresionada <> 8 And intTeclaPresionada <> 32 Then
                    intTeclaPresionada = 0
                    Beep
                End If
            End If
            
            If intTeclaPresionada = 0 Then intTeclaPresionada = 0 Else intTeclaPresionada = intTmpTecla 'HMC
        
        Case "A" ' Alfanumérico
            If (intTeclaPresionada < Asc("0") Or intTeclaPresionada > Asc("9")) And (intTeclaPresionada < Asc("A") Or intTeclaPresionada > Asc("Z")) And (intTeclaPresionada < Asc("a") Or intTeclaPresionada > Asc("z")) Then
                If intTeclaPresionada <> 8 Then
                    intTeclaPresionada = 0
                    Beep
                End If
            End If
'            If intTeclaPresionada < Asc("A") Or intTeclaPresionada > Asc("Z") Then
'                If intTeclaPresionada <> 8 Then
'                    intTeclaPresionada = 0
'                    Beep
'                End If
'            End If
'            If intTeclaPresionada < Asc("a") Or intTeclaPresionada > Asc("z") Then
'                If intTeclaPresionada <> 8 Then
'                    intTeclaPresionada = 0
'                    Beep
'                End If
'            End If
      
        Case "LC" 'Solo Letras y Caracteres
            If intTeclaPresionada < Asc("A") Or intTeclaPresionada > Asc("Z") Then
                If intTeclaPresionada <> 8 Then
                    intTeclaPresionada = 0
                    Beep
                End If
            End If
            If intTeclaPresionada < Asc("a") Or intTeclaPresionada > Asc("z") Then
                If intTeclaPresionada <> 8 Then
                    intTeclaPresionada = 0
                    Beep
                End If
            End If
            
        Case "AC" 'Alfanumerico y Caracteres
            If intTeclaPresionada < Asc("0") Or intTeclaPresionada > Asc("9") Then
                If intTeclaPresionada <> 8 Then
                    intTeclaPresionada = 0
                    Beep
                End If
            End If
            If intTeclaPresionada < Asc("A") Or intTeclaPresionada > Asc("Z") Then
                If intTeclaPresionada <> 8 Then
                    intTeclaPresionada = 0
                    Beep
                End If
            End If
            If intTeclaPresionada < Asc("a") Or intTeclaPresionada > Asc("z") Then
                If intTeclaPresionada <> 8 Then
                    intTeclaPresionada = 0
                    Beep
                End If
            End If
            
        Case "M" 'Solo Números con formato
            Dim intPosicion As Integer
            Dim intPosicionPunto As Integer
            Dim intLongitudTexto As Integer
            Dim intLongitudSeleccion As Integer
            
            With txtControl
                If intTeclaPresionada = Asc(".") Then
                    intLongitudTexto = Len(.Text)
                    If .SelLength <= intLongitudTexto Then
                        intTeclaPresionada = 0
                        intPosicion = InStr(txtControl, ".")
                        If intPosicion > 0 Then .SelStart = intPosicion
                    Else
                        .SelText = "0.00"
                        intPosicion = InStr(txtControl, ".")
                        If intPosicion > 0 Then .SelStart = intPosicion
                    End If
                Else
                    If intTeclaPresionada < Asc("0") Or intTeclaPresionada > Asc("9") Then
                        intTeclaPresionada = 0
                        Beep
                    End If
                    
                    If intTeclaPresionada >= Asc("0") And intTeclaPresionada <= Asc("9") Then
                    
                        intLongitudTexto = Len(txtControl)
                        If .SelLength < intLongitudTexto Then
                            intPosicionPunto = InStr(txtControl, ".")
                            intPosicion = .SelStart
                            If intPosicion >= intPosicionPunto Then
                                .SelLength = 1
                            End If
                            
                            If intPosicion >= (intPosicionPunto + intDecimales) Then
                                intTeclaPresionada = 0
                                Beep
                            End If
                        End If
                    End If
                End If
            End With
                                
    End Select
    
End Sub
    
Function ValiText(Keyin As Integer, Validatestring As String, Editable As Boolean) As Integer
    'Validacion HMC
    Dim Validatelist As String
    Dim Keyout As Integer
    
    Select Case Validatestring
    
        Case "L"    ' Letras
        
            If Editable = True Then
                Validatelist = UCase("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ ") & Chr(8) 'SI SE USA EL BACKSPACE
            Else
                Validatelist = UCase("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ ")
            End If
    
        Case "N"    ' Numeros
    
            If Editable = True Then
                Validatelist = UCase("0123546789- ") & Chr(8) 'SI SE USA EL BACKSPACE
            Else
                Validatelist = UCase("0123456789- ")
            End If
    
        Case "AN"    ' AlfaNumerico
    
            If Editable = True Then
                Validatelist = UCase("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123546789-/ ") & Chr(8) 'SI SE USA EL BACKSPACE
            Else
                Validatelist = UCase("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789-/ ")
            End If
        
    End Select
    
    If InStr(1, Validatelist, UCase(Chr(Keyin)), 1) > 0 Then
        Keyout = Keyin
    Else
        Keyout = 0
        Beep
    End If
    
    ValiText = Keyout
            
End Function

Public Function FormatoConasev_9(ByRef lblControl As String, ByVal intLongitud As Integer) As String

    Dim LongitudControl         As Integer, intContadorPrincipal           As Integer
    Dim intContadorSecundario   As Integer
    
    LongitudControl = Len(lblControl)
    
    For intContadorPrincipal = 1 To LongitudControl
    
        If Mid(lblControl, intContadorPrincipal, 1) = "." Then 'Buscamos el separador decimal
            lblControl = Replace(lblControl, ".", "", 1, Len(lblControl))
            If Len(Mid(lblControl, 1, intContadorPrincipal)) <= intLongitud Then
                For intContadorSecundario = 1 To intLongitud - Len(Mid(lblControl, 1, intContadorPrincipal - 1))
                    lblControl = 0 & lblControl
                Next
            Else
                lblControl = "Error"
            End If
            
            Exit For
        End If
    Next
       
    FormatoConasev_9 = lblControl 'SOLO para ENTEROS

End Function

Public Function FormatoConasev_9_V9(ByRef lblControl As String, ByVal intLongitud1 As Integer, Optional ByVal intLongitud2 As Integer) As String

    Dim LongitudControl             As String, intContadorPrincipal             As Integer
    Dim intContadorSecundario       As Integer
    Dim strNewlblControl            As String, strOldlblControl                 As String
    Dim strSigno                    As String
    
    strSigno = Trim(IIf(Mid(Trim(lblControl), 1, 1) = "-" Or Mid(Trim(lblControl), 1, 1) = "+", Mid(Trim(lblControl), 1, 1), ""))
    
    strOldlblControl = Replace(lblControl, "-", "", 1, Len(lblControl))
    strNewlblControl = Replace(lblControl, "-", "", 1, Len(lblControl))
    
    For intContadorPrincipal = 1 To Len(Trim(lblControl))
        If Mid(strNewlblControl, intContadorPrincipal, 1) = "." Then                      'Buscamos el separador decimal
            If Len(Mid(strNewlblControl, 1, intContadorPrincipal)) <= intLongitud1 Then
                For intContadorSecundario = 1 To intLongitud1 - Len(Mid(strNewlblControl, 1, intContadorPrincipal - 1))
                    strNewlblControl = 0 & strNewlblControl               ' Agregamos 0 a la izquierda
                Next
            End If
            If Len(Mid(strOldlblControl, intContadorPrincipal, (Len(Trim(strOldlblControl)) - intContadorPrincipal))) <= intLongitud2 Then
                For intContadorSecundario = 1 To intLongitud2 - Len(Mid(strOldlblControl, intContadorPrincipal, (Len(Trim(strOldlblControl)) - intContadorPrincipal)))
                    strNewlblControl = strNewlblControl & 0               ' Agregamos 0 a la derecha
                Next intContadorSecundario
            End If
            Exit For
        End If
    Next
       
    strNewlblControl = Replace(strNewlblControl, ".", "", 1, Len(lblControl)) 'Eliminamos el . decimal
    FormatoConasev_9_V9 = (strSigno & strNewlblControl)                       'Agregamos el signo


End Function
'Public Function FormatoConasev_9(ByRef lblControl As String, ByVal intLongitud As Integer) As String
'
'    Dim LongitudControl         As Integer, intContadorPrincipal           As Integer
'    Dim intContadorSecundario   As Integer
'
'    LongitudControl = Len(lblControl)
'
'    For intContadorPrincipal = 1 To LongitudControl
'
'        If Mid(lblControl, intContadorPrincipal, 1) = "." Then 'Buscamos el separador decimal
'            lblControl = Replace(lblControl, ".", "", 1, Len(lblControl))
'            If Len(Mid(lblControl, 1, intContadorPrincipal)) <= intLongitud Then
'                For intContadorSecundario = 1 To intLongitud - Len(Mid(lblControl, 1, intContadorPrincipal - 1))
'                    lblControl = 0 & lblControl
'                Next
'            Else
'                lblControl = "Error"
'            End If
'
'            Exit For
'        End If
'
'    Next
'
'    FormatoConasev_9 = lblControl 'SOLO para ENTEROS
'
'End Function




Public Sub FormatoGrilla(ByRef DBGrid As TDBGrid)
                
                
        '************* Normal *********************
        DBGrid.Styles(0).BackColor = vbWindowBackground
        DBGrid.Styles(0).Font.Bold = True
        DBGrid.Styles(0).ForeColor = &H80000012
        
        '************* Cabecera de columnas *********
        DBGrid.Styles(1).BackColor = vbActiveTitleBar
        DBGrid.Styles(1).Font.Bold = True
        DBGrid.Styles(1).ForeColor = vbHighlightText
        
        '************* Seleccion ********************
        DBGrid.Styles(3).BackColor = vbWindowText
        DBGrid.Styles(3).Font.Bold = True
        DBGrid.Styles(3).Font.Italic = False
        DBGrid.Styles(3).ForeColor = vbWindowBackground
        
        '************* Caption *******************
        DBGrid.Styles(4).BackColor = vbActiveTitleBar
        DBGrid.Styles(4).Font.Bold = True
        DBGrid.Styles(4).ForeColor = vbHighlightText
        
        DBGrid.HeadingStyle = DBGrid.Styles(1)
        DBGrid.HighlightRowStyle = DBGrid.Styles(3)
        DBGrid.SelectedStyle = DBGrid.Styles(3)
        DBGrid.Style = DBGrid.Styles(0)
        
        DBGrid.Splits(0).HeadingStyle = DBGrid.Styles(1)
        DBGrid.Splits(0).HighlightRowStyle = DBGrid.Styles(3)
        DBGrid.Splits(0).SelectedStyle = DBGrid.Styles(3)
        DBGrid.Splits(0).Style = DBGrid.Styles(0)
        
        
        
        DBGrid.Refresh
    
End Sub

Public Sub ImprimeComprobanteCobro(ByVal strCodFondo As String, ByVal numRegistro As Integer, ByVal strTD As String, ByVal strNumDoc As String, ByVal strSerie As String, ByRef strMsgError As String, Optional strIndTotal As String = Valor_Caracter)
Dim wcont, wsum         As Integer
Dim wfila, wcolu        As Integer
Dim csql                As String

Dim rst                 As New ADODB.Recordset
Dim rstObj              As New ADODB.Recordset
Dim p                   As Object
Dim indPrinter          As Boolean
Dim strCampos           As String
Dim intScale            As Integer

Dim numEntreLineasAdicional As Integer
Dim i As Integer
 
On Error GoTo err

    '******************************************************
    'SELECCIONAMOS IMPRESORA
    '******************************************************
        indPrinter = False
        If strTD = "01" Or strTD = "07" Or strTD = "08" Then 'FACTURA
            For Each p In Printers
               If Right(UCase(p.DeviceName), 7) = "FACTURA" Then
                  Set Printer = p
                  indPrinter = True
                  Exit For
               End If
            Next p
        ElseIf strTD = "03" Then 'BOLETA
                For Each p In Printers
                    If Right(UCase(p.DeviceName), 6) = "BOLETA" Then
                        Set Printer = p
                        indPrinter = True
                        Exit For
                    End If
                Next p
        End If
        
    
        If indPrinter = False Then
        For Each p In Printers
                    If Right(UCase(p.DeviceName), 7) = "GENERAL" Then
                        Set Printer = p
                        indPrinter = True
                        Exit For
                    End If
                Next p
        End If
                
        If indPrinter = False Then
            For Each p In Printers
               If p.Port = "LPT1:" Then
                  Set Printer = p
                  Exit For
               End If
            Next p
        End If

        '******************************************************
        intScale = 6
        Printer.ScaleMode = intScale
        Printer.FontName = "Roman 12cpi"
        Printer.FontBold = False
        Printer.FontSize = 8

    '**********************************************************************
    'CABECERA
    '**********************************************************************
        'seleccionar los campos a imprimir de la tabla objDocventas
        csql = "SELECT GlsCampo " & _
               "FROM objRegistroVenta " & _
               "WHERE CodAdministradora = '" & gstrCodAdministradora & "' AND CodFondo = '" & strCodFondo & "' " & _
                 "AND tipoObj = 'C' AND GlsCampo <> '' AND indImprime = 1 " & _
                 "AND CodTipoComprobante = '" & strTD & "' AND SerieComprobante = '" & strSerie & "' " & _
               "ORDER BY impY, impX"
        rst.Open csql, gstrConnectConsulta, adOpenForwardOnly, adLockReadOnly
        Do While Not rst.EOF
            strCampos = strCampos & "" & rst.Fields("GlsCampo") & ","
            rst.MoveNext
        Loop
        If strCampos = "" Then
            strMsgError = "No existe configuración de impresión para el Tipo de Documento y la Serie"
            GoTo err
        End If
        strCampos = Left(strCampos, Len(strCampos) - 1)
        rst.Close
        
        'traemos la data de lo campos seleccionados arriba
        csql = "SELECT " & strCampos & " FROM RegistroVenta " & _
               "WHERE CodAdministradora = '" & gstrCodAdministradora & "' AND CodFondo = '" & strCodFondo & "' " & _
                 "AND NumRegistro = " & numRegistro
        rst.Open csql, gstrConnectConsulta, adOpenStatic, adLockReadOnly
        
        If Not rst.EOF Then
            For i = 0 To rst.Fields.Count - 1
                'traigo datos de impreison por en nombre del campo de la tabla objDocventas
                csql = "SELECT '" & (rst.Fields(i) & "") & "' AS valor, tipoDato, impLongitud, impX, impY, decimales,0 AS intNumFilas FROM objRegistroVenta " & _
                       "WHERE CodAdministradora = '" & gstrCodAdministradora & "' AND CodFondo = '" & strCodFondo & "' " & _
                         "AND GlsCampo = '" & rst.Fields(i).Name & "' AND indImprime = 1 " & _
                         "AND CodTipoComprobante = '" & strTD & "' AND SerieComprobante = '" & strSerie & "'"
        
                rstObj.Open csql, gstrConnectConsulta, adOpenForwardOnly, adLockReadOnly
                If Not rstObj.EOF Then
                    ImprimeXY rstObj.Fields("valor") & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY"), rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), Val("" & rstObj.Fields("intNumFilas")), strMsgError
                    If strMsgError <> "" Then GoTo err
                End If
                rstObj.Close
            Next
        End If
        rst.Close
        '**********************************************************************
        
    '**********************************************************************
    'DETALLE
    '**********************************************************************
    Dim strTipoDato As String
    Dim intLong As Integer
    Dim intX    As Integer
    Dim intY As Integer
    Dim intDec As Integer
    
    If strIndTotal = Valor_Caracter Then
        wcont = 1
        wsum = 0
        
        'seleccionar los campos a imprimir de la tabla objDocventas ' la forma de Pago se guarda como una glosa en la tabla doc ventas
        strCampos = ""
        csql = "SELECT GlsCampo " & _
               "FROM objRegistroVenta " & _
               "WHERE CodAdministradora = '" & gstrCodAdministradora & "' AND CodFondo = '" & strCodFondo & "' " & _
                 "AND tipoObj = 'D' AND GlsCampo <> '' AND indImprime = 1 " & _
                 "AND CodTipoComprobante = '" & strTD & "' AND SerieComprobante = '" & strSerie & "' " & _
               "ORDER BY impY,impX"
        rst.Open csql, gstrConnectConsulta, adOpenForwardOnly, adLockReadOnly
        Do While Not rst.EOF
            strCampos = strCampos & "" & rst.Fields("GlsCampo") & ","
            rst.MoveNext
        Loop
        If strCampos = "" Then
            strMsgError = "No existe configuración de impresión para el Tipo de Documento y la Serie"
            GoTo err
        End If
        strCampos = Left(strCampos, Len(strCampos) - 1)
        rst.Close
        
        'traemos la data de los campos seleccionados arriba
        csql = "SELECT " & strCampos & " FROM RegistroVentaDetalle " & _
               "WHERE CodAdministradora = '" & gstrCodAdministradora & "' AND CodFondo = '" & strCodFondo & "' " & _
                 "AND NumRegistro = " & numRegistro
        rst.Open csql, gstrConnectConsulta, adOpenStatic, adLockReadOnly
        wsum = 0
        Do While Not rst.EOF
            For i = 0 To rst.Fields.Count - 1
                'traigo datos de impresion por en nombre del campo de la tabla objDocventas
                csql = "SELECT '" & (rst.Fields(i) & "") & "' AS valor, tipoDato, impLongitud, impX, impY, decimales,0 AS intNumFilas FROM objRegistroVenta " & _
                       "WHERE CodAdministradora = '" & gstrCodAdministradora & "' AND CodFondo = '" & strCodFondo & "' " & _
                         "AND GlsCampo = '" & rst.Fields(i).Name & "' AND indImprime = 1 " & _
                         "AND CodTipoComprobante = '" & strTD & "' AND SerieComprobante = '" & strSerie & "'"
                rstObj.Open csql, gstrConnectConsulta, adOpenForwardOnly, adLockReadOnly
                If Not rstObj.EOF Then
                    strTipoDato = rstObj.Fields("tipoDato")
                    intLong = rstObj.Fields("impLongitud")
                    intX = rstObj.Fields("impX")
                    intY = rstObj.Fields("impY")
                    intDec = Val("" & rstObj.Fields("Decimales"))
    

                    ImprimeXY rst.Fields(i) & "", strTipoDato, intLong, intY + wsum, intX, intDec, 0, strMsgError
                    If strMsgError <> "" Then GoTo err
                        
                End If
                rstObj.Close
            Next
            rst.MoveNext
            wsum = wsum + IIf(intScale = 6, 4, 1) + numEntreLineasAdicional
        Loop
        rst.Close
    End If
    
    
    '**********************************************************************
    'DETALLE SOLO TOTALES
    '**********************************************************************
    If strIndTotal = Valor_Indicador Then
        wcont = 1
        wsum = 0
        
        'seleccionar los campos a imprimir de la tabla objDocventas ' la forma de Pago se guarda como una glosa en la tabla doc ventas
        strCampos = ""
        csql = "SELECT GlsCampo " & _
               "FROM objRegistroVenta " & _
               "WHERE CodAdministradora = '" & gstrCodAdministradora & "' AND CodFondo = '" & strCodFondo & "' " & _
                 "AND tipoObj = 'E' AND GlsCampo <> '' AND indImprime = 1 " & _
                 "AND CodTipoComprobante = '" & strTD & "' AND SerieComprobante = '" & strSerie & "' " & _
               "ORDER BY impY,impX"
        rst.Open csql, gstrConnectConsulta, adOpenForwardOnly, adLockReadOnly
        Do While Not rst.EOF
            strCampos = strCampos & "" & rst.Fields("GlsCampo") & ","
            rst.MoveNext
        Loop
        If strCampos = "" Then
            strMsgError = "No existe configuración de impresión para el Tipo de Documento y la Serie"
            GoTo err
        End If
        strCampos = Left(strCampos, Len(strCampos) - 1)
        rst.Close
        
        'traemos la data de los campos seleccionados arriba
        csql = "SELECT " & strCampos & " FROM RegistroVenta " & _
               "WHERE CodAdministradora = '" & gstrCodAdministradora & "' AND CodFondo = '" & strCodFondo & "' " & _
                 "AND NumRegistro = " & numRegistro
        rst.Open csql, gstrConnectConsulta, adOpenStatic, adLockReadOnly
        wsum = 0
        Do While Not rst.EOF
            For i = 0 To rst.Fields.Count - 1
                'traigo datos de impresion por en nombre del campo de la tabla objDocventas
                csql = "SELECT '" & (rst.Fields(i) & "") & "' AS valor, tipoDato, impLongitud, impX, impY, decimales,0 AS intNumFilas FROM objRegistroVenta " & _
                       "WHERE CodAdministradora = '" & gstrCodAdministradora & "' AND CodFondo = '" & strCodFondo & "' " & _
                         "AND GlsCampo = '" & rst.Fields(i).Name & "' AND indImprime = 1 AND tipoObj = 'E' " & _
                         "AND CodTipoComprobante = '" & strTD & "' AND SerieComprobante = '" & strSerie & "'"
                rstObj.Open csql, gstrConnectConsulta, adOpenForwardOnly, adLockReadOnly
                If Not rstObj.EOF Then
                    strTipoDato = rstObj.Fields("tipoDato")
                    intLong = rstObj.Fields("impLongitud")
                    intX = rstObj.Fields("impX")
                    intY = rstObj.Fields("impY")
                    intDec = Val("" & rstObj.Fields("Decimales"))
    

                    ImprimeXY rst.Fields(i) & "", strTipoDato, intLong, intY + wsum, intX, intDec, 0, strMsgError
                    If strMsgError <> "" Then GoTo err
                        
                End If
                rstObj.Close
            Next
            rst.MoveNext
            wsum = wsum + IIf(intScale = 6, 4, 1) + numEntreLineasAdicional
        Loop
        rst.Close
    End If

                
    '------------------------------------------------------------------------------------------------
    'IMPRIME TOTALES
    '------------------------------------------------------------------------------------------------
        wsum = 0
        
        strCampos = ""
        csql = "SELECT GlsCampo " & _
               "FROM objRegistroVenta " & _
               "WHERE CodAdministradora = '" & gstrCodAdministradora & "' AND CodFondo = '" & strCodFondo & "' " & _
                 "AND tipoObj = 'T' AND GlsCampo <> '' AND indImprime = 1 " & _
                 "AND CodTipoComprobante = '" & strTD & "' AND SerieComprobante = '" & strSerie & "' " & _
               "ORDER BY impY, impX"
        rst.Open csql, gstrConnectConsulta, adOpenForwardOnly, adLockReadOnly
        Do While Not rst.EOF
            strCampos = strCampos & "" & rst.Fields("GlsCampo") & ","
            rst.MoveNext
        Loop
        If Len(strCampos) > 0 Then
            strCampos = Left(strCampos, Len(strCampos) - 1)
        End If
        rst.Close
        
        'traemos la data de lo campos seleccionados arriba
        If Len(strCampos) > 0 Then
            csql = "SELECT " & strCampos & " FROM RegistroVenta " & _
                   "WHERE CodAdministradora = '" & gstrCodAdministradora & "' AND CodFondo = '" & strCodFondo & "' " & _
                     "AND NumRegistro = " & numRegistro
            rst.Open csql, gstrConnectConsulta, adOpenStatic, adLockReadOnly
            
            If Not rst.EOF Then
                For i = 0 To rst.Fields.Count - 1
                    csql = "SELECT '" & (rst.Fields(i) & "") & "' AS valor, tipoDato, impLongitud, impX, impY, decimales,0 AS intNumFilas FROM objRegistroVenta " & _
                           "WHERE CodAdministradora = '" & gstrCodAdministradora & "' AND CodFondo = '" & strCodFondo & "' " & _
                             "AND GlsCampo = '" & rst.Fields(i).Name & "' AND indImprime = 1 " & _
                             "AND CodTipoComprobante = '" & strTD & "' AND SerieComprobante = '" & strSerie & "'"
                    
                    
                    rstObj.Open csql, gstrConnectConsulta, adOpenForwardOnly, adLockReadOnly
                    If Not rstObj.EOF Then
                        ImprimeXY rstObj.Fields("valor") & "", rstObj.Fields("tipoDato"), rstObj.Fields("impLongitud"), rstObj.Fields("impY"), rstObj.Fields("impX"), Val("" & rstObj.Fields("Decimales")), 0, strMsgError
                        If strMsgError <> "" Then GoTo err
                    End If
                    rstObj.Close
                Next
            End If
            rst.Close
        End If

        Set rst = Nothing
        Set rstObj = Nothing
                               
                                       
'        Printer.Print Chr$(149)
        
        Printer.Print ""
        Printer.Print ""
                               
        Printer.EndDoc
        Exit Sub
err:
        If strMsgError = "" Then strMsgError = err.Description
        Printer.KillDoc

End Sub

Private Sub ImprimeXY(varData As Variant, strTipoDato As String, intTamanoCampo As Integer, intFila As Integer, intColu As Integer, intDecimales As Integer, intFilas As Integer, ByRef strMsgError As String)
    Dim i As Integer
    Dim strDec  As String
    Dim indFinWhile As Boolean
    Dim intFilaImp As Integer
    Dim intIndiceInicio As Integer
    
    On Error GoTo err
    Select Case strTipoDato
        Case "T"   'texto
             
             If (intFilas = 0 Or intFilas = 1) Or Len(varData) <= intTamanoCampo Then
                
                Printer.CurrentY = intFila
                Printer.CurrentX = intColu
                
                Printer.Print Left(varData, intTamanoCampo)
             Else
                indFinWhile = True
                intFilaImp = 0
                intIndiceInicio = 1
                
                Do While (indFinWhile = True)
                    If intFilaImp < intFilas Then
                        intFilaImp = intFilaImp + 1
                        
                        Printer.CurrentY = intFila
                        Printer.CurrentX = intColu
                        Printer.Print Mid(varData, intIndiceInicio, intTamanoCampo)
                        
                        intFila = intFila + 5
                        
                        intIndiceInicio = intIndiceInicio + intTamanoCampo
                    Else
                        indFinWhile = False
                    End If
                Loop
             End If
        Case "F"   'Fecha
             Printer.CurrentY = intFila
             Printer.CurrentX = intColu
             Printer.Print Left(Format(varData, "dd/mm/yyyy"), intTamanoCampo)
        Case "H"   'Hora
             Printer.CurrentY = intFila
             Printer.CurrentX = intColu
             Printer.Print Left(Format(varData, "hh:MM"), intTamanoCampo)
        Case "Y"   'Fecha y Hora
             Printer.CurrentY = intFila
             Printer.CurrentX = intColu
             Printer.Print Left(Format(varData, "dd/mm/yyyy hh:MM"), intTamanoCampo)
        Case "N"     'numerico
            Printer.CurrentY = intFila
            Printer.CurrentX = intColu
                    
            'asig. la cantidad de decimales
            For i = 1 To intDecimales
                strDec = strDec & "0"
            Next
            
            If Val(varData) >= 0 Then
                If intDecimales > 0 Then
                    Printer.Print Right((Space(intTamanoCampo) & Format(varData, "#,###,##0." & strDec)), intTamanoCampo)
                Else
                    Printer.Print Right((Space(intTamanoCampo) & Format(varData, "#,###,##0" & strDec)), intTamanoCampo)
                End If
            Else
                Printer.CurrentX = intColu - 2
                Printer.Print "(" & Right((Space(intTamanoCampo) & Format(varData, "#,###,##0." & strDec)), intTamanoCampo - 2) & ")"
            End If
        End Select
    Exit Sub
err:
     If strMsgError = "" Then strMsgError = err.Description
End Sub

