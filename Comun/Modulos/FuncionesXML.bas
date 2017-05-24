Attribute VB_Name = "FuncionesXML"
Option Explicit

Private Const SignoMayor As String = "&gt;"
Private Const SignoMenor As String = "&lt;"
Private Const SignoAmpersand  As String = "&amp;"
Private Const SignoApostrofe  As String = "&apos;"
Private Const SignoComillas  As String = "&quot;"

Public Function EjecutaProcedimiento(lstrNombreSP As String, ParamArray varValores() As Variant) As String
Dim cn As ADODB.Connection
Dim cm As ADODB.Command
Dim intI As Integer

Dim auxGrabar As Integer
Dim CanIntentos As Integer

On Error Resume Next

CanIntentos = 3

Set cn = New ADODB.Connection
Set cm = New ADODB.Command
cn.ConnectionString = gstrConnectConsulta
cn.CursorLocation = adUseClient
cn.Open
cm.CommandTimeout = 0
cm.ActiveConnection = cn
cm.CommandType = adCmdStoredProc
cm.CommandText = lstrNombreSP

For intI = 0 To UBound(varValores)
    cm.Parameters(intI + 1).Value = varValores(intI)
Next

Grabar:
auxGrabar = auxGrabar + 1
cm.Execute intI

If cn.Errors.Count > 0 Then
    If cn.Errors.Item(0).NativeError = 2627 Then
        If auxGrabar <= CanIntentos Then
            GoTo Grabar
        Else
            EjecutaProcedimiento = "Nro de intentos al grabar agotados. Error al guardar los datos"
            If cn.State = 1 Then cn.Close
            Set cn = Nothing
            Set cm = Nothing
        End If
    Else
        EjecutaProcedimiento = cn.Errors(0).Description
        If cn.State = 1 Then cn.Close
        Set cn = Nothing
        Set cm = Nothing
    End If
End If
End Function

Public Function reemplazarConstantesXML(strValor As String, _
                                        Optional BSignoMayor As Boolean = True, _
                                        Optional BSignoMenor As Boolean = True, _
                                        Optional BSignoAmpersand As Boolean = True, _
                                        Optional BSignoApostrofe As Boolean = True, _
                                        Optional BSignoComillas As Boolean = True _
                                        ) As String
    Dim strRpt As String
    If BSignoMayor = True Then strRpt = Replace(strValor, ">", SignoMayor)
    If BSignoMenor = True Then strRpt = Replace(strRpt, "<", SignoMenor)
    If BSignoAmpersand = True Then strRpt = Replace(strRpt, "&", SignoAmpersand)
    If BSignoApostrofe = True Then strRpt = Replace(strRpt, "'", SignoApostrofe)
    reemplazarConstantesXML = strRpt
End Function


Public Function CrearXML(ByVal xmlDoc As DOMDocument60, Optional intInicioValores As Boolean = False) As String
Dim strResult As String
Dim i As Integer
On Error GoTo err

strResult = "<?xml version=""1.0"" encoding=""windows-1252""?>"
strResult = strResult & "<ROOT><PARAMETERS "
'--------------------------

If intInicioValores = False Then
    For i = 3 To xmlDoc.documentElement.childNodes.Length - 1
    
        
    
        strResult = strResult & xmlDoc.documentElement.childNodes(i).nodeName & _
            "=""" & reemplazarConstantesXML(xmlDoc.documentElement.childNodes(i).Text) & """ "
    Next
Else
    For i = 0 To xmlDoc.documentElement.childNodes.Length - 1
        strResult = strResult & xmlDoc.documentElement.childNodes(i).nodeName & _
            "=""" & reemplazarConstantesXML(xmlDoc.documentElement.childNodes(i).Text) & """ "
    Next
End If
strResult = strResult & "></PARAMETERS></ROOT>"
CrearXML = strResult
On Error GoTo 0
Exit Function
err:
MsgBox "CrearXML: " & err.Description
End Function

Public Function DataProcedimiento(ByVal lstrNombreSP As String, _
                                  ByRef strMsgError As String, _
                                  ParamArray varValores() As Variant) As ADODB.Recordset
    Dim cm   As ADODB.Command
    Dim intI As Integer

    Dim result As New ADODB.Recordset
    
    On Error GoTo err

    Set cm = New ADODB.Command
    cm.CommandTimeout = 0
    cm.ActiveConnection = gstrConnectConsulta
    cm.CommandType = adCmdStoredProc
    cm.CommandText = lstrNombreSP

    For intI = 0 To UBound(varValores)
        cm.Parameters(intI + 1).Value = varValores(intI)
    Next
        Set result = cm.Execute
        'Set cm = Nothing

        On Error GoTo 0
        
       Set DataProcedimiento = result
        Exit Function
err:

        If Left(err.Description, 1) = "@" Then
            strMsgError = Mid(err.Description, 2, Len(err.Description) - 1)
        Else
            strMsgError = err.Description

        End If

        Set cm = Nothing

End Function


Public Function DataProcedimientoTipoCursor(ByVal lstrNombreSP As String, rstCursorType As CursorTypeEnum, _
                                    rstLockType As LockTypeEnum, _
                                    ByRef strMsgError As String, _
                                    ParamArray varValores() As Variant) As ADODB.Recordset
                                    
      Dim cm As ADODB.Command
      Dim intI As Integer
      Dim rst As ADODB.Recordset
10    On Error GoTo err
20    Set cm = New ADODB.Command
30    cm.CommandTimeout = 0
40    cm.ActiveConnection = gstrConnectConsulta
50    cm.CommandType = adCmdStoredProc
60    cm.CommandText = lstrNombreSP
      cm.ActiveConnection.CursorLocation = adUseClient
70    For intI = 0 To UBound(varValores)
80        cm.Parameters(intI + 1).Value = varValores(intI)
90    Next
      Set rst = New Recordset
      rst.CursorType = rstCursorType
      rst.CursorLocation = adUseClient
      rst.LockType = rstLockType
      Set rst.Source = cm
      rst.Open
100   Set DataProcedimientoTipoCursor = rst
110   Set cm = Nothing
120   On Error GoTo 0
130   Exit Function
err:
140   If strMsgError = "" Then strMsgError = "ModFuncionesSQL (DataProcedimientoTipoRst) - (" & Erl & ")" & err.Description
150   Set cm = Nothing
End Function

Public Function EjecutaProcedimientoReturn(lstrNombreSP As String, strMsgError As String, ParamArray varValores() As Variant) As String
Dim cn As ADODB.Connection
Dim cm As ADODB.Command
Dim intI As Integer
Dim auxGrabar As Integer
Dim CanIntentos As Integer

On Error Resume Next
CanIntentos = 3

Set cn = New Connection
Set cm = New ADODB.Command
cn.ConnectionString = gstrConnectConsulta
cn.CommandTimeout = 0
cn.ConnectionTimeout = 0
cn.CursorLocation = adUseClient
cn.Open
cm.CommandTimeout = 0
cm.ActiveConnection = cn
cm.CommandType = adCmdStoredProc
cm.CommandText = lstrNombreSP
For intI = 0 To UBound(varValores)
    cm.Parameters(intI + 1).Value = varValores(intI)
Next

Grabar:
auxGrabar = auxGrabar + 1
cm.Execute intI

    If cn.Errors.Count = 0 Then
        'Recorriendo el conjunto de resultados
        EjecutaProcedimientoReturn = IIf(IsNull(cm.Parameters(UBound(varValores) + 2).Value), "", cm.Parameters(UBound(varValores) + 2).Value)
        If cn.State = 1 Then cn.Close
        Set cn = Nothing
        Set cm = Nothing
    Else
        If cn.Errors.Item(0).NativeError = 2627 Then
            If auxGrabar <= CanIntentos Then
                GoTo Grabar
            Else
                strMsgError = "Nro de intentos al grabar agotados. Error al guardar los datos"
                If cn.State = 1 Then cn.Close
                Set cn = Nothing
                Set cm = Nothing
            End If
        Else
            strMsgError = cn.Errors(0).Description
            If cn.State = 1 Then cn.Close
            Set cn = Nothing
            Set cm = Nothing
        End If
    End If
End Function


Public Sub VariosValorProcedimiento(ByVal lstrNombreSP As String, ByRef strMsgError As String, ByRef ArrRpta() As Variant, ParamArray varValores() As Variant)
Dim cm As ADODB.Command
Dim intI As Integer
Dim j As Integer
Dim auxGrabar As Integer
Dim CanIntentos As Integer

On Error Resume Next

CanIntentos = 3
Set cm = New ADODB.Command
cm.CommandTimeout = 0
cm.ActiveConnection = gstrConnectConsulta
cm.CommandType = adCmdStoredProc
cm.CommandText = lstrNombreSP

For intI = 0 To UBound(varValores)
    cm.Parameters(intI + 1).Value = varValores(intI)
Next

Grabar:
auxGrabar = auxGrabar + 1
cm.Execute

    If cm.ActiveConnection.Errors.Count = 0 Then
        'Recorriendo el conjunto de resultados
        For j = 0 To cm.Parameters.Count - (intI + 2)
            If IsNull(cm.Parameters(intI + j + 1)) = False Then
                ArrRpta(j) = Trim(CStr(cm.Parameters(intI + j + 1)))
            Else
                ArrRpta(j) = ""
    
            End If
        Next
        Set cm = Nothing
    Else
        If cm.ActiveConnection.Errors.Item(0).NativeError = 2627 Then
            If auxGrabar <= CanIntentos Then
                GoTo Grabar
            Else
                strMsgError = "Nro de intentos al grabar agotados. Error al guardar los datos"
                Set cm = Nothing
            End If
        Else
            strMsgError = err.Description
            Set cm = Nothing
        End If
    End If
End Sub


Public Sub ConfGrid(g As dxDBGrid, indMod As Boolean, Optional mostrarFooter As Boolean, Optional mostrarGroupPanel As Boolean, Optional mostrarBandas As Boolean)
     With g.Options
        '***
        If indMod Then
            .Set (egoEditing)
            .Set (egoCanDelete)
            .Set (egoCanInsert)
            '.Set (egoCanAppend)
        End If
        If mostrarBandas Then .Set (egoShowBands)
        .Set (egoTabs)
        .Set (egoTabThrough)
        .Set (egoImmediateEditor)
        .Set (egoShowIndicator)
        .Set (egoCanNavigation)
        .Set (egoHorzThrough)
        .Set (egoVertThrough)
        '.Set (egoAutoWidth)
        .Set (egoEnterShowEditor)
        .Set (egoEnterThrough)
        .Set (egoShowButtonAlways)
        .Set (egoColumnSizing)
        .Set (egoColumnMoving)
        .Set (egoTabThrough)
        .Set (egoConfirmDelete)
        .Set (egoCancelOnExit)
        .Set (egoLoadAllRecords)
        .Set (egoShowHourGlass)
        .Set (egoUseBookmarks)
        .Set (egoUseLocate)
        .Set (egoAutoCalcPreviewLines)
        .Set (egoBandSizing)
        .Set (egoBandMoving)
        .Set (egoDragScroll)
        .Set (egoAutoSort)
        .Set (egoExpandOnDblClick)
        .Set (egoNameCaseInsensitive)
        If mostrarFooter Then .Set (egoShowFooter)
        .Set (egoShowGrid)
        .Set (egoShowButtons)
        .Set (egoShowHeader)
        .Set (egoShowPreviewGrid)
        .Set (egoShowBorder)
        '.Set (egoShowRowFooter)
        '''''.Set (egoDynamicLoad)
        
        If mostrarGroupPanel Then .Set (egoShowGroupPanel)
        .Set (egoEnableNodeDragging)
        .Set (egoDragCollapse)
        .Set (egoDragExpand)
        .Set (egoDragScroll)
        .Set (egoEnableNodeDragging)
    End With
End Sub

Public Sub mostrarDatosGridSQL(g As dxDBGrid, r As Recordset, ByRef strMsgError As String, Optional strKeyField As String = "Item")
On Error GoTo err
    With g
        .DefaultFields = False
    
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Active = False
           
        Set .DataSource = r
    
        .Dataset.Active = True
        .KeyField = strKeyField
        .Dataset.Edit
        .Dataset.Post
        
    End With
Exit Sub
err:
If strMsgError = "" Then strMsgError = err.Description
End Sub


Public Function CrearFormXML(frmFormulario As Form, Optional NomFrame As Object, Optional X As DOMDocument60) As DOMDocument60
Dim c As Control, C1 As Control
Dim F As Frame
Dim XMLTemp As DOMDocument60
On Error GoTo CrearFormXML_Error
Set XMLTemp = New DOMDocument60

'CreaParametrosDeValidacion XMLTemp
'********************************************
If TypeName(X) = "Nothing" Then 'CREA UN NUEVO DOCUMENTO XML
    'CreaParametrosDeValidacion XMLTemp
Else 'INSERTA LOS VALORES EN UN DOCUMENTO XML EXISTENTE
    Set XMLTemp = X
End If
'*******************************************
For Each c In frmFormulario
    If TypeName(NomFrame) <> "Nothing" Then
        If c.Tag <> "" Then
            If TypeName(c) = "TextBox" Or TypeName(c) = "ComboBox" Or TypeName(c) = "CheckBox" Or TypeName(c) = "DTPicker" Or TypeName(c) = "TAMTextBox" Or TypeName(c) = "Label" Then
                If c.Container.Name = NomFrame.Name Then
                    Select Case LCase(TypeName(c))
                        Case Is = "textbox", Is = "combobox"
                            CreaParametrosXML XMLTemp, c.Tag, , c.Text
                        Case Is = "optionbutton", Is = "tamtextbox"
                            CreaParametrosXML XMLTemp, c.Tag, , c.Value
                        Case Is = "dtpicker"
                            CreaParametrosXML XMLTemp, c.Tag, , Format(c.Value, "mm/dd/yyyy")  ' JCB
                        Case Is = "checkbox"
                            If c.Value = 0 Or c.Value = 1 Then CreaParametrosXML XMLTemp, c.Tag, , c.Value
                        Case Is = "label"
                            If IsNumeric(c.Caption) = True Then
                                CreaParametrosXML XMLTemp, c.Tag, , CDbl(c.Caption)
                            Else
                                CreaParametrosXML XMLTemp, c.Tag, , c.Caption
                            End If
                        Case Is = "frame"
                            Set F = c
                            For Each C1 In frmFormulario
                                If C1.Tag <> "" Then
                                    If TypeName(C1) = "TextBox" Or TypeName(C1) = "ComboBox" Or TypeName(C1) = "CheckBox" Or TypeName(C1) = "DTPicker" Or TypeName(c) = "TAMTextBox" Or TypeName(c) = "Label" Then
                                        If C1.Container = F Then
                                            Select Case LCase(TypeName(C1))
                                                Case Is = "textbox", Is = "combobox"
                                                    CreaParametrosXML XMLTemp, c.Tag, , c.Text
                                                Case Is = "optionbutton", Is = "tamtextbox"
                                                    CreaParametrosXML XMLTemp, c.Tag, , c.Value
                                                Case Is = "dtpicker"
                                                    CreaParametrosXML XMLTemp, c.Tag, , Format(c.Value, "mm/dd/yyyy")   'JCB
                                                Case Is = "checkbox"
                                                    If c.Value = 0 Or c.Value = 1 Then CreaParametrosXML XMLTemp, c.Tag, , c.Value
                                                Case Is = "label"
                                                    If IsNumeric(c.Caption) = True Then
                                                        CreaParametrosXML XMLTemp, c.Tag, , CDbl(c.Caption)
                                                    Else
                                                        CreaParametrosXML XMLTemp, c.Tag, , c.Caption
                                                    End If
                                            End Select
                                        End If
                                    End If
                                End If
                            Next
                    End Select
                End If
            End If
        End If
    Else
        If c.Tag <> "" Then
            Select Case LCase(TypeName(c))
                Case Is = "textbox", Is = "combobox"
                    CreaParametrosXML XMLTemp, c.Tag, , c.Text
                Case Is = "optionbutton", Is = "tamtextbox"
                    CreaParametrosXML XMLTemp, c.Tag, , c.Value
                Case Is = "dtpicker"
                    CreaParametrosXML XMLTemp, c.Tag, , Format(c.Value, "mm/dd/yyyy")   'JCB
                Case Is = "checkbox"
                    If c.Value = 0 Or c.Value = 1 Then CreaParametrosXML XMLTemp, c.Tag, , c.Value
                Case Is = "label"
                    If IsNumeric(c.Caption) = True Then
                        CreaParametrosXML XMLTemp, c.Tag, , CDbl(c.Caption)
                    Else
                        CreaParametrosXML XMLTemp, c.Tag, , c.Caption
                    End If
            End Select
        End If
    End If
Next
Set CrearFormXML = XMLTemp
Set XMLTemp = Nothing
On Error GoTo 0
Exit Function
CrearFormXML_Error:
Set XMLTemp = Nothing
Set CrearFormXML = Nothing
MsgBox "Error " & err.Number & " (" & err.Description & ") CrearFormXML"
End Function

Public Sub XMLDetalleGrid(ByRef objXML As DOMDocument60, ByVal strNomEntidad As String, ByVal g As dxDBGrid, ByVal strNomCampos As String, ByRef strMsgError As String, Optional strCampoCond As String, Optional strDatoCond As String, Optional optFilaInicio As Integer = 1)
    'Dim objXML As MSXML2.DOMDocument
    Dim objElem As MSXML2.IXMLDOMElement
    Dim objParent As MSXML2.IXMLDOMElement
    Dim lngPos As Long, lngParent As Long
    Dim i As Integer, j As Integer, aux As Integer
    Dim lblnSuccess As Boolean
    Dim NomCampos() As String, ArrayCols() As String
    Dim indCumpleCondicion As Boolean
    
    On Error GoTo ErrCreaXMLDetalleGrid
    
    NomCampos = Split(strNomCampos, ",")
    'ArrayCols = Split(strNumColumnas, ",")
    ' iniciando el documento xml
    If objXML Is Nothing Then
        Set objXML = New MSXML2.DOMDocument60
        Set objXML.documentElement = objXML.createElement("ROOT")
    'Else
    '    lblnSuccess = objXML.loadXML(xmlDoc)
    End If
    Set objParent = objXML.documentElement
    'Recorriendo todas las filas de una rejilla
    If g.Count > 0 Then
        g.Dataset.DisableControls
                
        g.Dataset.First
        Do While Not g.Dataset.EOF
            indCumpleCondicion = True
            If strCampoCond <> "" Then
                    If g.Columns.ColumnByFieldName(strCampoCond).Value <> strDatoCond Then indCumpleCondicion = False
            End If
            
            If indCumpleCondicion Then
                Set objElem = objParent.appendChild(objXML.createElement(strNomEntidad))
                For j = 0 To UBound(NomCampos)
                    'añadiendo los atributos, solo para las columnas especificadas
                    objElem.setAttribute NomCampos(j), "" & g.Columns.ColumnByFieldName(NomCampos(j)).Value
                Next
            End If
            
            g.Dataset.Next
        Loop
        g.Dataset.EnableControls
    End If
    'Set objXML = Nothing
    Set objElem = Nothing
    Set objParent = Nothing
    Exit Sub
ErrCreaXMLDetalleGrid:
    If strMsgError = "" Then strMsgError = "(XMLDetalleGrid) - " & err.Description
End Sub
Public Sub XMLADORecordset(ByRef objXML As DOMDocument60, ByVal strNomDocumento As String, ByVal strNomEntidad As String, ByVal adoRecordset As ADODB.Recordset, ByRef strMsgError As String, Optional strNomCampos As String = "", Optional strCampoCond As String, Optional strDatoCond As String, Optional optFilaInicio As Integer = 1)
    'Dim objXML As MSXML2.DOMDocument
    Dim objElem As MSXML2.IXMLDOMElement
    Dim objParent As MSXML2.IXMLDOMElement
    Dim lngPos As Long, lngParent As Long
    Dim i As Integer, j As Integer, aux As Integer
    Dim lblnSuccess As Boolean
    Dim NomCampos() As String, ArrayCols() As String
    Dim indCumpleCondicion As Boolean
    Dim adoField As ADODB.Field
    Dim n As Integer
    
    On Error GoTo ErrCreaXMLADORecordset
    
    NomCampos = Split(strNomCampos, ",")
    'ArrayCols = Split(strNumColumnas, ",")
    ' iniciando el documento xml
    If objXML Is Nothing Then
        Set objXML = New MSXML2.DOMDocument60
        Set objXML.documentElement = objXML.createElement(strNomDocumento)
    'Else
    '    lblnSuccess = objXML.loadXML(xmlDoc)
    End If
    Set objParent = objXML.documentElement
    
    n = 0
    
    'Recorriendo todas las filas de una rejilla
    If adoRecordset.RecordCount > 0 Then
                
        If UBound(NomCampos) = -1 Then
            For Each adoField In adoRecordset.Fields
                'Set adoField = ADODB.Field
                ReDim Preserve NomCampos(n)
                NomCampos(n) = adoField.Name
                n = n + 1
            Next
        End If
                
        adoRecordset.MoveFirst
        Do While Not adoRecordset.EOF
            indCumpleCondicion = True
            If strCampoCond <> "" Then
                If adoRecordset.Fields(strCampoCond).Value <> strDatoCond Then indCumpleCondicion = False
            End If
            
            If indCumpleCondicion Then
                Set objElem = objParent.appendChild(objXML.createElement(strNomEntidad))
                For j = 0 To UBound(NomCampos)
                    'añadiendo los atributos, solo para las columnas especificadas
                    objElem.setAttribute NomCampos(j), "" & adoRecordset.Fields(NomCampos(j)).Value
                Next
            End If
            
            adoRecordset.MoveNext
        Loop
    End If
    'Set objXML = Nothing
    Set objElem = Nothing
    Set objParent = Nothing
    Exit Sub
ErrCreaXMLADORecordset:
    If strMsgError = "" Then strMsgError = "(XMLADORecordset) - " & err.Description
End Sub

'Public Function CreaParametrosXML(xmlDoc As DOMDocument60, strGlsElemento As String, Optional strGlsSubElemento As String, Optional strGlsValor As String) As Boolean
'On Error GoTo ErrCreaParametrosXML
'Dim lblnSuccess As Boolean
'Dim Element     As IXMLDOMElement, SubElement As IXMLDOMElement
'Dim newNode     As IXMLDOMNode
'If Trim(strGlsElemento) = "" Then
'    CreaParametrosXML = False
'    Exit Function
'End If
'lblnSuccess = True
'If xmlDoc Is Nothing Then
'    Set xmlDoc = New DOMDocument60
'End If
'If xmlDoc.childNodes.length <= 0 Then
'    lblnSuccess = xmlDoc.loadXML("<parameters/>")
'End If
'If lblnSuccess = True Then
'    Set Element = xmlDoc.documentElement.selectSingleNode(strGlsElemento)
'    If Element Is Nothing Then
'        Set newNode = xmlDoc.createNode(NODE_ELEMENT, strGlsElemento, "")
'        xmlDoc.documentElement.appendChild newNode
'        Set Element = xmlDoc.documentElement.selectSingleNode(strGlsElemento)
'    End If
'    If Trim(strGlsSubElemento) <> "" Then
'        Set SubElement = Element.selectSingleNode(strGlsSubElemento)
'        If SubElement Is Nothing Then
'            Set newNode = xmlDoc.createNode(NODE_ELEMENT, strGlsSubElemento, "")
'            Element.appendChild newNode
'        End If
'    End If
'    newNode.Text = strGlsValor
'    Set newNode = Nothing
'    Set Element = Nothing
'    Set SubElement = Nothing
'    CreaParametrosXML = True
'Else
'    CreaParametrosXML = False
'End If
'Exit Function
'ErrCreaParametrosXML:
'CreaParametrosXML = False
'End Function

Public Sub mostrarDatosFormSQL(F As Form, r As Recordset, ByRef strMsgError As String, Optional ByVal objContenedor As Object)
Dim Ctrl As Control
On Error GoTo err
If TypeName(objContenedor) = "Nothing" Then
    For Each Ctrl In F
        If Ctrl.Tag <> "" Then
            If TypeOf Ctrl Is DTPicker Then
                Ctrl.Value = r(Ctrl.Tag).Value
            ElseIf TypeOf Ctrl Is OptionButton Then
                If r(Ctrl.Tag).Value = "Verdadero" Or _
                r(Ctrl.Tag).Value = "True" Then
                    Ctrl.Value = True
                Else
                    Ctrl.Value = False
                End If
            ElseIf TypeOf Ctrl Is CheckBox Then
                If CBool(r(Ctrl.Tag).Value) = True Then
                    Ctrl.Value = 1
                Else
                    Ctrl.Value = 0
                End If
            ElseIf TypeOf Ctrl Is Label Then
                Ctrl.Caption = r(Ctrl.Tag).Value
            Else
                Ctrl.Text = "" & r(Ctrl.Tag).Value
            End If
        End If
    Next
Else
    For Each Ctrl In F
            If Ctrl.Tag <> "" Then
                If Ctrl.Container.Name = objContenedor.Name Then
                    If TypeOf Ctrl Is DTPicker Then
                        Ctrl.Value = r(Ctrl.Tag).Value
                    ElseIf TypeOf Ctrl Is OptionButton Then
                        If r(Ctrl.Tag).Value = "Verdadero" Or _
                        r(Ctrl.Tag).Value = "True" Then
                            Ctrl.Value = True
                        Else
                            Ctrl.Value = False
                        End If
                    ElseIf TypeOf Ctrl Is CheckBox Then
                        Ctrl.Value = Val(r(Ctrl.Tag).Value)
                    ElseIf TypeOf Ctrl Is Label Then
                        Ctrl.Caption = r(Ctrl.Tag).Value
                    Else
                        Ctrl.Text = "" & r(Ctrl.Tag).Value
                    End If
                End If
            End If
'        End If
    Next
End If
Exit Sub
err:
If strMsgError = "" Then strMsgError = err.Description
End Sub

Public Sub mostrarDatosGridRS(g As dxDBGrid, r As Recordset, ByRef strMsgError As String)
Dim i As Integer
On Error GoTo err

Do While Not r.EOF
    For i = 0 To r.Fields.Count - 1
        g.Dataset.Edit
        g.Columns.ColumnByName(r.Fields(i).Name).Value = r.Fields(i)
        g.Dataset.Post
    Next
    
    g.Dataset.Insert
    
    r.MoveNext
Loop

g.Dataset.Delete

Exit Sub
err:
If strMsgError = "" Then strMsgError = err.Description
End Sub

Public Function traerCampo(tabla As String, campoTraer As String, campoComp As String, Valor As String, Optional condicion As String) As String
    Dim rs              As New ADODB.Recordset
    Dim csql            As String
    
    If Trim(campoComp) <> "" Then
        csql = "Select " & campoTraer + " From " + tabla + " where " + campoComp + " = '" & Valor & "'"
    
        If condicion <> "" Then
            If UCase(Left(Trim(condicion), 2)) = "OR" Or UCase(Left(Trim(condicion), 3)) = "AND" Then
                csql = csql + " " + condicion
            Else
                csql = csql + " and " & condicion
            End If
        End If
    
    Else
        csql = "Select " & campoTraer + " From " + tabla
    
        If condicion <> "" Then
            csql = csql & " where "
            If UCase(Left(Trim(condicion), 2)) = "OR" Or UCase(Left(Trim(condicion), 3)) = "AND" Then
                csql = csql + " " + condicion
            Else
                csql = csql & condicion
            End If
        End If
    
    
    End If
    
'    If condicion <> "" Then
'        If UCase(Left(Trim(condicion), 2)) = "OR" Or UCase(Left(Trim(condicion), 3)) = "AND" Then
'            csql = csql + " " + condicion
'        Else
'            csql = csql + " and " & condicion
'        End If
'    End If
    
    rs.Open csql, gstrConnectConsulta, adOpenForwardOnly, adLockReadOnly
    
    traerCampo = ""
    If Not rs.EOF Then
        If Not IsNull(rs.Fields(0)) Then traerCampo = (rs.Fields(0))
    End If
    rs.Close
    Set rs = Nothing

End Function

Public Function CrearXMLDetalle(ByVal xmlDoc As DOMDocument60) As String
Dim strResult As String
CrearXMLDetalle = "<?xml version=""1.0"" encoding=""windows-1252""?>" & xmlDoc.xml
On Error GoTo 0
Exit Function
CrearXML_Error:
MsgBox "Error " & err.Number & " (" & err.Description & ") CrearXMLDetalle"
End Function


Public Sub MostrarAyudaTablaTexto(strTabla As String, strCampoCodigo As String, _
strCampoGlosa As String, strParCodigo As String, strParDescripcion As String, Optional strCondicional As String = "")

Dim strCODIGO As String, strDescripcion As String

'If frmAyuda.Execute(strTabla, strCampoCodigo, strCampoGlosa, strCodigo, strDescripcion, strCondicional) Then
'    strParCodigo = strCodigo
'    strParDescripcion = strDescripcion
'Else
    
'End If

End Sub

Public Sub XMLDetalleListView(ByRef objXML As DOMDocument60, ByVal strNomEntidad As String, ByVal g As ListView, ByVal strNumColumnas As String, ByVal strNomColumnas As String, ByRef strMsgError As String)
    Dim objElem As MSXML2.IXMLDOMElement
    Dim objParent As MSXML2.IXMLDOMElement
    Dim i As Integer, j As Integer
    Dim lblnSuccess As Boolean
    Dim NumColumnas() As String, NomColumnas() As String
    
    On Error GoTo err
    
    NumColumnas = Split(strNumColumnas, ",")
    NomColumnas = Split(strNomColumnas, ",")
    ' iniciando el documento xml
    If objXML Is Nothing Then
        Set objXML = New MSXML2.DOMDocument60
        Set objXML.documentElement = objXML.createElement("ROOT")
    'Else
    '    lblnSuccess = objXML.loadXML(xmlDoc)
    End If
    Set objParent = objXML.documentElement
    'Recorriendo todas las filas de una rejilla
    For i = 1 To g.ListItems.Count
        Set objElem = objParent.appendChild(objXML.createElement(strNomEntidad))
        For j = 0 To UBound(NumColumnas)
            'añadiendo los atributos, solo para las columnas especificadas
            If NumColumnas(j) = 0 Then
                objElem.setAttribute NomColumnas(j), "" & g.ListItems(i).Text
            Else
                objElem.setAttribute NomColumnas(j), "" & g.ListItems(i).SubItems(NumColumnas(j))
            End If
        Next
    Next
    
    Set objElem = Nothing
    Set objParent = Nothing
    Exit Sub
err:
    If strMsgError = "" Then strMsgError = err.Description
End Sub
Public Sub XMLADORecordsetD(ByVal strXMLDoc As String, ByVal strNomDocumento As String, ByRef adoRecordset As ADODB.Recordset, ByRef strMsgError As String)
    Dim objXML As New MSXML2.DOMDocument60
    Dim objElem As MSXML2.IXMLDOMElement
    Dim objSub As MSXML2.IXMLDOMElement
    
    Dim i As Integer
    Dim lblnSuccess As Boolean
    
    'XMLADORecordsetD: Esta rutina toma un XML y lo convierte en Recordset
    
    On Error GoTo ErrCreaXMLADORecordsetD
    
    lblnSuccess = objXML.loadXML(strXMLDoc)

    Set objElem = objXML.selectSingleNode("//" & strNomDocumento)

    For Each objSub In objElem.childNodes
        Debug.Print objSub.nodeName
        
        If objSub.Attributes.Length > 0 Then
            adoRecordset.AddNew
            For i = 0 To objSub.Attributes.Length - 1
                adoRecordset.Fields(objSub.Attributes(i).nodeName) = objSub.Attributes(i).nodeValue
                Debug.Print objSub.Attributes(i).nodeName & " - " & objSub.Attributes(i).nodeValue
            Next i
            adoRecordset.Update
        End If
    
    Next
    
    Set objXML = Nothing
    Set objElem = Nothing
    Set objSub = Nothing
    Exit Sub
ErrCreaXMLADORecordsetD:
    If strMsgError = "" Then strMsgError = "(XMLADORecordsetD) - " & err.Description
End Sub
Public Function CreaParametrosXML(xmlDoc As DOMDocument60, strGlsElemento As String, Optional strGlsSubElemento As String, Optional strGlsValor As String) As Boolean
On Error GoTo ErrCreaParametrosXML
Dim lblnSuccess As Boolean
Dim Element     As IXMLDOMElement, SubElement As IXMLDOMElement
Dim newNode     As IXMLDOMNode
If Trim(strGlsElemento) = "" Then
    CreaParametrosXML = False
    Exit Function
End If
lblnSuccess = True
If xmlDoc Is Nothing Then
    Set xmlDoc = New DOMDocument60
End If
If xmlDoc.childNodes.Length <= 0 Then
    lblnSuccess = xmlDoc.loadXML("<parameters/>")
End If
If lblnSuccess = True Then
    Set Element = xmlDoc.documentElement.selectSingleNode(strGlsElemento)
    If Element Is Nothing Then
        Set newNode = xmlDoc.createNode(NODE_ELEMENT, strGlsElemento, "")
        xmlDoc.documentElement.appendChild newNode
        Set Element = xmlDoc.documentElement.selectSingleNode(strGlsElemento)
    End If
    If Trim(strGlsSubElemento) <> "" Then
        Set SubElement = Element.selectSingleNode(strGlsSubElemento)
        If SubElement Is Nothing Then
            Set newNode = xmlDoc.createNode(NODE_ELEMENT, strGlsSubElemento, "")
            Element.appendChild newNode
        End If
    End If
    newNode.Text = strGlsValor
    Set newNode = Nothing
    Set Element = Nothing
    Set SubElement = Nothing
    CreaParametrosXML = True
Else
    CreaParametrosXML = False
End If
Exit Function
ErrCreaParametrosXML:
CreaParametrosXML = False
End Function


Public Sub XMLFormularioControlTag(ByRef objXML As DOMDocument60, ByVal strNomDocumento As String, ByVal strNomEntidad As String, F As Form, ByRef strMsgError As String, Optional ByVal objContenedor As Object)
    Dim Ctrl As Control
    Dim objElem As MSXML2.IXMLDOMElement
    Dim objParent As MSXML2.IXMLDOMElement
    
    On Error GoTo err

    If objXML Is Nothing Then
        Set objXML = New MSXML2.DOMDocument60
        Set objXML.documentElement = objXML.createElement(strNomDocumento)
    End If
    
    Set objParent = objXML.documentElement

    'Set objElem = objParent.appendChild(objXML.createElement(strNomEntidad))
        
    If TypeName(objContenedor) = "Nothing" Then
        For Each Ctrl In F
            If Ctrl.Tag <> "" Then
                If TypeOf Ctrl Is DTPicker Then
                    objParent.setAttribute Ctrl.Tag, "" & Ctrl.Value
                ElseIf TypeOf Ctrl Is OptionButton Then
                    objParent.setAttribute Ctrl.Tag, "" & Ctrl.Value
                ElseIf TypeOf Ctrl Is CheckBox Then
                    objParent.setAttribute Ctrl.Tag, "" & Ctrl.Value
                ElseIf TypeOf Ctrl Is ComboBox Then
                    objParent.setAttribute Ctrl.Tag, "" & Ctrl.ItemData(Ctrl.ListIndex)
                ElseIf TypeOf Ctrl Is Label Then
                    objParent.setAttribute Ctrl.Tag, "" & Ctrl.Caption
                Else
                    objParent.setAttribute Ctrl.Tag, "" & Ctrl.Text
                End If
            End If
        Next
    Else
        For Each Ctrl In F
            If Ctrl.Tag <> "" Then
                If Ctrl.Container.Name = objContenedor.Name Then
                    If TypeOf Ctrl Is DTPicker Then
                        objParent.setAttribute Ctrl.Tag, "" & Ctrl.Value
                    ElseIf TypeOf Ctrl Is OptionButton Then
                        objParent.setAttribute Ctrl.Tag, "" & Ctrl.Value
                    ElseIf TypeOf Ctrl Is CheckBox Then
                        objParent.setAttribute Ctrl.Tag, "" & Ctrl.Value
                    ElseIf TypeOf Ctrl Is ComboBox Then
                        objParent.setAttribute Ctrl.Tag, "" & Ctrl.ItemData(Ctrl.ListIndex)
                    ElseIf TypeOf Ctrl Is Label Then
                        objParent.setAttribute Ctrl.Tag, "" & Ctrl.Caption
                    Else
                        objParent.setAttribute Ctrl.Tag, "" & Ctrl.Text
                    End If
                End If
            End If
        Next
    End If
    
    Exit Sub
err:
    If strMsgError = "" Then strMsgError = err.Description
End Sub

