VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmKardexInversion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta del Kardex"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6540
   ScaleWidth      =   10185
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   8580
      TabIndex        =   9
      Top             =   5700
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
      Caption0        =   "&Salir"
      Tag0            =   "9"
      ToolTipText0    =   "Salir"
      UserControlWidth=   1200
   End
   Begin VB.Frame fraKardex 
      Height          =   5535
      Left            =   30
      TabIndex        =   5
      Top             =   0
      Width           =   10095
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmKardexInversion.frx":0000
         Height          =   2895
         Left            =   360
         OleObjectBlob   =   "frmKardexInversion.frx":001A
         TabIndex        =   4
         Top             =   2280
         Width           =   9375
      End
      Begin VB.CheckBox chkUltimo 
         Caption         =   "Ultimo Movimiento"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   1800
         Width           =   2175
      End
      Begin VB.ComboBox cboFondo 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   6135
      End
      Begin VB.ComboBox cboTipoInstrumento 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   840
         Width           =   6135
      End
      Begin VB.ComboBox cboTitulo 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1320
         Width           =   6135
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         Caption         =   "Fondo"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   8
         Top             =   405
         Width           =   450
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         Caption         =   "Instrumento"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   7
         Top             =   915
         Width           =   825
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         Caption         =   "Título"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   6
         Top             =   1320
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmKardexInversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Ficha de Consulta del Kardex de Cartera"
Option Explicit

Dim arrFondo()      As String, arrTipoInstrumento()     As String
Dim arrTitulo()     As String
Dim strCodFondo     As String, strCodTipoInstrumento    As String
Dim strCodTitulo    As String, strSQL                   As String
Dim strEstado       As String, strCodMoneda             As String
Dim adoConsulta     As ADODB.Recordset
Dim indSortAsc      As Boolean, indSortDesc             As Boolean

Public Sub Adicionar()

End Sub

Public Sub Eliminar()

End Sub

Public Sub Grabar()

End Sub


Public Sub Imprimir()

End Sub

Public Sub Modificar()

End Sub
Public Sub SubImprimir(Index As Integer)

    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    Dim strFechaDesde           As String, strFechaHasta        As String
    Dim strSeleccionRegistro    As String
        
    Select Case Index
        Case 1
            gstrNameRepo = "InversionKardex2"
            
            strSeleccionRegistro = "{InversionKardex.FechaMovimiento} IN 'Fch1' TO 'Fch2'"
            gstrSelFrml = strSeleccionRegistro
            frmRangoFecha.Show vbModal
                        
            If gstrSelFrml <> "0" Then
                Set frmReporte = New frmVisorReporte
    
                ReDim aReportParamS(6)
                ReDim aReportParamFn(5)
                ReDim aReportParamF(5)
                            
                aReportParamFn(0) = "Usuario"
                aReportParamFn(1) = "FechaDesde"
                aReportParamFn(2) = "FechaHasta"
                aReportParamFn(3) = "Hora"
                aReportParamFn(4) = "Fondo"
                aReportParamFn(5) = "NombreEmpresa"
                            
                aReportParamF(0) = gstrLogin
                aReportParamF(1) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10)
                aReportParamF(2) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10)
                aReportParamF(3) = Format(Time(), "hh:mm:ss")
                aReportParamF(4) = Trim(cboFondo.Text)
                aReportParamF(5) = gstrNombreEmpresa & Space(1)
                            
                aReportParamS(0) = strCodFondo
                aReportParamS(1) = gstrCodAdministradora
                aReportParamS(2) = Convertyyyymmdd(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10))
                aReportParamS(3) = Convertyyyymmdd(DateAdd("d", 1, CVDate(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10))))
                aReportParamS(4) = strCodMoneda
                aReportParamS(5) = strCodTitulo
                aReportParamS(6) = Valor_Caracter
                
                If chkUltimo.Value Then aReportParamS(6) = Valor_Indicador
                
            End If
            
    
        Case 2
            gstrNameRepo = "InversionCarteraCliente"
            
            strSeleccionRegistro = "{InversionCarteraCliente.FechaMovimiento} IN 'Fch1' TO 'Fch2'"
            gstrSelFrml = strSeleccionRegistro
            'frmRangoFecha.lblRango(0).Visible = False
            'frmRangoFecha.dtpFechaInicial.Visible = False
            'esto se cambio para probar
            frmRangoFecha.lblRango(1).Visible = False
            frmRangoFecha.dtpFechaFinal.Visible = False
            frmRangoFecha.Show vbModal
            'gstrFchAl = frmRangoFecha.dtpFechaInicial
                        
            If gstrSelFrml <> "0" Then
                Set frmReporte = New frmVisorReporte
    
                ReDim aReportParamS(2)
                ReDim aReportParamFn(5)
                ReDim aReportParamF(5)
                            
                aReportParamFn(0) = "Usuario"
                aReportParamFn(1) = "FechaDesde"
                aReportParamFn(2) = "FechaHasta"
                aReportParamFn(3) = "Hora"
                aReportParamFn(4) = "Fondo"
                aReportParamFn(5) = "NombreEmpresa"
                            
                aReportParamF(0) = gstrLogin
                aReportParamF(1) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10)
                aReportParamF(2) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10)
                aReportParamF(3) = Format(Time(), "hh:mm:ss")
                aReportParamF(4) = Trim(cboFondo.Text)
                aReportParamF(5) = gstrNombreEmpresa & Space(1)
                            
                aReportParamS(0) = strCodFondo
                aReportParamS(1) = gstrCodAdministradora
                'aReportParamS(2) = Convertyyyymmdd(gstrFchAl)
                'esto se modifico para probar
                aReportParamS(2) = Convertyyyymmdd(gstrFchDel)
                'aReportParamS(3) = strCodParticipe
                'aReportParamS(3) = strCodTipoInstrumento 'ACR
                
            End If
    
    
    End Select

    If gstrSelFrml = "0" Then Exit Sub
    
    gstrSelFrml = Valor_Caracter
    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"

    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())

    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
    frmReporte.Show vbModal

    Set frmReporte = Nothing

    Screen.MousePointer = vbNormal
    
End Sub
'Public Sub SubImprimir(Index As Integer)
'
'    Dim frmReporte              As frmVisorReporte
'    Dim aReportParamS(), aReportParamF(), aReportParamFn()
'    Dim strFechaDesde           As String, strFechaHasta        As String
'    Dim strSeleccionRegistro    As String
'
'    Select Case Index
'        Case 1
'            gstrNameRepo = "InversionKardex"
'
'            strSeleccionRegistro = "{InversionKardex.FechaMovimiento} IN 'Fch1' TO 'Fch2'"
'            gstrSelFrml = strSeleccionRegistro
'            frmRangoFecha.Show vbModal
'
'            If gstrSelFrml <> "0" Then
'                Set frmReporte = New frmVisorReporte
'
'                ReDim aReportParamS(6)
'                ReDim aReportParamFn(5)
'                ReDim aReportParamF(5)
'
'                aReportParamFn(0) = "Usuario"
'                aReportParamFn(1) = "FechaDesde"
'                aReportParamFn(2) = "FechaHasta"
'                aReportParamFn(3) = "Hora"
'                aReportParamFn(4) = "Fondo"
'                aReportParamFn(5) = "NombreEmpresa"
'
'                aReportParamF(0) = gstrLogin
'                aReportParamF(1) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10)
'                aReportParamF(2) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10)
'                aReportParamF(3) = Format(Time(), "hh:mm:ss")
'                aReportParamF(4) = Trim(cboFondo.Text)
'                aReportParamF(5) = gstrNombreEmpresa & Space(1)
'
'                aReportParamS(0) = strCodFondo
'                aReportParamS(1) = gstrCodAdministradora
'                aReportParamS(2) = Convertyyyymmdd(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10))
'                aReportParamS(3) = Convertyyyymmdd(DateAdd("d", 1, CVDate(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10))))
'                aReportParamS(4) = strCodMoneda
'                aReportParamS(5) = strCodTitulo
'                aReportParamS(6) = Valor_Caracter
'                If chkUltimo.Value Then aReportParamS(6) = Valor_Indicador
'
'            End If
'
'
'        Case 2
'            gstrNameRepo = "InversionCarteraCliente"
'
'            strSeleccionRegistro = "{InversionCarteraCliente.FechaMovimiento} IN 'Fch1' TO 'Fch2'"
'            gstrSelFrml = strSeleccionRegistro
'            'frmRangoFecha.lblRango(0).Visible = False
'            'frmRangoFecha.dtpFechaInicial.Visible = False
'            'esto se cambio para probar
'            frmRangoFecha.lblRango(1).Visible = False
'            frmRangoFecha.dtpFechaFinal.Visible = False
'            frmRangoFecha.Show vbModal
'            'gstrFchAl = frmRangoFecha.dtpFechaInicial
'
'            If gstrSelFrml <> "0" Then
'                Set frmReporte = New frmVisorReporte
'
'                ReDim aReportParamS(2)
'                ReDim aReportParamFn(5)
'                ReDim aReportParamF(5)
'
'                aReportParamFn(0) = "Usuario"
'                aReportParamFn(1) = "FechaDesde"
'                aReportParamFn(2) = "FechaHasta"
'                aReportParamFn(3) = "Hora"
'                aReportParamFn(4) = "Fondo"
'                aReportParamFn(5) = "NombreEmpresa"
'
'                aReportParamF(0) = gstrLogin
'                aReportParamF(1) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10)
'                aReportParamF(2) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10)
'                aReportParamF(3) = Format(Time(), "hh:mm:ss")
'                aReportParamF(4) = Trim(cboFondo.Text)
'                aReportParamF(5) = gstrNombreEmpresa & Space(1)
'
'                aReportParamS(0) = strCodFondo
'                aReportParamS(1) = gstrCodAdministradora
'                'aReportParamS(2) = Convertyyyymmdd(gstrFchAl)
'                'esto se modifico para probar
'                aReportParamS(2) = Convertyyyymmdd(gstrFchDel)
'                'aReportParamS(3) = strCodParticipe
'
'            End If
'
'
'    End Select
'
'    If gstrSelFrml = "0" Then Exit Sub
'
'    gstrSelFrml = Valor_Caracter
'    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"
'
'    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())
'
'    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
'    frmReporte.Show vbModal
'
'    Set frmReporte = Nothing
'
'    Screen.MousePointer = vbNormal
'
'End Sub
Public Sub Salir()

    Unload Me
    
End Sub

Public Sub Accion(nAccion As ButtonAction)
    
    Select Case nAccion
                                
        Case vReport
            Call Imprimir
        
        Case vExit
            Call Salir
        
    End Select
    
End Sub

Private Sub cboFondo_Click()

    Dim adoRegistro As ADODB.Recordset
    
    strCodFondo = Valor_Caracter
    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(arrFondo(cboFondo.ListIndex))
    
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Fecha Vigente, Moneda, Tipo de Cambio ***
        .CommandText = "{ call up_ACSelDatosParametro(23,'" & strCodFondo & "','" & gstrCodAdministradora & "') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            gdatFechaActual = CVDate(adoRegistro("FechaCuota"))
            gstrFechaActual = Convertyyyymmdd(gdatFechaActual)
            
            gstrPeriodoActual = CStr(Year(gdatFechaActual))
            gstrMesActual = Format(Month(gdatFechaActual), "00")
            
            strCodMoneda = Trim(adoRegistro("CodMoneda"))
                        
            frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
            Me.Refresh
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
End Sub


Private Sub cboTipoInstrumento_Click()

    strCodTipoInstrumento = Valor_Caracter
    If cboTipoInstrumento.ListIndex < 0 Then Exit Sub
    
    strCodTipoInstrumento = Trim(arrTipoInstrumento(cboTipoInstrumento.ListIndex))
        
    strSQL = "SELECT DISTINCT II.CodTitulo CODIGO," & _
        "(RTRIM(II.Nemotecnico) + ' ' + RTRIM(II.CodTitulo) + ' ' + RTRIM(II.DescripTitulo)) DESCRIP " & _
        "FROM InstrumentoInversion II JOIN InversionKardex IK ON(IK.CodTitulo=II.CodTitulo) " & _
        "WHERE SaldoFinal > 0 AND II.CodFile='" & strCodTipoInstrumento & "' AND " & _
        "IK.CodFondo='" & strCodFondo & "' AND IK.CodAdministradora='" & gstrCodAdministradora & "' " & _
        "ORDER BY DESCRIP"
    
    CargarControlLista strSQL, cboTitulo, arrTitulo(), Valor_Caracter
        
    If cboTitulo.ListCount > 0 Then cboTitulo.ListIndex = 0
                        
End Sub

Private Sub cboTitulo_Click()

    strCodTitulo = Valor_Caracter
    If cboTitulo.ListIndex < 0 Then Exit Sub
    
    strCodTitulo = Trim(arrTitulo(cboTitulo.ListIndex))
    
    Call Buscar
    
End Sub

Private Sub chkUltimo_Click()

    cboTitulo_Click
    
End Sub

Private Sub Form_Activate()

    frmMainMdi.stbMdi.Panels(3).Text = Me.Caption
    Call CargarReportes
    
End Sub

Private Sub Form_Deactivate()

    Call OcultarReportes
    
End Sub


Private Sub Form_Load()

    Call InicializarValores
    Call CargarListas
    Call CargarReportes
    Call Buscar
    Call DarFormato
    
    Call ValidarPermisoUsoControl(Trim(gstrLoginUS), Me, Trim(App.Title) + Separador_Codigo_Objeto + _
    gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)
    
    CentrarForm Me
    
End Sub

Private Sub DarFormato()

    Dim intCont As Integer
    Dim elemento As Object
    
    For intCont = 0 To (lblDescrip.Count - 1)
        Call FormatoEtiqueta(lblDescrip(intCont), vbLeftJustify)
    Next
    
    For Each elemento In Me.Controls
    
        If TypeOf elemento Is TDBGrid Then
            Call FormatoGrilla(elemento)
        End If
    
    Next
            
End Sub

Public Sub Buscar()

    Set adoConsulta = New ADODB.Recordset

    strSQL = "SELECT NumKardex,FechaMovimiento,SaldoInicial,SaldoFinal," & _
        "CantMovimiento,(SaldoFinal / ValorNominal) CantTitulo," & _
        "CASE TipoMovimiento WHEN 'E' THEN 'Entrada' ELSE 'Salida' END TipoMovimiento " & _
        "FROM InversionKardex IK JOIN InstrumentoInversion II ON(II.CodTitulo=IK.CodTitulo) " & _
        "WHERE IK.CodTitulo='" & strCodTitulo & "' AND " & _
        "IK.CodFondo='" & strCodFondo & "' AND IK.CodAdministradora='" & gstrCodAdministradora & "'"
        
    If chkUltimo.Value Then strSQL = strSQL & "AND IndUltimoMovimiento='X' "
    strSQL = strSQL & "ORDER BY FechaMovimiento"

    strEstado = Reg_Defecto
    With adoConsulta
        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open strSQL
    End With

    tdgConsulta.DataSource = adoConsulta

    If adoConsulta.RecordCount > 0 Then strEstado = Reg_Consulta

    Me.MousePointer = vbDefault
    
End Sub

Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Kardex de Cartera"
    
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Text = "Kardex de Cartera 2"
    
    
End Sub
Private Sub CargarListas()
                            
    '*** Fondos ***
    strSQL = "{ call up_ACSelDatosParametro(74,'" & gstrCodAdministradora & "','" & gstrCodFondoContable & "') }"
    CargarControlLista strSQL, cboFondo, arrFondo(), Valor_Caracter
    
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
    
    '*** Tipo de Instrumento ***
    strSQL = "SELECT CodFile CODIGO,DescripFile DESCRIP FROM InversionFile WHERE IndInstrumento='X' AND IndVigente='X' ORDER BY DescripFile"
    CargarControlLista strSQL, cboTipoInstrumento, arrTipoInstrumento(), Sel_Defecto
    
    If cboTipoInstrumento.ListCount > 0 Then cboTipoInstrumento.ListIndex = 0
        
End Sub
Private Sub InicializarValores()
    
    strEstado = Reg_Defecto
                    
    Set cmdSalir.FormularioActivo = Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmKardexInversion = Nothing
    Call OcultarReportes
    frmMainMdi.stbMdi.Panels(3).Text = "Acción"
    
End Sub

Private Sub tdgConsulta_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)

    If ColIndex = 2 Then
        Call DarFormatoValor(Value, Decimales_Monto)
    End If
    
    If ColIndex = 4 Then
        Call DarFormatoValor(Value, Decimales_Monto)
    End If
    
    If ColIndex = 5 Then
        Call DarFormatoValor(Value, Decimales_Monto)
    End If
    
    If ColIndex = 6 Then
        Call DarFormatoValor(Value, Decimales_Monto)
    End If
    
End Sub

Private Sub tdgConsulta_HeadClick(ByVal ColIndex As Integer)
    
    Dim strColNameTDB  As String
    Static numColindex As Integer
    Static strPrevColumTDB As String
    '** agregar para que no se raye la seleccion de registro con ordenamiento
    strColNameTDB = tdgConsulta.Columns(ColIndex).DataField
    
    If strColNameTDB = strPrevColumTDB Then
        If indSortAsc Then
            indSortAsc = False
            indSortDesc = True
        Else
            indSortAsc = True
            indSortDesc = False
        End If
    Else
        indSortAsc = True
        indSortDesc = False
    End If
    '***

    tdgConsulta.Splits(0).Columns(numColindex).HeadingStyle.ForegroundPicture = Null

    Call OrdenarDBGrid(ColIndex, adoConsulta, tdgConsulta)
    
    numColindex = ColIndex
    
    '****
    strPrevColumTDB = strColNameTDB
    '***
    
End Sub
