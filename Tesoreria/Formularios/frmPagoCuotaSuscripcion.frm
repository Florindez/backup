VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmPagoCuotaSuscripcion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pago de Cuotas por Suscripción"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5520
   ScaleWidth      =   7875
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   6000
      TabIndex        =   3
      Top             =   4680
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
      Caption0        =   "&Salir"
      Tag0            =   "9"
      ToolTipText0    =   "Salir"
      UserControlWidth=   1200
   End
   Begin TAMControls2.ucBotonEdicion2 cmdOpcion 
      Height          =   735
      Left            =   600
      TabIndex        =   2
      Top             =   4680
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   1296
      Buttons         =   2
      Caption0        =   "&Modificar"
      Tag0            =   "3"
      ToolTipText0    =   "Modificar"
      Caption1        =   "&Buscar"
      Tag1            =   "5"
      ToolTipText1    =   "Buscar"
      UserControlWidth=   2700
   End
   Begin TabDlg.SSTab tabPagos 
      Height          =   4245
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   7488
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Lista"
      TabPicture(0)   =   "frmPagoCuotaSuscripcion.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraCriterio"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "tdgConsulta"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "frmPagoCuotaSuscripcion.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraDatos"
      Tab(1).Control(1)=   "cmdAccion"
      Tab(1).ControlCount=   2
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   -70800
         TabIndex        =   8
         Top             =   3360
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   1296
         Buttons         =   2
         Caption0        =   "&Guardar"
         Tag0            =   "2"
         ToolTipText0    =   "Guardar"
         Caption1        =   "&Cancelar"
         Tag1            =   "8"
         ToolTipText1    =   "Cancelar"
         UserControlWidth=   2700
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmPagoCuotaSuscripcion.frx":0038
         Height          =   1695
         Left            =   240
         OleObjectBlob   =   "frmPagoCuotaSuscripcion.frx":0052
         TabIndex        =   18
         Top             =   1920
         Width           =   6825
      End
      Begin VB.Frame fraCriterio 
         Caption         =   "Criterios de Búsqueda"
         Height          =   1305
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   6825
         Begin VB.ComboBox cboParticipe 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   760
            Width           =   5055
         End
         Begin VB.ComboBox cboFondo 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   360
            Width           =   5055
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Partícipe"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   6
            Left            =   360
            TabIndex        =   17
            Top             =   780
            Width           =   645
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fondo"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   12
            Top             =   360
            Width           =   450
         End
      End
      Begin VB.Frame fraDatos 
         Caption         =   "Datos"
         Height          =   2775
         Left            =   -74760
         TabIndex        =   6
         Top             =   480
         Width           =   6855
         Begin MSComCtl2.DTPicker dtpFechaLiquidacion 
            Height          =   285
            Left            =   1800
            TabIndex        =   4
            Top             =   1740
            Width           =   1820
            _ExtentX        =   3201
            _ExtentY        =   503
            _Version        =   393216
            Format          =   175964161
            CurrentDate     =   38949
         End
         Begin VB.TextBox txtMontoPago 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1800
            MaxLength       =   12
            TabIndex        =   7
            Text            =   " "
            Top             =   2160
            Width           =   1820
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Partícipe"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   5
            Left            =   360
            TabIndex        =   20
            Top             =   920
            Width           =   795
         End
         Begin VB.Label lblDescripParticipe 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1800
            TabIndex        =   19
            Top             =   900
            Width           =   4575
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   4
            Left            =   360
            TabIndex        =   16
            Top             =   2180
            Width           =   450
         End
         Begin VB.Label lblNumSecuencial 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1800
            TabIndex        =   15
            Top             =   1320
            Width           =   1820
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fondo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   360
            TabIndex        =   14
            Top             =   500
            Width           =   540
         End
         Begin VB.Label lblDescripFondo 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1800
            TabIndex        =   13
            Top             =   480
            Width           =   4575
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Secuencial"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   10
            Top             =   1340
            Width           =   795
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Liquidación"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   9
            Top             =   1755
            Width           =   810
         End
      End
   End
End
Attribute VB_Name = "frmPagoCuotaSuscripcion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strCodFondo         As String, strCodParticipe      As String
Dim strEstado           As String, strSQL               As String
Dim curMontoEmitido     As Currency
Dim adoConsulta         As ADODB.Recordset
Dim indSortAsc          As Boolean, indSortDesc         As Boolean

Private Sub cboFondo_Click()

    Dim adoRegistro     As ADODB.Recordset
    
    strCodFondo = Valor_Caracter
    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(garrFondo(cboFondo.ListIndex))
    
    '*** Participe con Cronograma ***
    If gstrFormulario = "frmConfirmacionSolicitud" Then
        strSQL = "SELECT DISTINCT PPS.CodParticipe CODIGO, DescripParticipe DESCRIP " & _
            "FROM ParticipePagoSuscripcionTmp PPS JOIN ParticipeContrato PC ON(PC.CodParticipe=PPS.CodParticipe) " & _
            "WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
    Else
        strSQL = "SELECT DISTINCT PPS.CodParticipe CODIGO, DescripParticipe DESCRIP " & _
            "FROM ParticipePagoSuscripcion PPS JOIN ParticipeContrato PC ON(PC.CodParticipe=PPS.CodParticipe) " & _
            "WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
    End If
    CargarControlLista strSQL, cboParticipe, garrParticipe(), Sel_Defecto
    
    If cboParticipe.ListCount > 0 Then cboParticipe.ListIndex = 0
    
    '*** Obtener Monto Emitido ***
    Set adoRegistro = New ADODB.Recordset
        
    adoComm.CommandText = "SELECT MontoEmitido FROM Fondo WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
    Set adoRegistro = adoComm.Execute
    
    If Not adoRegistro.EOF Then
        curMontoEmitido = CCur(adoRegistro("MontoEmitido"))
    End If
    adoRegistro.Close: Set adoRegistro = Nothing
    
    Call Buscar

End Sub

Private Sub cboParticipe_Click()

    gstrCodParticipe = Valor_Caracter
    If cboParticipe.ListIndex < 0 Then Exit Sub
    
    gstrCodParticipe = Trim(garrParticipe(cboParticipe.ListIndex))
    
    Call Buscar
    
End Sub

Private Sub Form_Activate()

    Call CargarReportes
    
End Sub

Private Sub CargarReportes()

'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Cronograma Cuotas por Partícipe"
    
End Sub

Public Sub Accion(nAccion As ButtonAction)
    
    Select Case nAccion
        
        Case vNew
            Call Adicionar
        Case vModify
            Call Modificar
        Case vSearch
            Call Buscar
        Case vReport
            Call Imprimir
        Case vSave
            Call Grabar
        Case vCancel
            Call Cancelar
        Case vExit
            Call Salir
        
    End Select
    
End Sub

Public Sub Salir()

    Unload Me
    
End Sub
Public Sub Cancelar()

    cmdOpcion.Visible = True
    With tabPagos
        .TabEnabled(0) = True
        .Tab = 0
    End With
    strEstado = Reg_Consulta
    
End Sub
Public Sub Grabar()
                        
    Dim intAccion       As Integer, lngNumError         As Integer
    Dim curMontoTotal   As Currency, strNumSolicitud    As String
    
    If strEstado = Reg_Consulta Then Exit Sub
            
    On Error GoTo CtrlError
    
    If strEstado = Reg_Edicion Then
        If TodoOk() Then
            frmMainMdi.stbMdi.Panels(3).Text = "Grabar Cronograma de Pagos de Cuotas..."
            
            strNumSolicitud = tdgConsulta.Columns(0).Value
            adoConsulta.MoveFirst
            curMontoTotal = 0
                                                
            Do While Not adoConsulta.EOF
                
                If adoConsulta.Fields("NumSecuencial") <> CInt(lblNumSecuencial.Caption) Then
                    curMontoTotal = curMontoTotal + CDbl(adoConsulta.Fields("MontoLiquidacion"))
                End If
                
                adoConsulta.MoveNext
            Loop
            
            curMontoTotal = curMontoTotal + CCur(txtMontoPago.Text)
                                    
            If curMontoTotal > curMontoEmitido Then
                MsgBox "El Monto Total excede el Monto Emitido, verifique!", vbCritical, Me.Caption
                Exit Sub
            End If
            
            Me.MousePointer = vbHourglass
            
            With adoComm
                .CommandText = "UPDATE ParticipePagoSuscripcion SET " & _
                    "FechaLiquidacion='" & Convertyyyymmdd(dtpFechaLiquidacion.Value) & "'," & _
                    "MontoLiquidacion=" & CDec(txtMontoPago.Text) & " " & _
                    "WHERE NumSecuencial=" & CInt(lblNumSecuencial.Caption) & " AND " & _
                    "NumSolicitud='" & strNumSolicitud & "' AND " & _
                    "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
                adoConn.Execute .CommandText
            End With
        
            Me.MousePointer = vbDefault
                            
            MsgBox Mensaje_Edicion_Exitosa, vbExclamation
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
            cmdOpcion.Visible = True
            With tabPagos
                .TabEnabled(0) = True
                .Tab = 0
            End With
            Call Buscar
        End If
    End If
    Exit Sub
    
CtrlError:
    Me.MousePointer = vbDefault
    intAccion = ControlErrores
    Select Case intAccion
        Case 0: Resume
        Case 1: Resume Next
        Case 2: Exit Sub
        Case Else
            lngNumError = err.Number
            err.Raise Number:=lngNumError
            err.Clear
    End Select
    
End Sub

Private Function TodoOk() As Boolean
        
    TodoOk = False
            
    If CCur(txtMontoPago.Text) = 0 Then
        MsgBox "El Monto de Pago no puede ser cero!.", vbCritical, gstrNombreEmpresa
        Exit Function
    End If
        
    '*** Si todo paso OK ***
    TodoOk = True
  
End Function
Public Sub Imprimir()

End Sub

Public Sub Buscar()
        
    Dim strSQL As String
    
    If gstrFormulario = "frmConfirmacionSolicitud" Then
        strSQL = "SELECT NumSolicitud,NumSecuencial,FechaLiquidacion,MontoLiquidacion,(CASE FechaPago WHEN '01/01/1900' THEN NULL ELSE FechaPago END) FechaPago " & _
            "FROM ParticipePagoSuscripcionTmp  " & _
            "WHERE CodParticipe='" & gstrCodParticipe & "' AND " & _
            "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' " & _
            "ORDER BY NumSecuencial"
    Else
        strSQL = "SELECT NumSolicitud,NumSecuencial,FechaLiquidacion,MontoLiquidacion,(CASE FechaPago WHEN '01/01/1900' THEN NULL ELSE FechaPago END) FechaPago " & _
            "FROM ParticipePagoSuscripcion  " & _
            "WHERE CodParticipe='" & gstrCodParticipe & "' AND " & _
            "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' " & _
            "ORDER BY NumSecuencial"
    End If
    
    Set adoConsulta = New ADODB.Recordset
    
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
    
End Sub

Public Sub Modificar()

    If strEstado = Reg_Consulta Then
        strEstado = Reg_Edicion
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabPagos
            .TabEnabled(0) = False
            .Tab = 1
        End With
    End If
    
End Sub

Private Sub LlenarFormulario(strModo As String)

    Dim intNumSecuencial    As Integer
    
    Select Case strModo
        Case Reg_Edicion
            intNumSecuencial = CInt(tdgConsulta.Columns(1).Value)
            
            lblDescripFondo.Caption = Trim(cboFondo.Text)
            lblDescripParticipe.Caption = Trim(cboParticipe.Text)
            lblNumSecuencial.Caption = CStr(intNumSecuencial)
            
            dtpFechaLiquidacion.Value = tdgConsulta.Columns(2).Value
            txtMontoPago.Text = CStr(CCur(tdgConsulta.Columns(3).Value))
    
    End Select
    
End Sub

Public Sub SubImprimir(index As Integer) 'No se encuentra el reporte PagoCuotaSuscripcion.RPT

'    Dim frmReporte              As frmVisorReporte
'    Dim aReportParamS(), aReportParamF(), aReportParamFn()
'    Dim strFechaDesde           As String, strFechaHasta        As String
'    Dim intAccion               As Integer
'    Dim lngNumError             As Long
'
'    If tabPagos.Tab = 1 Then Exit Sub
'
'    Select Case index
'        Case 1
'            gstrNameRepo = "PagoCuotaSuscripcion"
'
'            Set frmReporte = New frmVisorReporte
'
'            ReDim aReportParamS(1)
'            ReDim aReportParamFn(2)
'            ReDim aReportParamF(2)
'
'            aReportParamFn(0) = "Usuario"
'            aReportParamFn(1) = "Hora"
'            aReportParamFn(2) = "NombreEmpresa"
'
'            aReportParamF(0) = gstrLogin
'            aReportParamF(1) = Format(Time(), "hh:mm:ss")
'            aReportParamF(2) = gstrNombreEmpresa & Space(1)
'
'            aReportParamS(0) = "001"
'            aReportParamS(1) = gstrCodAdministradora
'
'    End Select
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
        
End Sub
Public Sub Adicionar()
    
    frmMainMdi.stbMdi.Panels(3).Text = "Adicionar Perfil..."
                
    strEstado = Reg_Adicion
    LlenarFormulario strEstado
    cmdOpcion.Visible = False
    With tabPagos
        .TabEnabled(0) = False
        .Tab = 1
    End With
                
End Sub
Private Sub Form_Deactivate()
    
    Call OcultarReportes
    Unload Me
    
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
        
    Call ValidaExisteTipoCambio(Codigo_TipoCambio_SBS, gstrFechaActual)
    
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
Private Sub CargarListas()
    
    '*** Fondos ***
    strSQL = "{ call up_ACSelDatosParametro(74,'" & gstrCodAdministradora & "','" & gstrCodFondoContable & "') }"
    CargarControlLista strSQL, cboFondo, garrFondo(), Valor_Caracter
    
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
        
End Sub
Private Sub InicializarValores()

    '*** Valores Iniciales ***
    strEstado = Reg_Defecto
    tabPagos.Tab = 0
    
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgConsulta.Columns(0).Width = tdgConsulta.Width * 0.01 * 16
    tdgConsulta.Columns(1).Width = tdgConsulta.Width * 0.01 * 9
    tdgConsulta.Columns(2).Width = tdgConsulta.Width * 0.01 * 16
    tdgConsulta.Columns(3).Width = tdgConsulta.Width * 0.01 * 34
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
        
    Call OcultarReportes
    gstrCodParticipe = Valor_Caracter
    Set frmPagoCuotaSuscripcion = Nothing
    
End Sub

Private Sub tabPagos_Click(PreviousTab As Integer)

    Select Case tabPagos.Tab
        Case 1
            If gstrFormulario = "frmConfirmacionSolicitud" Then tabPagos.Tab = 0
            If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vModify)
            If strEstado = Reg_Defecto Then tabPagos.Tab = 0
        
    End Select
    
End Sub

Private Sub tdgConsulta_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
        
    If ColIndex = 3 Then
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

Private Sub txtMontoPago_Change()

    Call FormatoCajaTexto(txtMontoPago, Decimales_Monto)
    
End Sub

Private Sub txtMontoPago_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtMontoPago, Decimales_Monto)
    
End Sub

