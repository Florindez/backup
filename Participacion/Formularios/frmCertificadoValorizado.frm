VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmCertificadoValorizado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estado de Cuenta por Partícipe"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7920
   ScaleWidth      =   10695
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   8400
      TabIndex        =   5
      Top             =   7080
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
      Caption0        =   "&Salir"
      Tag0            =   "9"
      Visible0        =   0   'False
      ToolTipText0    =   "Salir"
      UserControlWidth=   1200
   End
   Begin TAMControls2.ucBotonEdicion2 cmdOpcion 
      Height          =   735
      Left            =   600
      TabIndex        =   4
      Top             =   7080
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
      Caption0        =   "&Buscar"
      Tag0            =   "5"
      Visible0        =   0   'False
      ToolTipText0    =   "Buscar"
      UserControlWidth=   1200
   End
   Begin VB.Frame fraCertificado 
      Caption         =   "Criterios de Búsqueda"
      Height          =   6855
      Left            =   140
      TabIndex        =   6
      Top             =   120
      Width           =   10360
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta2 
         Bindings        =   "frmCertificadoValorizado.frx":0000
         Height          =   1815
         Left            =   240
         OleObjectBlob   =   "frmCertificadoValorizado.frx":001B
         TabIndex        =   25
         Top             =   4800
         Width           =   9855
      End
      Begin VB.ComboBox cboFondo 
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   7800
      End
      Begin VB.ComboBox cboTipoDocumento 
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1125
         Width           =   2900
      End
      Begin VB.TextBox txtNumDocumento 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         TabIndex        =   12
         Top             =   1500
         Width           =   2900
      End
      Begin VB.CommandButton cmdBusqueda 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9680
         TabIndex        =   1
         ToolTipText     =   "Búsqueda de Partícipe"
         Top             =   765
         Width           =   375
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmCertificadoValorizado.frx":41C8
         Height          =   1935
         Left            =   240
         OleObjectBlob   =   "frmCertificadoValorizado.frx":41E2
         TabIndex        =   11
         Top             =   2640
         Width           =   9855
      End
      Begin MSComCtl2.DTPicker dtpFechaDesde 
         Height          =   285
         Left            =   7185
         TabIndex        =   2
         Top             =   1125
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   503
         _Version        =   393216
         Format          =   197656577
         CurrentDate     =   38069
      End
      Begin MSComCtl2.DTPicker dtpFechaHasta 
         Height          =   285
         Left            =   7185
         TabIndex        =   3
         Top             =   1500
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   503
         _Version        =   393216
         Format          =   197656577
         CurrentDate     =   38069
      End
      Begin VB.Label lblValorCuota 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7185
         TabIndex        =   24
         Top             =   1860
         Width           =   2895
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         Caption         =   "Valor Cuota"
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
         Index           =   9
         Left            =   5520
         TabIndex        =   23
         Top             =   1875
         Width           =   1005
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Desde"
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
         Height          =   285
         Index           =   8
         Left            =   5520
         TabIndex        =   22
         Top             =   1140
         Width           =   855
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Hasta"
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
         Height          =   285
         Index           =   7
         Left            =   5520
         TabIndex        =   21
         Top             =   1515
         Width           =   855
      End
      Begin VB.Label lblDescrip 
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
         Height          =   285
         Index           =   4
         Left            =   240
         TabIndex        =   20
         Top             =   405
         Width           =   1095
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         Caption         =   "Num.Documento"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   19
         Top             =   1515
         Width           =   1200
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Tipo Documento"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   18
         Top             =   1140
         Width           =   1455
      End
      Begin VB.Label lblDescripParticipe 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2280
         TabIndex        =   17
         Top             =   765
         Width           =   7360
      End
      Begin VB.Label lblDescripTipoParticipe 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2280
         TabIndex        =   16
         Top             =   1860
         Width           =   2900
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Partícipe"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   15
         Top             =   1875
         Width           =   1005
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         Caption         =   "Partícipe"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   780
         Width           =   645
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Actual Cuotas"
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
         Left            =   240
         TabIndex        =   10
         Top             =   2220
         Width           =   1740
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Inversión"
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
         Index           =   2
         Left            =   5520
         TabIndex        =   9
         Top             =   2220
         Width           =   1335
      End
      Begin VB.Label lblSaldoInversion 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7185
         TabIndex        =   8
         Top             =   2205
         Width           =   2895
      End
      Begin VB.Label lblSaldoCuotas 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2280
         TabIndex        =   7
         Top             =   2205
         Width           =   2900
      End
   End
End
Attribute VB_Name = "frmCertificadoValorizado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrFondo()              As String

Dim strCodFondo             As String
Dim strCodTipoDocumento     As String, strCodMoneda         As String
Dim strFechaDesde           As String, strFechaHasta        As String
Dim strEstado               As String
Dim adoConsulta             As ADODB.Recordset
Dim adoConsulta2            As ADODB.Recordset
Dim indSortAsc              As Boolean, indSortDesc         As Boolean
Dim strCodCertificado  As String

Public Sub Accion(nAccion As ButtonAction)
    
    Select Case nAccion
        
        Case vNew
            Call Adicionar
        Case vQuery
            Call Modificar
        Case vDelete
            Call Eliminar
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
Public Sub Adicionar()

End Sub
Public Sub Buscar()
          
    On Error GoTo Error1            '/**/ HMC Habilitamos la rutina de Errores.

    Dim strSql      As String
    
    Set adoConsulta = New ADODB.Recordset
    
    strFechaDesde = Convertyyyymmdd(dtpFechaDesde.Value)
    strFechaHasta = Convertyyyymmdd(DateAdd("d", 1, dtpFechaHasta.Value))
            
    strSql = "SELECT NumCertificado,FechaSuscripcion,ValorCuota,CantCuotas,(ValorCuota * CantCuotas) Inversion," & _
        "FechaRedencion=CASE Convert(char(8),FechaRedencion,112) WHEN '19000101' THEN NULL ELSE FechaRedencion END," & _
        "(CantCuotas * " & CDbl(lblValorCuota.Caption) & ") InversionValorizada " & _
        "FROM ParticipeCertificado " & _
        "WHERE (FechaOperacion>='" & strFechaDesde & "' AND FechaOperacion<'" & strFechaHasta & "') AND " & _
        "CodParticipe='" & gstrCodParticipe & "' AND CodFondo='" & strCodFondo & "' AND " & _
        "CodAdministradora='" & gstrCodAdministradora & "' " & _
        "ORDER BY FechaSuscripcion"
    
    strEstado = Reg_Defecto
    With adoConsulta
        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open strSql
    End With
        
    tdgConsulta.DataSource = adoConsulta
    
    If adoConsulta.RecordCount > 0 Then strEstado = Reg_Consulta
    Call CargarTransferencias
    Exit Sub
    
Error1:
    MsgBox DescripcionError & vbNewLine & DescripcionTecnica & err.Description, vbExclamation, TituloError ' Mostrar Error
    
End Sub

Public Sub Cancelar()

End Sub

Public Sub Eliminar()

End Sub

Public Sub Grabar()

End Sub

Public Sub Imprimir()
                   
End Sub

Public Sub Modificar()

End Sub

Public Sub ObtenerSaldosParticipacion()

    Dim adoRegistro         As ADODB.Recordset
    Dim strFechaHastaMenos1 As String
    
    Set adoRegistro = New ADODB.Recordset
    '*** Saldos ***
    strFechaDesde = Convertyyyymmdd(dtpFechaDesde.Value)
    strFechaHasta = Convertyyyymmdd(DateAdd("d", 1, dtpFechaHasta.Value))
    strFechaHastaMenos1 = Convertyyyymmdd(dtpFechaHasta.Value)
    
    adoComm.CommandText = "SELECT SUM(CantCuotas) CantCuotas," & _
        "(SUM(CantCuotas) * (SELECT Round(ValorCuotaFinal, 8) FROM FondoValorCuota WHERE (FechaCuota>='" & strFechaHastaMenos1 & "' AND FechaCuota<'" & strFechaHasta & "') AND CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "')) SaldoInversion " & _
        "FROM ParticipeCertificado " & _
        "WHERE FechaOperacion < '" & strFechaHasta & "' AND (FechaRedencion='19000101' OR FechaRedencion  > '" & strFechaDesde & "') AND " & _
        "CodParticipe='" & gstrCodParticipe & "' AND CodFondo='" & strCodFondo & "' AND " & _
        "CodAdministradora='" & gstrCodAdministradora & "'"
    Set adoRegistro = adoComm.Execute

    If Not adoRegistro.EOF Then
        If IsNull(adoRegistro("CantCuotas")) Then
            lblSaldoCuotas.Caption = "0"
        Else
            lblSaldoCuotas.Caption = CStr(adoRegistro("CantCuotas"))
        End If
        If IsNull(adoRegistro("SaldoInversion")) Then
            lblSaldoInversion.Caption = "0"
        Else
            lblSaldoInversion.Caption = CStr(adoRegistro("SaldoInversion"))
        End If
    End If
    adoRegistro.Close: Set adoRegistro = Nothing
    
End Sub

Public Sub Salir()

    Unload Me
    
End Sub

Public Sub SubImprimir(Index As Integer)

    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    Dim strFechaDesde           As String, strFechaHasta        As String
    Dim strFechaHastaMas1Dia    As String
        
    strFechaDesde = Convertyyyymmdd(dtpFechaDesde.Value)
    'strFechaHasta = Convertyyyymmdd(DateAdd("d", 1, dtpFechaHasta.Value))
    strFechaHasta = Convertyyyymmdd(dtpFechaHasta.Value)
    'strFechaHasta = Convertyyyymmdd(DateAdd("d", 1, dtpFechaHasta.Value))
        
    If Not TodoOK() Then Exit Sub
        
    Select Case Index
        Case 1
            gstrNameRepo = "EstadoCuenta"
                        
            Set frmReporte = New frmVisorReporte

            ReDim aReportParamS(5)
            ReDim aReportParamFn(4)
            ReDim aReportParamF(4)
                        
            aReportParamFn(0) = "Fondo"
            aReportParamFn(1) = "Moneda"
            aReportParamFn(2) = "NombreEmpresa"
            aReportParamFn(3) = "Usuario"
            aReportParamFn(4) = "Hora"
            
            aReportParamF(0) = Trim(cboFondo.Text)
            aReportParamF(1) = ObtenerDescripcionMoneda(strCodMoneda)
            aReportParamF(2) = gstrNombreEmpresa
            aReportParamF(3) = gstrLogin
            aReportParamF(4) = Format(Time(), "hh:mm:ss")
                        
            aReportParamS(0) = strCodFondo
            aReportParamS(1) = gstrCodAdministradora
            aReportParamS(2) = gstrCodParticipe
            aReportParamS(3) = strFechaDesde 'dtpFechaDesde.Value
            aReportParamS(4) = strFechaHasta 'dtpFechaHasta.Value
            aReportParamS(5) = strCodMoneda
            
        Case 2
            gstrNameRepo = "EstadoCuentaM"
                        
            Set frmReporte = New frmVisorReporte

            ReDim aReportParamS(5)
            ReDim aReportParamFn(4)
            ReDim aReportParamF(4)
                       
            aReportParamFn(0) = "Fondo"
            aReportParamFn(1) = "Moneda"
            aReportParamFn(2) = "NombreEmpresa"
            aReportParamFn(3) = "Usuario"
            aReportParamFn(4) = "Hora"
            
            aReportParamF(0) = Trim(cboFondo.Text)
            aReportParamF(1) = ObtenerDescripcionMoneda(strCodMoneda)
            aReportParamF(2) = gstrNombreEmpresa
            aReportParamF(3) = gstrLogin
            aReportParamF(4) = Format(Time(), "hh:mm:ss")
                        
            aReportParamS(0) = strCodFondo
            aReportParamS(1) = gstrCodAdministradora
            aReportParamS(2) = "%"
            aReportParamS(3) = strFechaDesde 'dtpFechaDesde.Value
            aReportParamS(4) = strFechaHasta 'dtpFechaHasta.Value
            aReportParamS(5) = strCodMoneda
            
        Case 11
            gstrNameRepo = "EstadoCuentaCertificado"
                        
            Set frmReporte = New frmVisorReporte

            ReDim aReportParamS(4)
            ReDim aReportParamFn(4)
            ReDim aReportParamF(4)
                       
            aReportParamFn(0) = "Fondo"
            aReportParamFn(1) = "Moneda"
            aReportParamFn(2) = "NombreEmpresa"
            aReportParamFn(3) = "Usuario"
            aReportParamFn(4) = "Hora"
            
            aReportParamF(0) = Trim(cboFondo.Text)
            aReportParamF(1) = ObtenerDescripcionMoneda(strCodMoneda)
            aReportParamF(2) = gstrNombreEmpresa
            aReportParamF(3) = gstrLogin
            aReportParamF(4) = Format(Time(), "hh:mm:ss")
                        
            aReportParamS(0) = strCodFondo
            aReportParamS(1) = gstrCodAdministradora
            aReportParamS(2) = gstrCodParticipe
            aReportParamS(3) = strCodCertificado
            aReportParamS(4) = gstrFechaActual
       
        Case 3: Exit Sub
        
    End Select

    gstrSelFrml = Valor_Caracter
    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"

    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())

    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
    frmReporte.Show vbModal

    Set frmReporte = Nothing

    Screen.MousePointer = vbNormal
    
End Sub

Private Sub cboFondo_Click()

    Dim adoRegistro As ADODB.Recordset, adoTemporal As ADODB.Recordset
    
    strCodFondo = Valor_Caracter
    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(arrFondo(cboFondo.ListIndex))
    
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Fecha Vigente, Valor Cuota, Valor Cuota T-1, Moneda y Cantidad Inicial de Cuotas del Fondo ***
        .CommandText = "{ call up_ACSelDatosParametro(23,'" & strCodFondo & "','" & gstrCodAdministradora & "') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            
            '*** Fecha Máxima de Consulta ***
            dtpFechaDesde.MaxDate = adoRegistro("FechaCuota")
            dtpFechaHasta.MaxDate = adoRegistro("FechaCuota")
            
            dtpFechaDesde.Value = adoRegistro("FechaCuota") 'DateAdd("d", -1, CVDate(adoRegistro("FechaCuota")))
            dtpFechaHasta.Value = dtpFechaDesde.Value
            
            strCodMoneda = Trim(adoRegistro("CodMoneda"))
            lblValorCuota.Caption = CStr(adoRegistro("ValorCuotaInicial"))
            
            gdatFechaActual = adoRegistro("FechaCuota")
            frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
            

        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    lblDescrip(2).Caption = "Saldo Inversión" & Space(1) & ObtenerSignoMoneda(strCodMoneda)
    
End Sub

Private Sub cboTipoDocumento_Click()

    strCodTipoDocumento = Valor_Caracter
    If cboTipoDocumento.ListIndex < 0 Then Exit Sub
    
    strCodTipoDocumento = Trim(garrTipoDocumento(cboTipoDocumento.ListIndex))
    
End Sub

Private Sub cmdBusqueda_Click()

    gstrFormulario = Me.Name
    frmBusquedaParticipeP.Show vbModal
    
End Sub

Private Sub Form_Activate()

    Call CargarReportes
    
End Sub

Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Estado de Cuenta Individual"
        
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Text = "Estado de Cuenta Masivo"
    
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo11").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo11").Text = "Estado de Cuenta Participe"
        
        
End Sub
Private Sub Form_Deactivate()

    ReDim garrTipoDocumento(0)
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
Private Sub CargarListas()
    
    Dim intRegistro As Integer
    Dim strSql      As String
    
    '*** Fondos ***
    strSql = "{ call up_ACSelDatosParametro(74,'" & gstrCodAdministradora & "','" & gstrCodFondoContable & "') }"
    CargarControlLista strSql, cboFondo, arrFondo(), Valor_Caracter
    
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
        
    '*** Tipo Documento Identidad ***
    strSql = "{ call up_ACSelDatos(11) }"
    CargarControlLista strSql, cboTipoDocumento, garrTipoDocumento(), Sel_Defecto
    
    If cboTipoDocumento.ListCount > 0 Then cboTipoDocumento.ListIndex = 0
                        
End Sub
Private Sub InicializarValores()

    strEstado = Reg_Defecto
    
    '*** Verificando Nivel de Acceso de Usuario ***
'    strNivAcceso = AccesoForm(gstrNomOpc, gstrNumInd)
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    
    
End Sub
Private Sub Form_Unload(Cancel As Integer)

    Call OcultarReportes
    Set frmCertificadoValorizado = Nothing
    gstrCodParticipe = Valor_Caracter
    frmMainMdi.stbMdi.Panels(3).Text = "Acción..."
        
End Sub

Private Sub optVigencia_Click(Index As Integer, Value As Integer)

'    If cmbCodFondo.ListIndex > 0 And strCodPart <> "" Then
'        '*** Configurar la grilla ***
'        ReDim aGrdCnf(1 To 10)
''        aGrdCnf(1).TitDes = "Fondo Mutuo"
''        aGrdCnf(1).DatNom = "DSC_FOND"
'
'        aGrdCnf(1).TitDes = "Nro. Cert."
'        aGrdCnf(1).DatNom = "NRO_CERT"
'
'        aGrdCnf(2).TitDes = "Vig."
'        aGrdCnf(2).DatNom = "FLG_VIGE"
'
'        aGrdCnf(3).TitDes = "Fec. Susc"
'        aGrdCnf(3).DatNom = "FCH_SUSC"
'
'        aGrdCnf(4).TitDes = "Fec. Rede"
'        aGrdCnf(4).DatNom = "FCH_REDE"
'
'        aGrdCnf(5).TitDes = "Oper."
'        aGrdCnf(5).DatNom = "TIP_OPER"

'        aGrdCnf(6).TitDes = "Fec. Crea"
'        aGrdCnf(6).DatNom = "FCH_CREA"
'
'        aGrdCnf(7).TitDes = "Cnt. Cuotas"
'        aGrdCnf(7).DatNom = "CNT_CUOT"
'
'        aGrdCnf(8).TitDes = "Valor Total"
'        aGrdCnf(8).DatNom = "SLD_FINA"
'
'        aGrdCnf(9).TitDes = "FLAG GARA"
'        aGrdCnf(9).DatNom = "FLG_GARA"
'
'        aGrdCnf(10).TitDes = "FLAG CUST"
'        aGrdCnf(10).DatNom = "FLG_CUST"
'
'        '*** Configurar la grilla ***
'        ReDim aGrdCnf(1 To 5)
'        aGrdCnf(1).TitDes = "Fondo"
'        aGrdCnf(1).TitJus = 2
'        aGrdCnf(1).DatNom = "DSC_FOND"
'        aGrdCnf(1).DatAnc = 2235
'
'        aGrdCnf(2).TitDes = "Moneda"
'        aGrdCnf(2).DatNom = "COD_MONE"
'        aGrdCnf(2).DatAnc = 1365
'        aGrdCnf(2).DatJus = 2
'        aGrdCnf(2).TitJus = 2
'
'        aGrdCnf(3).TitDes = "Valor Cuota"
'        aGrdCnf(3).DatNom = "VAL_CUOT"
'        aGrdCnf(3).DatAnc = 1530
'        aGrdCnf(3).DatJus = 1
'        aGrdCnf(3).TitJus = 2
'
'        aGrdCnf(4).TitDes = "Cnt. Cuotas"
'        aGrdCnf(4).DatNom = "CNT_CUOT"
'        aGrdCnf(4).DatAnc = 1455
'        aGrdCnf(4).DatFmt = "C"
'        aGrdCnf(4).DatJus = 1
'        aGrdCnf(4).TitJus = 2
'
'        aGrdCnf(5).TitDes = "Valor Total"
'        aGrdCnf(5).DatNom = "SLD_FINA"
'        aGrdCnf(5).DatAnc = 1605
'        aGrdCnf(5).DatFmt = "N"
'        aGrdCnf(5).DatJus = 1
'        aGrdCnf(5).TitJus = 2
'
'        '** Estado del Certificado
'        If optVigencia(0).Value = True Then
'            'adoComm.CommandText = "Sp_INF_ConsCertif '14', '" & Format(MhDFchCuota.Text, "yyyymmdd") & "', '" & strCodPart & "'"
'            gstrSQL = "Sp_INF_ConsCertif '14', '" & Convertyyyymmdd(MhDFchCuota.Value) & "', '" & strCodPart & "'"
'        ElseIf optVigencia(1).Value = True Then
'            'adoComm.CommandText = "Sp_INF_ConsCertif '15', '" & Format(MhDFchCuota.Text, "yyyymmdd") & "', '" & strCodPart & "'"
'            gstrSQL = "Sp_INF_ConsCertif '15', '" & Convertyyyymmdd(MhDFchCuota.Value) & "', '" & strCodPart & "'"
'        End If
'        adoComm.CommandText = gstrSQL
'    End If
'
'    If strCodPart <> "" Then
'        Set adoRecord = adoComm.Execute
'        Call LlenarGrid(grdCertif, adoRecord, aGrdCnf(), adirreg())
'        frmINFConsCertif.Refresh
'    End If

End Sub

Private Sub lblSaldoCuotas_Change()

    Call FormatoMillarEtiqueta(lblSaldoCuotas, Decimales_CantCuota)
    
End Sub

Private Sub lblSaldoInversion_Change()

    Call FormatoMillarEtiqueta(lblSaldoInversion, Decimales_Monto)
    
End Sub

Private Sub lblValorCuota_Change()

    Call FormatoMillarEtiqueta(lblValorCuota, Decimales_ValorCuota_Cierre)
    
End Sub

Private Sub tdgConsulta_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)

    If ColIndex = 2 Then
        Call DarFormatoValor(Value, Decimales_CantCuota)
    End If
    
    If ColIndex = 3 Then
        Call DarFormatoValor(Value, Decimales_ValorCuota)
    End If
    
    If ColIndex = 5 Then
        Call DarFormatoValor(Value, Decimales_Monto)
    End If
    
    If ColIndex = 6 Then
        Call DarFormatoValor(Value, Decimales_Monto)
    End If
    
    strCodCertificado = Me.tdgConsulta.Columns(0)
    
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

Private Sub tdgConsulta2_HeadClick(ByVal ColIndex As Integer)
    
    Dim strColNameTDB  As String
    Static numColindex As Integer
    Static strPrevColumTDB As String
    '** agregar para que no se raye la seleccion de registro con ordenamiento
    strColNameTDB = tdgConsulta2.Columns(ColIndex).DataField
    
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

    tdgConsulta2.Splits(0).Columns(numColindex).HeadingStyle.ForegroundPicture = Null

    Call OrdenarDBGrid(ColIndex, adoConsulta2, tdgConsulta2)
    
    numColindex = ColIndex
    
    '****
    strPrevColumTDB = strColNameTDB
    '***
    
End Sub

Private Sub txtNumDocumento_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Call ObtenerDatosParticipe
        Call ObtenerSaldosParticipacion
        Call Buscar
    End If
    
End Sub

Private Sub ObtenerDatosParticipe()
    
    Dim adoRegistro         As ADODB.Recordset

    Set adoRegistro = New ADODB.Recordset
    adoRegistro.CursorLocation = adUseClient
    adoRegistro.CursorType = adOpenStatic

    adoComm.CommandText = "SELECT PC.CodParticipe,AP1.DescripParametro TipoIdentidad,PCD.NumIdentidad,DescripParticipe,FechaIngreso,PCD.TipoIdentidad CodIdentidad,PC.TipoMancomuno, AP2.DescripParametro DescripMancomuno " & _
    "FROM ParticipeContratoDetalle PCD JOIN ParticipeContrato PC " & _
    "ON(PCD.CodParticipe=PC.CodParticipe AND PCD.TipoIdentidad='" & strCodTipoDocumento & "' AND PCD.NumIdentidad='" & Trim(txtNumDocumento.Text) & "') " & _
    "JOIN AuxiliarParametro AP1 ON(AP1.CodParametro=PCD.TipoIdentidad AND CodTipoParametro='TIPIDE') " & _
    "JOIN AuxiliarParametro AP2 ON(AP2.CodParametro=PC.TipoMancomuno AND AP2.CodTipoParametro='TIPMAN')"
    adoRegistro.Open adoComm.CommandText, adoConn

    If Not adoRegistro.EOF Then
        If adoRegistro.RecordCount > 1 Then
            gstrFormulario = Me.Name
            frmBusquedaParticipeP.optCriterio(1).Value = vbChecked
            frmBusquedaParticipeP.txtNumDocumento = Trim(txtNumDocumento.Text)
            Call frmBusquedaParticipeP.Buscar
            frmBusquedaParticipeP.Show vbModal
        Else
            gstrCodParticipe = Trim(adoRegistro("CodParticipe"))
            lblDescripTipoParticipe.Caption = Trim(adoRegistro("DescripMancomuno"))
            lblDescripParticipe.Caption = Trim(adoRegistro("DescripParticipe"))
        End If
    End If
    adoRegistro.Close: Set adoRegistro = Nothing
            
End Sub
Private Sub CargarTransferencias()

    Dim adoRegistro         As ADODB.Recordset
    Dim IndTransferente     As Integer, IndTransferido      As Integer
    Dim strSql              As String

    Set adoConsulta2 = New ADODB.Recordset

    IndTransferente = 0
    IndTransferido = 0
    strFechaDesde = Convertyyyymmdd(dtpFechaDesde.Value)
    strFechaHasta = Convertyyyymmdd(DateAdd("d", 1, dtpFechaHasta.Value))
    
    Set adoRegistro = New ADODB.Recordset
    adoRegistro.CursorLocation = adUseClient
    adoRegistro.CursorType = adOpenStatic
     
    With adoComm
        .CommandText = "SELECT NumOperacion FROM ParticipeOperacion " & _
            "WHERE CodFondo = '" & strCodFondo & "'  AND CodAdministradora = '" & gstrCodAdministradora & "' AND " & _
            "(FechaOperacion >= '" & strFechaDesde & "' AND FechaOperacion < '" & strFechaHasta & "') AND " & _
            "CodParticipe = '" & gstrCodParticipe & "' AND TipoOperacion LIKE '" & Codigo_Operacion_Transferencia & "' + '%' "
        adoRegistro.Open .CommandText, adoConn
    
        If Not adoRegistro.EOF Then IndTransferente = 1
        adoRegistro.Close
        
        .CommandText = "SELECT NumOperacion FROM ParticipeCertificado " & _
            "WHERE   CodFondo = '" & strCodFondo & "'  AND CodAdministradora = '" & gstrCodAdministradora & "' AND " & _
            "(FechaOperacion >= '" & strFechaDesde & "' AND FechaOperacion < '" & strFechaHasta & "') AND " & _
            "CodParticipe = '" & gstrCodParticipe & "' AND TipoOperacion LIKE '" & Codigo_Operacion_Transferencia & "' + '%' "
        adoRegistro.Open adoComm.CommandText, adoConn
        
        If Not adoRegistro.EOF Then IndTransferido = 1
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    If IndTransferente = 1 And IndTransferido = 1 Then
        strSql = "SELECT  PO.CodFondo, PO.NumOperacion NumOperacion, TSD.CodCorto TipoOperacion,PO.FechaOperacion FechaOperacion," & _
            "PC.DescripParticipe DescripTransferente,PCT.DescripParticipe DescripTransferido,PO.CantCuotas CantCuotas " & _
            "FROM  ParticipeOperacion PO JOIN ParticipeContrato PC " & _
            "ON(PC.CodParticipe = PO.CodParticipe AND PO.CodFondo = '" & strCodFondo & "' AND PO.CodAdministradora = '" & gstrCodAdministradora & "')" & _
            "JOIN ParticipeCertificado CP ON(CP.CodFondo=PO.CodFondo AND CP.CodAdministradora=PO.CodAdministradora AND CP.NumOperacion=PO.NumOperacion )" & _
            "JOIN ParticipeContrato PCT ON(PCT.CodParticipe = CP.CodParticipe)" & _
            "JOIN TipoSolicitudDetalle TSD ON(TSD.CodDetalleTipoSolicitud = PO.ClaseOperacion AND TSD.CodTipoSolicitud = PO.TipoOperacion)" & _
            "WHERE" & _
            "(PO.FechaOperacion >= '" & strFechaDesde & "'         AND " & _
            "PO.FechaOperacion  <  '" & strFechaHasta & "')        AND " & _
            "PO.TipoOperacion LIKE '" & Codigo_Operacion_Transferencia & "' + '%'                     AND " & _
            "PO.CodParticipe    =  '" & gstrCodParticipe & "'"
    ElseIf IndTransferente = 1 Then
        strSql = "SELECT  PO.CodFondo, PO.NumOperacion NumOperacion, TSD.CodCorto TipoOperacion,PO.FechaOperacion FechaOperacion," & _
            "PC.DescripParticipe DescripTransferente,PCT.DescripParticipe DescripTransferido,PO.CantCuotas CantCuotas " & _
            "FROM  ParticipeOperacion PO JOIN ParticipeContrato PC " & _
            "ON(PC.CodParticipe = PO.CodParticipe AND PO.CodFondo = '" & strCodFondo & "' AND PO.CodAdministradora = '" & gstrCodAdministradora & "')" & _
            "JOIN ParticipeCertificado CP ON(CP.CodFondo=PO.CodFondo AND CP.CodAdministradora=PO.CodAdministradora AND CP.NumOperacion=PO.NumOperacion )" & _
            "JOIN ParticipeContrato PCT ON(PCT.CodParticipe = CP.CodParticipe)" & _
            "JOIN TipoSolicitudDetalle TSD ON(TSD.CodDetalleTipoSolicitud = PO.ClaseOperacion AND TSD.CodTipoSolicitud = PO.TipoOperacion)" & _
            "WHERE" & _
            "(PO.FechaOperacion >= '" & strFechaDesde & "'         AND " & _
            "PO.FechaOperacion  <  '" & strFechaHasta & "')        AND " & _
            "PO.TipoOperacion LIKE '" & Codigo_Operacion_Transferencia & "' + '%'                     AND " & _
            "PO.CodParticipe    =  '" & gstrCodParticipe & "'"
    ElseIf IndTransferido = 1 Then
        strSql = "SELECT  PO.CodFondo, PO.NumOperacion NumOperacion, TSD.CodCorto TipoOperacion,PO.FechaOperacion FechaOperacion," & _
            "PC.DescripParticipe DescripTransferente,PCT.DescripParticipe DescripTransferido,PO.CantCuotas CantCuotas " & _
            "FROM  ParticipeOperacion PO JOIN ParticipeContrato PC " & _
            "ON(PC.CodParticipe = PO.CodParticipe AND PO.CodFondo = '" & strCodFondo & "' AND PO.CodAdministradora = '" & gstrCodAdministradora & "')" & _
            "JOIN ParticipeCertificado CP ON(CP.CodFondo=PO.CodFondo AND CP.CodAdministradora=PO.CodAdministradora AND CP.NumOperacion=PO.NumOperacion )" & _
            "JOIN ParticipeContrato PCT ON(PCT.CodParticipe = CP.CodParticipe)" & _
            "JOIN TipoSolicitudDetalle TSD ON(TSD.CodDetalleTipoSolicitud = PO.ClaseOperacion AND TSD.CodTipoSolicitud = PO.TipoOperacion)" & _
            "WHERE" & _
            "(PO.FechaOperacion >= '" & strFechaDesde & "'         AND " & _
            "PO.FechaOperacion  <  '" & strFechaHasta & "')        AND " & _
            "PO.CodParticipe    =  PC.CodParticipe                 AND " & _
            "PO.TipoOperacion LIKE '" & Codigo_Operacion_Transferencia & "' + '%'                     AND " & _
            "CP.CodParticipe    =  '" & gstrCodParticipe & "'"
    End If
        
    If strSql <> Valor_Caracter Then
        strEstado = Reg_Defecto
        With adoConsulta2
            .ActiveConnection = gstrConnectConsulta
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockBatchOptimistic
            .Open strSql
        End With
            
        tdgConsulta2.DataSource = adoConsulta2
        
        If adoConsulta2.RecordCount > 0 Then strEstado = Reg_Consulta
    End If
    
End Sub
Private Function TodoOK() As Boolean
                
    Dim adoConsulta As ADODB.Recordset
    Dim strMensaje  As String
    
    TodoOK = False
                
                                
'    Set adoConsulta = New ADODB.Recordset
'    '*** Se Realizó Cierre anteriormente ? ***
'    adoComm.CommandText = "{ call up_GNValidaCierreRealizado('" & _
'        strCodFondo & "','" & gstrCodAdministradora & "','" & Convertyyyymmdd(dtpFechaHasta.Value) & "','" & _
'        Convertyyyymmdd(DateAdd("d", 1, dtpFechaHasta.Value)) & "') }"
'    Set adoConsulta = adoComm.Execute
'
'    If Not adoConsulta.EOF Then
'        If Trim(adoConsulta("IndCierre")) = Valor_Caracter Then
'            MsgBox "El Cierre Diario del Día " & CStr(dtpFechaHasta.Value) & " no ha sido realizado aún.", vbCritical, Me.Caption
'            adoConsulta.Close: Set adoConsulta = Nothing
'            Exit Function
'        End If
'    End If
'    adoConsulta.Close

    '*** Si todo paso OK ***
    TodoOK = True
  
End Function

