VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmListaReportes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reportes"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6780
   ScaleWidth      =   6720
   Begin MSAdodcLib.Adodc adoReporte 
      Height          =   330
      Left            =   3630
      Top             =   6000
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5040
      Picture         =   "frmListaReportes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5940
      Width           =   1200
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   1800
      Picture         =   "frmListaReportes.frx":0582
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5940
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Vista &Previa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   360
      Picture         =   "frmListaReportes.frx":0BF3
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5940
      Width           =   1200
   End
   Begin VB.Frame fraReportes 
      Caption         =   "Parámetros"
      Height          =   2175
      Index           =   0
      Left            =   45
      TabIndex        =   0
      Top             =   120
      Width           =   6585
      Begin VB.CheckBox chkPersonalizado 
         Caption         =   "Personalizado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3450
         TabIndex        =   18
         Top             =   1800
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.CheckBox chkBalanceEleccion 
         Caption         =   "Check1"
         Height          =   255
         Left            =   1920
         TabIndex        =   15
         Top             =   1800
         Width           =   255
      End
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1040
         Width           =   4095
      End
      Begin VB.CheckBox chkSimulacion 
         Caption         =   "Simulación"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   300
         TabIndex        =   12
         Top             =   1800
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ComboBox cboTipoCartera 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1380
         Width           =   4095
      End
      Begin VB.ComboBox cboFondo 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   4095
      End
      Begin MSComCtl2.DTPicker dtpFechaDesde 
         Height          =   285
         Left            =   1800
         TabIndex        =   4
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   50135041
         CurrentDate     =   38779
      End
      Begin MSComCtl2.DTPicker dtpFechaHasta 
         Height          =   285
         Left            =   4560
         TabIndex        =   6
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   50135041
         CurrentDate     =   38779
      End
      Begin VB.Label lblBalanceEleccion 
         Caption         =   "Consolidado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   16
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cartera"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   13
         Top             =   1400
         Width           =   870
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         Caption         =   "Moneda Reporte"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   11
         Top             =   1065
         Width           =   1200
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         Caption         =   "Fondo"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   380
         Width           =   450
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicial"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   740
         Width           =   900
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   3360
         TabIndex        =   7
         Top             =   740
         Width           =   825
      End
   End
   Begin TrueOleDBGrid60.TDBGrid tdgReporte 
      Bindings        =   "frmListaReportes.frx":10DC
      Height          =   3375
      Left            =   60
      OleObjectBlob   =   "frmListaReportes.frx":10F5
      TabIndex        =   17
      Top             =   2400
      Width           =   6555
   End
End
Attribute VB_Name = "frmListaReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrFondo()          As String, arrTipoCartera()         As String
Dim arrReporte()        As String, arrMoneda()              As String
Dim strCodFondo         As String, strCodTipoCartera        As String
Dim strCodReporte       As String, strDescripMoneda         As String
Dim strCodModuloL       As String, strCodGrupoReporteL      As String
Dim strSQL              As String, strCodMoneda             As String
Dim dblTipoCambio       As Double

Public Sub Adicionar()

End Sub

Public Sub Buscar()
            
    If chkSimulacion.Enabled Then
        If chkSimulacion.Value Then
            
            If chkPersonalizado.Value Then chkPersonalizado.Value = False
            
            strSQL = "SELECT CodReporte,DescripReporte,IndRango,IndProcedimiento,NumParamFormulas," & _
                "NumParamProcedimiento,IndTipoFondo,IndOpciones,CodReporteAlterno,IndMoneda,IndPersonalizado FROM ControlReporte " & _
                "WHERE SUBSTRING(GrupoReporte,PATINDEX('%" & strCodGrupoReporteL & "%',GrupoReporte),1)='" & strCodGrupoReporteL & "' AND " & _
                "SUBSTRING(PerfilReporte,PATINDEX('%" & strCodModuloL & "%',PerfilReporte),1)='" & strCodModuloL & "' AND " & _
                "IndVigente='X' And IndSimulacion='X' " & _
                "ORDER BY DescripReporte"
        Else
            strSQL = "SELECT CodReporte,DescripReporte,IndRango,IndProcedimiento,NumParamFormulas," & _
                "NumParamProcedimiento,IndTipoFondo,IndOpciones,CodReporteAlterno,IndMoneda,IndPersonalizado FROM ControlReporte " & _
                "WHERE SUBSTRING(GrupoReporte,PATINDEX('%" & strCodGrupoReporteL & "%',GrupoReporte),1)='" & strCodGrupoReporteL & "' AND " & _
                "SUBSTRING(PerfilReporte,PATINDEX('%" & strCodModuloL & "%',PerfilReporte),1)='" & strCodModuloL & "' AND " & _
                "IndVigente='X' " & _
                "ORDER BY DescripReporte"
        End If
    End If

    With adoReporte
        .ConnectionString = gstrConnectConsulta
        .RecordSource = strSQL
        .Refresh
    End With
    
    tdgReporte.Refresh
    
End Sub

Public Sub CargarReportes(strCodModulo, strGrupoReporte)

    Dim adoRegistro     As ADODB.Recordset
    
    strCodModuloL = strCodModulo: strCodGrupoReporteL = strGrupoReporte
    If strCodModulo = "C" Then
        lblDescrip(4).Caption = "Tipo Instrumento"
        '*** Cartera del Fondo ***
        strSQL = "SELECT CodFile CODIGO,DescripFile DESCRIP FROM InversionFile WHERE CodEstructura<>'' AND IndVigente='X' ORDER BY DescripFile"
        CargarControlLista strSQL, cboTipoCartera, arrTipoCartera(), Sel_Defecto
    Else
        lblDescrip(4).Caption = "Tipo Cartera"
        '*** Cartera del Fondo ***
        strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPCAR' ORDER BY DescripParametro"
        CargarControlLista strSQL, cboTipoCartera, arrTipoCartera(), Sel_Defecto
    End If
    
    If cboTipoCartera.ListCount > 0 Then cboTipoCartera.ListIndex = 0
    
    Set adoRegistro = New ADODB.Recordset
    With adoComm
        
        .CommandText = "SELECT DescripParametro FROM AuxiliarParametro " & _
            "WHERE CodTipoParametro='GRPREP' AND ValorParametro='" & strGrupoReporte & "'"
                    
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            frmListaReportes.Caption = Trim(adoRegistro("DescripParametro"))
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
            
    strSQL = "SELECT CodReporte,DescripReporte,IndRango,IndProcedimiento,NumParamFormulas," & _
        "NumParamProcedimiento,IndTipoFondo,IndOpciones,CodReporteAlterno,IndPersonalizado FROM ControlReporte " & _
        "WHERE SUBSTRING(GrupoReporte,PATINDEX('%" & strGrupoReporte & "%',GrupoReporte),1)='" & strGrupoReporte & "' AND " & _
        "SUBSTRING(PerfilReporte,PATINDEX('%" & strCodModulo & "%',PerfilReporte),1)='" & strCodModulo & "' AND " & _
        "IndVigente='X' " & _
        "ORDER BY DescripReporte"
        
    With adoReporte
        .ConnectionString = gstrConnectConsulta
        .RecordSource = strSQL
        .Refresh
    End With
    
    tdgReporte.Refresh

End Sub

Public Sub Eliminar()

End Sub

Public Sub Grabar()

End Sub


Public Sub Imprimir()

End Sub

Public Sub Modificar()

End Sub

Private Sub cboFondo_Click()

    Dim adoRegistro As ADODB.Recordset
    
    strCodFondo = Valor_Caracter
    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(arrFondo(cboFondo.ListIndex))
    
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
    
        '*** Monedas contables del fondo ***
        strSQL = "{ call up_ACSelDatosParametro('70','" & strCodFondo & "','" & gstrCodAdministradora & "') }"
        CargarControlLista strSQL, cboMoneda, arrMoneda(), Valor_Caracter
        If cboMoneda.ListCount > 0 Then cboMoneda.ListIndex = 0
    
        '*** Fecha Vigente, Valor Cuota, Valor Cuota T-1, Moneda y Cantidad Inicial de Cuotas del Fondo ***
        .CommandText = "{ call up_ACSelDatosParametro(23,'" & strCodFondo & "','" & gstrCodAdministradora & "') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            gdatFechaActual = adoRegistro("FechaCuota")
            gdblTipoCambio = adoRegistro("ValorTipoCambio")
            gstrCodMoneda = adoRegistro("CodMoneda")
            dtpFechaDesde.Value = DateAdd("d", -1, gdatFechaActual)
            dtpFechaHasta.Value = dtpFechaDesde.Value
                        
            frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
End Sub

Private Sub cboMoneda_Click()

    strCodMoneda = Valor_Caracter
    If cboMoneda.ListIndex < 0 Then Exit Sub
    
    strCodMoneda = Trim(arrMoneda(cboMoneda.ListIndex))
    
End Sub

Private Sub cboTipoCartera_Click()

    strCodTipoCartera = Valor_Caracter
    If cboTipoCartera.ListIndex < 0 Then Exit Sub
        
    strCodTipoCartera = Trim(arrTipoCartera(cboTipoCartera.ListIndex))
        
End Sub

Private Sub chkPersonalizado_Click()
    If chkPersonalizado.Enabled Then
        If chkPersonalizado.Value Then
        
            If chkSimulacion.Value Then chkSimulacion.Value = False
        
            strSQL = "SELECT CodReporte,DescripReporte,IndRango,IndProcedimiento,NumParamFormulas," & _
                "NumParamProcedimiento,IndTipoFondo,IndOpciones,CodReporteAlterno,IndMoneda,IndPersonalizado FROM ControlReporte " & _
                "WHERE SUBSTRING(GrupoReporte,PATINDEX('%" & strCodGrupoReporteL & "%',GrupoReporte),1)='" & strCodGrupoReporteL & "' AND " & _
                "SUBSTRING(PerfilReporte,PATINDEX('%" & strCodModuloL & "%',PerfilReporte),1)='" & strCodModuloL & "' AND " & _
                "IndVigente='X' And IndPersonalizado='X' " & _
                "ORDER BY DescripReporte"
                
            
        Else
            strSQL = "SELECT CodReporte,DescripReporte,IndRango,IndProcedimiento,NumParamFormulas," & _
                "NumParamProcedimiento,IndTipoFondo,IndOpciones,CodReporteAlterno,IndMoneda,IndPersonalizado FROM ControlReporte " & _
                "WHERE SUBSTRING(GrupoReporte,PATINDEX('%" & strCodGrupoReporteL & "%',GrupoReporte),1)='" & strCodGrupoReporteL & "' AND " & _
                "SUBSTRING(PerfilReporte,PATINDEX('%" & strCodModuloL & "%',PerfilReporte),1)='" & strCodModuloL & "' AND " & _
                "IndVigente='X' " & _
                "ORDER BY DescripReporte"
                
        End If
    End If
   
    With adoReporte
        .ConnectionString = gstrConnectConsulta
        .RecordSource = strSQL
        .Refresh
    End With
   
    tdgReporte.Refresh
End Sub

Private Sub chkSimulacion_Click()

    Call Buscar
    
End Sub



Private Sub cmdImprimir_Click(Index As Integer)

    Dim intCont As Integer
    
    If tdgReporte.SelBookmarks.Count = 0 Then
        
        MsgBox "Debe seleccionar un registro...", vbCritical, Me.Caption
        Exit Sub
        
    End If
    
    Select Case Index
        Case 0, 1
            If Trim(tdgReporte.Columns(6).Value) = Valor_Indicador Then
                If cboFondo.ListIndex < 0 Then
                    MsgBox "Seleccione Fondo", vbCritical, Me.Caption
                    Exit Sub
                End If
            Else
                If cboFondo.ListIndex < 0 Then
                    MsgBox "Seleccione Fondo", vbCritical, Me.Caption
                    Exit Sub
                End If
            End If
                        
            If cboTipoCartera.Enabled Then
                If cboTipoCartera.ListIndex <= 0 Then
                    MsgBox "Seleccione el" & Space(1) & Trim(lblDescrip(4).Caption), vbCritical, Me.Caption
                    Exit Sub
                End If
            End If
                            
            '*** Control del Reporte ***
            CtrlReporte Index
                    
    End Select
    
End Sub

Private Sub cmdSalir_Click()

    Unload Me
    
End Sub

Private Sub dtpFechaHasta_Change()

    If Not dtpFechaDesde.Enabled Then
        dtpFechaDesde.Value = dtpFechaHasta.Value
    End If
    
End Sub

Private Sub Form_Activate()

    If gstrCodAdministradoraContable = Valor_Caracter And strCodFondo = Valor_Caracter Then
        MsgBox "No Existe Empresa de Trabajo Disponible", vbCritical, Me.Caption
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    Call InicializarValores
    Call CargarListas
    Call DarFormato
    
    
    
    
    Call ValidarPermisoUsoControl(Trim(gstrLoginUS), Me, Trim(App.Title) + Separador_Codigo_Objeto + _
    gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)
    
    
    
    CentrarForm Me
        
    Call ValidaExisteTipoCambio(Codigo_TipoCambio_SBS, gstrFechaActual)
    
End Sub

Private Sub CargarListas()
    
'    Dim intRegistro         As Integer
    
'    strSQL = "{ call up_ACSelDatos(70) }"
'    CargarControlLista strSQL, cboMoneda, arrMoneda(), Valor_Caracter
'
'    intRegistro = ObtenerItemLista(arrMoneda(), Codigo_Moneda_Local)
'    If intRegistro >= 0 Then cboMoneda.ListIndex = intRegistro
    
    If gstrCodAdministradoraContable = Valor_Caracter Then
        '*** Fondos ***
        strSQL = "{ call up_ACSelDatosParametro(74,'" & gstrCodAdministradora & "','" & gstrCodFondoContable & "') }"
'        strSQL = "{ call up_ACSelDatos(8) }"
        CargarControlLista strSQL, cboFondo, arrFondo(), Valor_Caracter
    Else
        strCodFondo = "000"
        cboFondo.AddItem Trim(frmMainMdi.txtEmpresa.Text)
        ReDim arrFondo(0)
        arrFondo(0) = "000"
        chkSimulacion.Visible = False
        lblDescrip(4).Visible = False
        cboTipoCartera.Visible = False
    End If
    
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
        
End Sub
Private Sub InicializarValores()
        
    '*** Valores Iniciales ***
    dtpFechaDesde.Value = gdatFechaActual
    dtpFechaHasta.Value = gdatFechaActual
    cboTipoCartera.Enabled = False
    
    
    
    If gstrCodAdministradoraContable <> Valor_Caracter Then lblDescrip(0).Caption = "Administradora"
        
End Sub
Private Sub DarFormato()

    Dim intCont As Integer
    Dim elemento As Object
    
    For intCont = 0 To (lblDescrip.Count - 1)
        Call FormatoEtiqueta(lblDescrip(intCont), vbLeftJustify)
    Next
    
    For intCont = 0 To (fraReportes.Count - 1)
        Call FormatoMarco(fraReportes(intCont))
    Next
    
    For Each elemento In Me.Controls
    
        If TypeOf elemento Is TDBGrid Then
            Call FormatoGrilla(elemento)
        End If
    
    Next
            
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmListaReportes = Nothing
    
End Sub

Private Sub dtpFechaDesde_Change()

    If ValidarFechaInicial(dtpFechaDesde.Value, strCodFondo, gstrCodAdministradora) Then
        dtpFechaHasta = dtpFechaDesde
    Else
        dtpFechaDesde.Value = DateAdd("d", 1, dtpFechaDesde.Value)
    End If
    
End Sub
Private Sub CtrlReporte(Index As Integer)
    
    Dim adoConsulta             As ADODB.Recordset
    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    Dim strFechaDesde           As String, strFechaHasta        As String
    Dim strIndProcedimiento     As String, strIndRango          As String
    Dim strCodGrupoReporte      As String, strIndTipoFondo      As String
    Dim intNumProcedimiento     As Integer, intNumFormula       As Integer
    Dim MonedaVaria As String
    Dim strValor As String
    Dim strCuenta As String
    Dim intCant As Integer
    'Dim hDC As Long, hBitmap As Long

    On Error GoTo Ctrl_Error
    
    'Load the bitmap into the memory
    'hBitmap = LoadImage(App.hInstance, "C:\\Spectrum Fondos\\Celfin\\Fuentes\\Imagenes\\LogoEmpresa.jpg", IMAGE_BITMAP, 320, 200, LR_LOADFROMFILE)
                
    If dtpFechaDesde.Enabled Then
        strFechaDesde = Convertyyyymmdd(dtpFechaDesde.Value)
        
        '*** BMM NUEVOS CAMBIOS ***
        If strCodReporte = "EstadoResultadoP" Then
            strFechaHasta = Convertyyyymmdd(dtpFechaHasta.Value)
        Else
            'strFechaHasta = Convertyyyymmdd(DateAdd("d", 1, dtpFechaHasta.Value)) --'del form de la siv
            strFechaHasta = Convertyyyymmdd(dtpFechaHasta.Value) '--original del form
        End If
        
        '*******
        
    Else
        strFechaDesde = Convertyyyymmdd(dtpFechaHasta.Value)
        strFechaHasta = Convertyyyymmdd(DateAdd("d", 1, dtpFechaHasta.Value))
    End If
    
    Set adoConsulta = New ADODB.Recordset
    With adoComm
        .CommandText = "SELECT PerfilReporte,GrupoReporte,IndProcedimiento,IndRango,NumParamFormulas,NumParamProcedimiento,IndTipoFondo FROM ControlReporte WHERE CodReporte='" & Trim(tdgReporte.Columns(0).Value) & "'"
        Set adoConsulta = .Execute

        If Not adoConsulta.EOF Then
            strIndProcedimiento = Trim(adoConsulta("IndProcedimiento")): strCodGrupoReporte = Trim(adoConsulta("GrupoReporte"))
            strIndRango = Trim(adoConsulta("IndRango")): intNumFormula = CInt(adoConsulta("NumParamFormulas"))
            intNumProcedimiento = CInt(adoConsulta("NumParamProcedimiento")): strIndTipoFondo = Trim(adoConsulta("IndTipoFondo"))
        End If
        adoConsulta.Close

'        If strCodFondo = "00" Then
'            If strIndTipoFondo = Valor_Caracter Then
'                MsgBox "Opción inválida para este reporte", vbCritical, Me.Caption
'                cboFondo.ListIndex = -1
'                Exit Sub
'            End If
'        End If
        
    End With
    
    Dim strnameRepoAux As String
    
    gstrNameRepo = strCodReporte
    strnameRepoAux = gstrNameRepo
    
    If chkPersonalizado.Value Then
        gstrNameRepo = strCodReporte + "P"
    End If
                
    Set frmReporte = New frmVisorReporte

    ReDim aReportParamS(intNumProcedimiento - 1) 'acr --6
    ReDim aReportParamFn(7)
    ReDim aReportParamF(7)
    
    '*** Parámetros mínimos ***
    aReportParamFn(0) = "Usuario"
    aReportParamFn(1) = "Hora"
    aReportParamFn(2) = "NombreEmpresa"
    aReportParamFn(3) = "Fondo"
    aReportParamFn(4) = "FechaDesde"
    aReportParamFn(5) = "FechaHasta"
    aReportParamFn(6) = "TipoCambio"
    aReportParamFn(7) = "Moneda"
    
    aReportParamF(0) = gstrLogin
    aReportParamF(1) = Format(Time(), "hh:mm:ss")
    aReportParamF(2) = gstrNombreEmpresa & Space(1)
    aReportParamF(3) = Trim(cboFondo.Text)
    aReportParamF(4) = CStr(dtpFechaDesde.Value)
    aReportParamF(5) = CStr(dtpFechaHasta.Value)
    aReportParamF(6) = gdblTipoCambio
    aReportParamF(7) = strDescripMoneda
    'If intNumFormula = 7 Then
        'If strCodMoneda = "01" Then
        'MonedaVaria = "02"
        'gstrNameRepo = "RegistroComprasParte1"
        'End If
        'If strCodMoneda = "02" Then
        'MonedaVaria = "01"
        'gstrNameRepo = "RegistroComprasParte2"
        'End If
        'strCodMoneda = MonedaVaria
    'End If
                    
    If intNumProcedimiento >= 1 Then
        aReportParamS(0) = strCodFondo
    End If
    If intNumProcedimiento >= 2 Then
    aReportParamS(1) = gstrCodAdministradora
    End If
    If intNumProcedimiento >= 3 Then
    aReportParamS(2) = strFechaDesde
    End If
    If intNumProcedimiento >= 4 Then
    aReportParamS(3) = strFechaHasta
    End If
    If intNumProcedimiento >= 5 Then
    aReportParamS(4) = strCodMoneda
    End If
    
    If gstrNameRepo <> "LibroBancos2" And gstrNameRepo <> "LimiteInstrumento" And gstrNameRepo <> "BalanceComprobacionSim" _
    And gstrNameRepo <> "BalanceComprobacion" And gstrNameRepo <> "LimiteMercado" And gstrNameRepo <> "LimiteMoneda" _
    And gstrNameRepo <> "LimitePlazo" And gstrNameRepo <> "SaldosDiarios" Then
        ReDim Preserve aReportParamS(6)
        aReportParamS(5) = gstrCodClaseTipoCambioFondo
        aReportParamS(6) = gstrValorTipoCambioCierre
    End If
    
    'If gstrNameRepo = "LibroBancos" Then
'        aReportParamS(2) = strFechaDesde 'CStr(dtpFechaDesde.Value)
'        aReportParamS(3) = strFechaHasta 'CStr(dtpFechaHasta.Value)
'        aReportParamS(4) = "%" 'strCodMoneda
'        aReportParamS(5) = "%" 'gstrCodClaseTipoCambioFondo
'        aReportParamS(6) = "%" 'gstrValorTipoCambioCierre
'    End If
    
    If gstrNameRepo = "BalanceComprobacionSim" Or gstrNameRepo = "BalanceComprobacion" Then
        'ReDim Preserve aReportParamS(7)
        aReportParamS(5) = "%" 'gstrCodClaseTipoCambioFondo
        aReportParamS(6) = "0" 'gstrValorTipoCambioCierre
        aReportParamS(7) = 8 'gstrValorTipoCambioCierre
    End If
   
    
    '*** Parámetros mínimos ***
    MonedaVaria = ""
    If intNumFormula = 8 Then
        ReDim Preserve aReportParamFn(8)
        ReDim Preserve aReportParamF(8)
        
        aReportParamFn(8) = "DescripInstrumento"
        aReportParamF(8) = Trim(cboTipoCartera.Text)
        
    End If
    
    
'    ***   ANTIGUAS  LINEAS ***
'    If intNumProcedimiento = 8 And gstrNameRepo <> "BalanceComprobacionSim" And gstrNameRepo <> "BalanceComprobacion" Then  ' Antes 6 -> corregir en tabla
'        ReDim Preserve aReportParamS(7)
'
'        If cboTipoCartera.Enabled Then
'            aReportParamS(7) = strCodTipoCartera
'        Else
'            aReportParamS(3) = Convertyyyymmdd(dtpFechaHasta.Value)
'            aReportParamS(4) = Convertyyyymmdd(DateAdd("d", 1, dtpFechaHasta.Value)) 'strFechaHasta
'            aReportParamS(5) = strCodMoneda
'            aReportParamS(6) = gstrCodClaseTipoCambioFondo
'            aReportParamS(7) = gstrValorTipoCambioCierre
'        End If
'    End If
    
'   ********************
    
    '********************BMM CAMBIOS REALIZADOS EN CONDICION*********'
    'If intNumProcedimiento = 8 And gstrNameRepo <> "BalanceComprobacionSim" And gstrNameRepo <> "BalanceComprobacion" Then  ' Antes 6 -> corregir en tabla
    If intNumProcedimiento = 8 And gstrNameRepo <> "BalanceComprobacionSim" And gstrNameRepo <> "BalanceComprobacion" And gstrNameRepo <> "EstadoResultadoP" And gstrNameRepo <> "CierreContableP" And gstrNameRepo <> "SaldosDiarios" Then ' Antes 6 -> corregir en tabla
        
        ReDim Preserve aReportParamS(7)
        
        If cboTipoCartera.Enabled Then
            aReportParamS(7) = strCodTipoCartera
        Else
            aReportParamS(3) = Convertyyyymmdd(dtpFechaHasta.Value)
            aReportParamS(4) = Convertyyyymmdd(DateAdd("d", 1, dtpFechaHasta.Value)) 'strFechaHasta
            aReportParamS(5) = strCodMoneda
            aReportParamS(6) = gstrCodClaseTipoCambioFondo
            aReportParamS(7) = gstrValorTipoCambioCierre
        End If
    End If
    
    '********************BMM CAMBIOS REALIZADOS EN CONDICION*********'
    If intNumProcedimiento = 8 And gstrNameRepo = "EstadoResultadoP" Then
    'ahora ya no tiene 8 parametros si no 9
        ReDim Preserve aReportParamS(8)
        aReportParamS(7) = strnameRepoAux
        aReportParamS(8) = "001"
    End If
    
    If intNumProcedimiento = 7 And gstrNameRepo = "CierreContableP" Then
    'ahora ya no tiene 8 parametros si no 9
        ReDim Preserve aReportParamS(8)
        aReportParamS(7) = strnameRepoAux
        aReportParamS(8) = "002"
    End If
    
    
    '*****************************************************************
    
    
    If intNumProcedimiento = 9 Then ' Antes 7 -> Corregir en tabla
        ReDim Preserve aReportParamS(8)
        
        aReportParamS(7) = Codigo_Listar_Todos
        aReportParamS(8) = Valor_Comodin
        
        If gstrNameRepo = "HistLibroMayor1" Then
           strValor = MsgBox("¿Desea consultar por cuenta?", vbYesNo, Me.Caption)
            If strValor = 6 Then
                aReportParamS(7) = Codigo_Listar_Todos
                strCuenta = InputBox("Ingrese el Numero de Cuenta de 6 digitos", Me.Caption)
                'intCant = Len(strCuenta)
                If strCuenta = "" Then
                    MsgBox "Cuenta en Blanco", vbInformation, Me.Caption
                    Exit Sub
                Else
                aReportParamS(8) = strCuenta
                End If
            End If
            If gstrNameRepo = "HistLibroMayor1" Then
                Call GenerarLibroMayor(strCodFondo, strCodMoneda, dtpFechaDesde.Value, dtpFechaHasta.Value)
            End If
       
        End If
    End If
    
    If gstrNameRepo = "BalanceComprobacion" And chkBalanceEleccion.Value And chkSimulacion.Value = False Then
        gstrNameRepo = "BalanceComprobacionConsolidado"
    End If
                                               
    If gstrNameRepo = "BalanceComprobacionSim" And chkBalanceEleccion.Value And chkSimulacion.Value Then
        gstrNameRepo = "BalanceComprobacionConsolidadoSim"
    End If
       
'    If gstrNameRepo = "NominaParticipes" Then
'        strValor = MsgBox("¿Desea agruparlo por partícipe?", vbYesNo, Me.Caption)
'        If strValor = 6 Then
'            gstrNameRepo = "NominaParticipesAgrupado"
'        End If
'    End If
                                          
    If gstrNameRepo = "InversionOperacion" Or gstrNameRepo = "FinanciamientoOperacion" Then
        ReDim Preserve aReportParamS(4)
    End If
                                          
    If gstrNameRepo = "OperaPendientes" Or gstrNameRepo = "InversionValorizacion" _
    Or gstrNameRepo = "FinanciamientosPendientes" Or gstrNameRepo = "FinanciamientoValorizacion" Then
        ReDim Preserve aReportParamS(2)
        aReportParamS(2) = Convertyyyymmdd(dtpFechaHasta.Value)
    End If
                                          
    If gstrNameRepo = "InversionValorizacion" Or gstrNameRepo = "FinanciamientoValorizacion" Then
        ReDim Preserve aReportParamS(3)
        aReportParamS(2) = Convertyyyymmdd(dtpFechaDesde.Value)
        aReportParamS(3) = Convertyyyymmdd(dtpFechaHasta.Value)
    End If
    
    '++REA 2015-07-20
    If gstrNameRepo = "FactPendiente" Or gstrNameRepo = "ComisionesGastos" Then
        ReDim Preserve aReportParamS(4)
        aReportParamS(4) = strCodMoneda
    End If
    '--REA 2015-07-20
    
    If gstrNameRepo = "CapitalInvertido" Then
        ReDim Preserve aReportParamS(4)
        aReportParamS(4) = gstrFechaActual
    End If
                              
    If gstrNameRepo = "OperaPorVencer" Then
        ReDim Preserve aReportParamS(3)
    End If
          
    If gstrNameRepo = "OperaVencidas" Then
        ReDim Preserve aReportParamS(2)
    End If
    
    If gstrNameRepo = "CambiosPatrimonio" Then
        ReDim Preserve aReportParamS(4)
        aReportParamS(4) = strCodMoneda
    End If
    
    Dim frmFiltro As New frmFiltroReporte2
        
    If gstrNameRepo = "HistOperaciones" Then
    ReDim Preserve aReportParamS(3)
        aReportParamS(2) = Convertyyyymmdd(CStr(dtpFechaHasta.Value))
        
        frmFiltro.strReporte = gstrNameRepo
        frmFiltro.Show 1
        
        If frmFiltro.blnCancelado Then
            Exit Sub
        End If
        
        aReportParamS(3) = frmFiltro.strCodEmisor
        
    End If
    
    If gstrNameRepo = "OperacionesPorVencer" Then
        ReDim Preserve aReportParamS(3)
        aReportParamS(2) = Convertyyyymmdd(CStr(dtpFechaHasta.Value))
    
        frmFiltro.strReporte = gstrNameRepo
        frmFiltro.Show 1
        
        If frmFiltro.blnCancelado Then
            Exit Sub
        End If
        
        aReportParamS(3) = frmFiltro.intDias
        
    End If
              
    gstrSelFrml = Valor_Caracter
    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"

    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())

    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
    frmReporte.Show vbModal

    Set frmReporte = Nothing

    Screen.MousePointer = vbNormal
    Exit Sub
    
Ctrl_Error:
    With err
        Select Case .Number
            Case 20504
                MsgBox "! Reporte NO EXISTE !" & vbNewLine & _
                        "CONSULTE con el Administrador del Sistema.", vbCritical, Me.Caption
            Case Else
                MsgBox " Error " & Str(.Number) & " -  " & .Description, vbCritical, Me.Caption
        End Select
    End With
    Me.MousePointer = vbDefault
    Resume Next
    
    
'        If strCodFondo <> "00" Then
'            .CommandText = "SELECT VAL_TCMB,VAL_PATR,VAL_ACTI,COD_MONE FROM FMCUOTAS WHERE FCH_CUOT='" & strFchInicio & "' AND COD_FOND='" & strCodFondo & "'"
'        Else
'            Select Case strTipFon
'                Case "T"
'                    .CommandText = "SELECT VAL_TCMB,VAL_PATR,VAL_ACTI,COD_MONE FROM FMCUOTAS WHERE FCH_CUOT='" & strFchInicio & "'"
'                Case "F"
'                    .CommandText = "SELECT VAL_TCMB,VAL_PATR,VAL_ACTI,COD_MONE FROM FMCUOTAS WHERE FCH_CUOT='" & strFchInicio & "' AND COD_FOND IN (SELECT COD_FOND FROM FMFONDOS WHERE TIP_FOND='F')"
'                Case "V"
'                    .CommandText = "SELECT VAL_TCMB,VAL_PATR,VAL_ACTI,COD_MONE FROM FMCUOTAS WHERE FCH_CUOT='" & strFchInicio & "' AND COD_FOND IN (SELECT COD_FOND FROM FMFONDOS WHERE TIP_FOND='V')"
'            End Select
'        End If
'        Set adoConsulta = .Execute
'
'        Do While Not adoConsulta.EOF
'            dblTC = CDbl(adoConsulta("VAL_TCMB"))
'            curPatr = curPatr + Abs(CCur(adoConsulta("VAL_PATR")))
'            curActivo = curActivo + Abs(CCur(adoConsulta("VAL_ACTI")))
'            strCodMone = CStr(adoConsulta("COD_MONE"))
'
'            adoConsulta.MoveNext
'        Loop
'        adoConsulta.Close
'
'        dblTasInte = 0: dblTasa = 0: curVanActu = 0: dblTasPdia = 0
'        If strCodFondo <> "00" Then
'            .CommandText = "SELECT SUM(VAN_ACTU) VAN_ACTU,SUM(TAS_PDIA) TAS_PDIA FROM tblCarteraInv WHERE FCH_CART='" & strFchInicio & "' AND COD_FOND='" & strCodFondo & "'"
'        Else
'            Select Case strTipFon
'                Case "T"
'                    .CommandText = "SELECT SUM(VAN_ACTU) VAN_ACTU,SUM(TAS_PDIA) TAS_PDIA FROM tblCarteraInv WHERE FCH_CART='" & strFchInicio & "'"
'                Case "F"
'                    .CommandText = "SELECT SUM(VAN_ACTU) VAN_ACTU,SUM(TAS_PDIA) TAS_PDIA FROM tblCarteraInv WHERE FCH_CART='" & strFchInicio & "' AND COD_FOND IN (SELECT COD_FOND FROM FMFONDOS WHERE TIP_FOND='F')"
'                Case "V"
'                    .CommandText = "SELECT SUM(VAN_ACTU) VAN_ACTU,SUM(TAS_PDIA) TAS_PDIA FROM tblCarteraInv WHERE FCH_CART='" & strFchInicio & "' AND COD_FOND IN (SELECT COD_FOND FROM FMFONDOS WHERE TIP_FOND='V')"
'            End Select
'        End If
'        Set adoConsulta = .Execute
'
'        If Not adoConsulta.EOF Then
'            If IsNull(adoConsulta("VAN_ACTU")) Then
'                curVanActu = 0
'            Else
'                curVanActu = CCur(adoConsulta("VAN_ACTU"))
'            End If
'        End If
'        adoConsulta.Close
'
'        .CommandText = "SELECT TAS_PDIA FROM tblCarteraInv WHERE FCH_CART='" & strFchInicio & "' AND COD_FOND='" & strCodFondo & "' "
'        .CommandText = .CommandText & "GROUP BY TAS_PDIA ORDER BY TAS_PDIA"
'        Set adoConsulta = .Execute
'
'        Do While Not adoConsulta.EOF
'            dblTasPdia = dblTasPdia + adoConsulta("TAS_PDIA")
'
'            adoConsulta.MoveNext
'        Loop
'        adoConsulta.Close: Set adoConsulta = Nothing
'
'    End With

'    With gobjReport
'        .Connect = gstrConnectReport
'        .ReportFileName = gstrRptPath & Trim(arrReportes(intRep)) & ".RPT"
'
'        .Formulas(0) = "User='" & gstrUID & "'"
'        .Formulas(1) = "FchProcIni='" & dtpFechaDesde.Text & "'"
'        .Formulas(2) = "DscFondo='" & Trim(cboFondo.Text) & "'"
'        .Formulas(3) = "Hora='" & Format(Time, "hh:mm") & "'"
'        .Formulas(4) = "TipCamb=" & dblTC
'        .Formulas(5) = "Moneda='" & strCodMone & "'"
'
'        Select Case strCodClaseReporte
'            Case "V" '*** Varios ***
'                .Formulas(0) = "Fondo='" & Trim$(cboFondo.Text) & "'"
'
'                If strIndProcedimiento = "X" Then
'                    .StoredProcParam(0) = strFchInicio
'                    .StoredProcParam(1) = strCodFondo
'                End If
'
'            Case "A" '*** Análisis ***
'
'                Select Case intNumFormula
'                    Case 7
'                        curPatr = curPatr / dblTC
'                        .Formulas(6) = "ValPatr=" & curPatr
'
'                    Case 8
'                        If strCodFondo <> "00" Then
'                            If strCodMone = "D" Then
'                                curPatr = curPatr / dblTC
'                                curActivo = curActivo / dblTC
'                            End If
'                        Else
'                            strCodMone = "D"
'                            curPatr = curPatr / dblTC
'                            curActivo = curActivo / dblTC
'                        End If
'                        .Formulas(6) = "ValPatr=" & curPatr
'                        .Formulas(7) = "Activo=" & curActivo
'
'                    Case 10
'                        If curVanActu > 0 Then
'                            dblTasInte = ((1 + (dblTasPdia / curVanActu)) ^ 365) - 1
'                        End If
'                        dblTasa = ((1 + (dblTasPdia / curPatr)) ^ 365) - 1
'
'                        If strCodFondo <> "00" Then
'                            If strCodMone = "D" Then
'                                curPatr = curPatr / dblTC
'                                curActivo = curActivo / dblTC
'                            End If
'                        Else
'                            strCodMone = "D"
'                            curPatr = curPatr / dblTC
'                            curActivo = curActivo / dblTC
'                        End If
'
'                        .Formulas(6) = "ValPatr=" & Format(curPatr, "0.00")
'                        .Formulas(7) = "Activo=" & Format(curActivo, "0.00")
'                        .Formulas(8) = "PorcCart=" & Format((dblTasInte * 100), "0.00")
'                        .Formulas(9) = "TasPatr=" & Format((dblTasa * 100), "0.00")
'
'                    Case 12
'                        With gadoComando
'                            Set adoConsulta = New ADODB.Recordset
'                            .CommandType = adCmdText
'
'                            dblValcuot = 0: intCntPart = 0
'                            .CommandText = "SELECT VAL_CALC,CNT_PART FROM FMCUOTAS WHERE FCH_CUOT='" & Format(dtpFechaDesde.Text, "yyyymmdd") & "' AND COD_FOND='" & strCodFondo & "'"
'                            Set adoConsulta = .Execute
'                            If Not adoConsulta.EOF Then
'                                dblValcuot = CDbl(adoConsulta("VAL_CALC")): intCntPart = CInt(adoConsulta("CNT_PART"))
'                            End If
'                            adoConsulta.Close
'
'                            datFchTmp = DateAdd("m", -6, dtpFechaDesde.Text): dblValCuot6 = 0: intDiaTmp = 0
'                            intDiaTmp = LUltDiaMes(CVDate(dtpFechaDesde.Text))
'                            If CInt(Left(dtpFechaDesde.Text, 2)) = intDiaTmp Then
'                                intUltDia = LUltDiaMes(datFchTmp)
'                            Else
'                                intUltDia = CInt(Left(dtpFechaDesde.Text, 2))
'                            End If
'
'                            datFchTmp = CVDate(Format(intUltDia, "00") & "/" & Mid(datFchTmp, 4, 2) & "/" & Right(datFchTmp, 4))
'                            .CommandText = "SELECT VAL_CALC FROM FMCUOTAS WHERE FCH_CUOT='" & Format(datFchTmp, "yyyymmdd") & "' AND COD_FOND='" & strCodFondo & "'"
'                            Set adoConsulta = .Execute
'                            If Not adoConsulta.EOF Then
'                                dblValCuot6 = CDbl(adoConsulta("VAL_CALC"))
'                            End If
'                            adoConsulta.Close
'
'                            datFchTmp = DateAdd("m", -12, dtpFechaDesde.Text): dblValCuot12 = 0: intDiaTmp = 0
'                            intDiaTmp = LUltDiaMes(CVDate(dtpFechaDesde.Text))
'                            If CInt(Left(dtpFechaDesde.Text, 2)) = intDiaTmp Then
'                                intUltDia = LUltDiaMes(datFchTmp)
'                            Else
'                                intUltDia = CInt(Left(dtpFechaDesde.Text, 2))
'                            End If
'
'                            datFchTmp = CVDate(Format(intUltDia, "00") & "/" & Mid(datFchTmp, 4, 2) & "/" & Right(datFchTmp, 4))
'                            .CommandText = "SELECT VAL_CALC FROM FMCUOTAS WHERE FCH_CUOT='" & Format(datFchTmp, "yyyymmdd") & "' AND COD_FOND='" & strCodFondo & "'"
'                            Set adoConsulta = .Execute
'                            If Not adoConsulta.EOF Then
'                                dblValCuot12 = CDbl(adoConsulta("VAL_CALC"))
'                            End If
'                            adoConsulta.Close: Set adoConsulta = Nothing
'                        End With
'
'                        .Formulas(6) = "ValPatr=" & curPatr
'                        .Formulas(7) = "Activo=" & curActivo
'                        .Formulas(8) = "ValCuota=" & dblValcuot
'                        .Formulas(9) = "NroPart=" & intCntPart
'                        .Formulas(10) = "ValCuot6=" & dblValCuot6
'                        .Formulas(11) = "ValCuot12=" & dblValCuot12
'
'                    Case 19
'                        Dim curLimMin0 As Currency, curLimMax0 As Currency
'                        Dim curLimMin1 As Currency, curLimMax1 As Currency
'                        Dim curLimMin2 As Currency, curLimMax2 As Currency
'                        Dim curLimMin3 As Currency, curLimMax3 As Currency
'                        Dim curLimMin4 As Currency, curLimMax4 As Currency
'                        Dim curLimMin5 As Currency, curLimMax5 As Currency
'
'                        curPatr = curPatr / dblTC
'
'                        With gadoComando
'                            Set adoConsulta = New ADODB.Recordset
'                            .CommandType = adCmdText
'
'                            .CommandText = "SELECT COD_FILE,MIN_LIMI,MAX_LIMI FROM tblCtrlLimites WHERE COD_LIMI='03' AND COD_FOND='00'"
'                            Set adoConsulta = .Execute
'                            Do While Not adoConsulta.EOF
'                                Select Case Trim(adoConsulta!COD_FILE)
'                                    Case "00"
'                                        curLimMin0 = adoConsulta("MIN_LIMI")
'                                        curLimMax0 = adoConsulta("MAX_LIMI")
'                                    Case "01"
'                                        curLimMin1 = adoConsulta("MIN_LIMI")
'                                        curLimMax1 = adoConsulta("MAX_LIMI")
'                                    Case "02"
'                                        curLimMin2 = adoConsulta("MIN_LIMI")
'                                        curLimMax2 = adoConsulta("MAX_LIMI")
'                                    Case "03"
'                                        curLimMin3 = adoConsulta("MIN_LIMI")
'                                        curLimMax3 = adoConsulta("MAX_LIMI")
'                                    Case "04"
'                                        curLimMin4 = adoConsulta("MIN_LIMI")
'                                        curLimMax4 = adoConsulta("MAX_LIMI")
'                                    Case "05"
'                                        curLimMin5 = adoConsulta("MIN_LIMI")
'                                        curLimMax5 = adoConsulta("MAX_LIMI")
'                                End Select
'                                adoConsulta.MoveNext
'                            Loop
'                            adoConsulta.Close: Set adoConsulta = Nothing
'                        End With
'
'                        .Formulas(6) = "ValPatr=" & curPatr
'                        .Formulas(7) = "LimMin0=" & curLimMin0
'                        .Formulas(8) = "LimMax0=" & curLimMax0
'                        .Formulas(9) = "LimMin1=" & curLimMin1
'                        .Formulas(10) = "LimMax1=" & curLimMax1
'                        .Formulas(11) = "LimMin2=" & curLimMin2
'                        .Formulas(12) = "LimMax2=" & curLimMax2
'                        .Formulas(13) = "LimMin3=" & curLimMin3
'                        .Formulas(14) = "LimMax3=" & curLimMax3
'                        .Formulas(15) = "LimMin4=" & curLimMin4
'                        .Formulas(16) = "LimMax4=" & curLimMax4
'                        .Formulas(17) = "LimMin5=" & curLimMin5
'                        .Formulas(18) = "LimMax5=" & curLimMax5
'
'                End Select
'
'                If strCodFondo <> "00" Then
'                    .SelectionFormula = "{tblCarteraInv.FCH_CART}='" & strFchInicio & "' AND {tblCarteraInv.COD_FOND}='" & strCodFondo & "'"
'                Else
'                    Select Case strTipFon
'                        Case "T"
'                            .SelectionFormula = "{tblCarteraInv.FCH_CART}='" & strFchInicio & "'"
'                        Case "F"
'                            .SelectionFormula = "{tblCarteraInv.FCH_CART}='" & strFchInicio & "' AND ({tblCarteraInv.COD_FOND}='02' OR {tblCarteraInv.COD_FOND}='03' OR {tblCarteraInv.COD_FOND}='04' OR {tblCarteraInv.COD_FOND}='06' OR {tblCarteraInv.COD_FOND}='07')"
'                        Case "V"
'                            .SelectionFormula = "{tblCarteraInv.FCH_CART}='" & strFchInicio & "' AND ({tblCarteraInv.COD_FOND}='01' OR {tblCarteraInv.COD_FOND}='05')"
'                    End Select
'                End If
'
'                If strIndRango = "X" Then .Formulas(10) = "FchProcFin='" & dtpFechaHasta.Text & "'"
'                If strIndRango = "X" Then .SelectionFormula = "({tblCarteraInv.FCH_CART} IN '" & strFchInicio & "' TO '" & strFchFinal & "') AND {tblCarteraInv.COD_FOND}='" & strCodFondo & "'"
'            Case "C" '*** Conasev ***
'                .Formulas(0) = "CodFond='" & strCodFondo & "'"
'                .Formulas(1) = "Fondo='" & Trim$(cboFondo.Text) & "'"
'                .Formulas(2) = "FchRep='" & dtpFechaDesde.Text & "'"
'                .Formulas(3) = "dd='" & Format$(Day(dtpFechaDesde.Text), "00") & "'"
'                .Formulas(4) = "mm='" & Format$(Month(dtpFechaDesde.Text), "00") & "'"
'                .Formulas(5) = "yy='" & Format$(Year(dtpFechaDesde.Text), "0000") & "'"
'                .Formulas(6) = "TipCamb=" & dblTC
'                .Formulas(7) = "CiaName='Santander Central Hispano SAF'"
'
'                If strIndProcedimiento = "X" Then
'                    .StoredProcParam(0) = strCodFondo
'                    .StoredProcParam(1) = strFchInicio
'                End If
'
'            Case "D" '*** Control Diario ***
'                MsgBox "Opción en Construcción"
''                Select Case intNumFormula
''                    Case 6
''                        Dim adoRecAux As ADODB.Recordset
''                        Dim dblTirCalc As Double, dblTasCier As Double
''                        Dim curValNomi As Currency, curValAct As Currency
''                        Dim strTipVac As String, strFile As String, strAnalitica As String
''
''                        With gadoComando
''
''                            dblTirCalc = 0: strTipVac = ""
''                            Set adoRecAux = New ADODB.Recordset
''
''                            .CommandText = "SELECT COD_FILE,COD_ANAL,VAN_ACTU,VAL_NOMI,TAS_CIER FROM tblCarteraInv WHERE FCH_CART='" & Format(dtpFechaDesde.Text, "yyyymmdd") & "' AND COD_FOND='" & strCodFondo & "'"
''                            Set adoConsulta = .Execute
''
''                            Do While adoConsulta.EOF
''                                strFile = Trim(adoConsulta("COD_FILE")): strAnalitica = Trim(adoConsulta("COD_ANAL"))
''                                curValNomi = CCur(adoConsulta("VAL_NOMI")): curValAct = CCur(adoConsulta("VAN_ACTU"))
''                                dblTasCier = CDbl(adoConsulta("TAS_CIER"))
''
''                                .CommandText = "SELECT TIP_VAC FROM FMBONOS WHERE COD_FILE='" & strFile & "' AND COD_ANAL='" & strAnalitica & "'"
''                                Set adoRecAux = .Execute
''                                If Not adoRecAux.EOF Then
''                                    strTipVac = IIf(IsNull(adoRecAux("TIP_VAC")), 0, Trim(adoRecAux("TIP_VAC")))
''                                End If
''                                adoRecAux.Close: Set adoRecAux = Nothing
''
''                                '*** Obtener Tir Base 365 ***
''                                dblTirCalc = TirNoPer(strFile, strAnalitica, DateAdd("d", 1, dtpFechaDesde.Text), DateAdd("d", 1, dtpFechaDesde.Text), curValAct, 0, curValNomi, curValNomi, dblTasCier, strTipVac)
''
''                                .CommandText = "UPDATE tblCarteraInv SET TAS_CALC=" & dblTirCalc & " WHERE COD_ANAL='" & strAnalitica
''
''
''                                adoConsulta.MoveNext
''                            Loop
''
''                        End With
''
''                End Select
'
'        End Select
'
'        If Index = 1 Then .Destination = crptToPrinter
'
'        If Index = 0 Then
'            .Destination = crptToWindow
'            .WindowState = crptNormal
'            .WindowTop = 0
'            .WindowLeft = 0
'            .WindowState = crptMaximized
'        End If
'
'        .Action = 1
'    End With
'
'    Me.MousePointer = vbDefault
'    Exit Sub
'
'Ctrl_Error:
'    GSwErr = True
'    With Err
'        Select Case .Number
'            Case 20504
'                MsgBox "! Reporte NO EXISTE !" & Chr$(10) & Chr$(10) & _
'                        "CONSULTE con el Administrador del Sistema.", vbCritical, Me.Caption
'            Case Else
'                MsgBox " Error " & Str$(.Number) & " -  " & .Description, vbCritical, Me.Caption
'        End Select
'    End With
'    Me.MousePointer = vbDefault
'    Resume Next
    
End Sub



Private Sub tdgReporte_SelChange(Cancel As Integer)
    
    Dim intRegistro As Integer, intContador         As Integer
    Dim strSQL      As String
    
    intContador = tdgReporte.SelBookmarks.Count - 1
    
    For intRegistro = 0 To intContador
    '-----se modifico porque se agrego un reporte mas -----------
        'tdgReporte.Row = tdgReporte.SelBookmarks(intRegistro) - 1
        'tdgReporte.Row = tdgReporte.SelBookmarks(intRegistro) - 1

        '*** Fechas ***
        dtpFechaDesde.Enabled = False
        If Trim(tdgReporte.Columns(2).Value) = Valor_Indicador Then dtpFechaDesde.Enabled = True
        
        '*** Moneda del reporte ***
        'cboMoneda.Enabled = False
        If Trim(tdgReporte.Columns(9).Value) = Valor_Indicador Then cboMoneda.Enabled = True
              
        '*** Fondos ***
        strSQL = "{ call up_ACSelDatos(8) }"
        If Trim(tdgReporte.Columns(6).Value) = Valor_Indicador Then
            CargarControlLista strSQL, cboFondo, arrFondo(), Sel_Todos
            
            If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
        Else
            If strCodFondo = Valor_Caracter Then
                CargarControlLista strSQL, cboFondo, arrFondo(), Valor_Caracter
                
                If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
            End If
        End If
        
        '*** Tipo de Cartera de Fondos ***
        cboTipoCartera.Enabled = False
        If Trim(tdgReporte.Columns(7).Value) = Valor_Indicador Then cboTipoCartera.Enabled = True
        
        strCodReporte = Trim(tdgReporte.Columns(0).Value)
        If chkSimulacion.Enabled Then
            If chkSimulacion.Value Then strCodReporte = Trim(tdgReporte.Columns(8).Value)
        End If
        
        If strCodReporte = "BalanceComprobacion" Or strCodReporte = "BalanceComprobacionSim" Then
            chkBalanceEleccion.Visible = True
            lblBalanceEleccion.Visible = True
        Else
            If strCodReporte <> "BalanceComprobacion" Then
                chkBalanceEleccion.Value = False
                chkBalanceEleccion.Visible = False
                lblBalanceEleccion.Visible = False
            End If
        End If
               
    Next
    
End Sub

