VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{442CAE95-1D41-47B1-BE83-6995DA3CE254}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmPeriodoContable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Períodos Contables"
   ClientHeight    =   6345
   ClientLeft      =   1245
   ClientTop       =   1395
   ClientWidth     =   6855
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6345
   ScaleWidth      =   6855
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   4920
      TabIndex        =   5
      Top             =   5520
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
      TabIndex        =   4
      Top             =   5520
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
      Caption0        =   "&Nuevo"
      Tag0            =   "0"
      ToolTipText0    =   "Nuevo"
      UserControlWidth=   1200
   End
   Begin VB.ComboBox cboPeriodo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   720
      Width           =   4440
   End
   Begin VB.Frame fraPeriodo 
      Caption         =   "Períodos"
      Height          =   4215
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   6375
      Begin MSDataGridLib.DataGrid dgdConsulta 
         Height          =   3600
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   6350
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "PeriodoContable"
            Caption         =   "PeriodoContable"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "MesContable"
            Caption         =   "MesContable"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "DescripPeriodo"
            Caption         =   "Periodo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "FechaInicio"
            Caption         =   "Fecha Inicial"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "FechaFinal"
            Caption         =   "Fecha Final"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "FechaApertura"
            Caption         =   "Fecha Apertura"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "FechaCierre"
            Caption         =   "Fecha Cierre"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            Locked          =   -1  'True
            BeginProperty Column00 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column01 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2940.095
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1365.165
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1244.976
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin VB.ComboBox cboFondo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   4440
   End
   Begin VB.Label lblDescrip 
      Caption         =   "Periodo"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   6
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblDescrip 
      Caption         =   "Fondo"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmPeriodoContable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrFondo()          As String, arrPeriodo()         As String

Dim strCodFondo         As String, strCodPeriodo        As String
Dim strCodMoneda        As String
Dim strEstado           As String

Public Sub Adicionar()

    If MsgBox(Mensaje_Adicion_Periodo, vbYesNo + vbQuestion + vbDefaultButton2, gstrNombreEmpresa) = vbYes Then
        frmMainMdi.stbMdi.Panels(3).Text = "Adicionar periodo contable..."
        
        strEstado = Reg_Adicion
        Me.MousePointer = vbHourglass
        Call GenerarPeriodo
        Me.MousePointer = vbDefault
        
        cboFondo_Click
        Me.Refresh
        
        MsgBox Mensaje_Adicion_Periodo_Exitoso, vbExclamation, gstrNombreEmpresa
        
    End If
    
End Sub

Public Sub Eliminar()

End Sub

Private Sub GenerarPeriodo()

    Dim datFechaInicioPeriodo   As Date, datFechaFinPeriodo     As Date
    Dim intPeriodo              As Integer
    
    intPeriodo = CInt(gstrPeriodoActual)
    intPeriodo = intPeriodo + 1
    
    datFechaInicioPeriodo = Convertddmmyyyy(Format(intPeriodo, "0000") & "0101")
    datFechaFinPeriodo = Convertddmmyyyy(Format(intPeriodo, "0000") & "1231")
    '*** Generar Periodo Contable del Fondo ***
    frmMainMdi.stbMdi.Panels(3).Text = "Generando Periodo Contable..."
    Call GenerarPeriodoContable(gstrTipoAdministradora, gstrCodAdministradora, strCodFondo, gstrCodMoneda, datFechaInicioPeriodo, datFechaFinPeriodo)
    
End Sub

Public Sub Grabar()

                
End Sub

Public Sub Imprimir()

End Sub

Public Sub Modificar()

End Sub

Public Sub Salir()

    Unload Me
    
End Sub

Public Sub SubImprimir(Index As Integer)

    Dim strSeleccionRegistro    As String
    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()

    Select Case Index
        Case 1
            gstrNameRepo = "PeriodoContable"
            Set frmReporte = New frmVisorReporte
            
            ReDim aReportParamS(2)
            ReDim aReportParamFn(2)
            ReDim aReportParamF(2)
            
            aReportParamFn(0) = "Usuario"
            aReportParamFn(1) = "Hora"
            aReportParamFn(2) = "NombreEmpresa"
                
            aReportParamF(0) = gstrLogin
            aReportParamF(1) = Format(Time(), "hh:mm:ss")
            aReportParamF(2) = gstrNombreEmpresa & Space(1)
                            
            aReportParamS(0) = strCodPeriodo
            aReportParamS(1) = strCodFondo
            aReportParamS(2) = gstrCodAdministradora
            
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

    Dim adoRegistro As ADODB.Recordset
    Dim strSQL      As String, intRegistro As Integer
    
    strCodFondo = Valor_Caracter
    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(arrFondo(cboFondo.ListIndex))
    
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Fecha Vigente, Valor Cuota, Valor Cuota T-1, Moneda y Cantidad Inicial de Cuotas del Fondo ***
        .CommandText = "{ call up_ACSelDatosParametro(50,'" & strCodFondo & "','" & gstrCodAdministradora & "') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            gdatFechaActual = adoRegistro("FechaFinal")
            gdblTipoCambio = 1 'adoRegistro("ValorTipoCambio")
            gstrCodMoneda = adoRegistro("CodMoneda")
            gstrPeriodoActual = Year(CVDate(adoRegistro("FechaFinal")))
            
            frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    '*** Periodos ***
    strSQL = "{ call up_CNSelPeriodoContableActual ('" & gstrPeriodoActual & "') }"
    CargarControlLista strSQL, cboPeriodo, arrPeriodo(), ""
    
    intRegistro = ObtenerItemLista(arrPeriodo(), gstrPeriodoActual)
    If intRegistro >= 0 Then cboPeriodo.ListIndex = intRegistro
    
End Sub

Private Sub cboPeriodo_Click()

    strCodPeriodo = Valor_Caracter
    If cboPeriodo.ListIndex < 0 Then Exit Sub
    
    strCodPeriodo = Trim(arrPeriodo(cboPeriodo.ListIndex))
    
    Call Buscar
    
End Sub

Private Sub cmdPrint_Click()

'    Dim frmRpt As frmReportViewer
'    Dim aReportParamS(), aReportParamF(), aReportParamFn()
'
'    Set frmRpt = New frmReportViewer
'
'    ReDim aReportParamS(0)
'    ReDim aReportParamFn(6)
'    ReDim aReportParamF(6)
'    gstrNameRepo = "FLPERANU"
''    With gobjReport
''        .Formulas(0) = "usuario= '" & gstrLogin & "'"
''        .Formulas(1) = "HORA='" & Format(Time(), "hh:mm") & "'"
''        .Formulas(2) = "Fondo='" & Left$(fondo.Text + Space(40), 40) & "'"
''        .Formulas(3) = "PrdDel='" & Prd1.Text & "-" & Mes1.Text & "'"
''        .Formulas(4) = "PrdAl='" & Prd2.Text & "-" & Mes2.Text & "'"
''        .Formulas(5) = "CodFond='" & cCodfon$ & "'"
''        .Formulas(6) = "cianame='" & gstrNombreEmpresa & "'"
''    End With
'    aReportParamFn(0) = "Usuario"
'    aReportParamFn(1) = "Hora"
'    aReportParamFn(2) = "Fondo"
'    aReportParamFn(3) = "PrdDel"
'    aReportParamFn(4) = "PrdAl"
'    aReportParamFn(5) = "CodFond"
'    aReportParamFn(6) = "CiaName"
'
'    aReportParamF(0) = gstrLogin
'    aReportParamF(1) = Format(Time(), "hh:mm:ss")
'    aReportParamF(2) = Left$(fondo.Text + Space(40), 40)
'    aReportParamF(3) = Prd1.Text & "-" & Mes1.Text
'    aReportParamF(4) = Prd2.Text & "-" & Mes2.Text
'    aReportParamF(5) = cCodfon$
'    aReportParamF(6) = gstrNombreEmpresa
'
'    gstrSelFrml = "{fmprdcon.PRD_CONT} IN '" & Prd1 & "' TO '" & Prd2 & "' AND {fmprdcon.MES_CONT} IN '" & Mes1 & "' TO '" & Mes2 & "'"
'    frmRpt.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"
'
'    Call frmRpt.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())
'
'    frmRpt.Caption = "Reporte - (" & gstrNameRepo & ")"
'    frmRpt.Show vbModal
'
'    Set frmRpt = Nothing
'
'    Screen.MousePointer = vbNormal
    
End Sub

Private Sub Form_Activate()

    Call CargarReportes
    
End Sub

Private Sub Form_Deactivate()

    Call OcultarReportes
    
End Sub

Private Sub Form_Load()

    Call InicializarValores
    Call CargarListas
    Call CargarReportes
    'Call Buscar
    Call DarFormato
    
    Call ValidarPermisoUsoControl(Trim(gstrLoginUS), Me, Trim(App.Title) + Separador_Codigo_Objeto + _
    gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)
    
    CentrarForm Me
    
End Sub

Private Sub DarFormato()
    
    Dim intCont As Integer
    
    For intCont = 0 To (lblDescrip.Count - 1)
        Call FormatoEtiqueta(lblDescrip(intCont), vbLeftJustify)
    Next
            
End Sub
Public Sub Buscar()

    Dim strSQL As String
    
    Dim adoRegistro As ADODB.Recordset
    
    Set adoRegistro = New ADODB.Recordset
  
    strEstado = Reg_Defecto
    
    strSQL = "{ call up_CNSelPeriodoContable ('" & _
                        strCodPeriodo & "','" & _
                        strCodFondo & "','" & _
                        gstrCodAdministradora & "') }"
    
    With adoRegistro
        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open strSQL
    End With
    
    Set dgdConsulta.DataSource = adoRegistro
    
    dgdConsulta.Refresh
    
    If adoRegistro.RecordCount > 0 Then strEstado = Reg_Consulta
    
End Sub

Public Sub Accion(nAccion As ButtonAction)
    
    Select Case nAccion
        
        Case vNew
            Call Adicionar
        Case vSearch
            Call Buscar
        Case vReport
            Call Imprimir
        Case vSave
            Call Grabar
        Case vExit
            Call Salir
        
    End Select
    
End Sub

Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Periodo Contable"
    
End Sub

Private Sub CargarListas()

    Dim strSQL  As String
    
    '*** Fondos ***
    strSQL = "{ call up_ACSelDatosParametro(74,'" & gstrCodAdministradora & "','" & gstrCodFondoContable & "') }"
    CargarControlLista strSQL, cboFondo, arrFondo(), ""
    
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
        
End Sub
Private Sub InicializarValores()

    strEstado = Reg_Defecto
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call OcultarReportes
    Set frmPeriodoContable = Nothing
    frmMainMdi.stbMdi.Panels(3).Text = "Acción"
    
End Sub
