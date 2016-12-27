VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmMovimientoAnalitica 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimientos por Cuenta Analítica"
   ClientHeight    =   6810
   ClientLeft      =   2070
   ClientTop       =   1215
   ClientWidth     =   13170
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6810
   ScaleWidth      =   13170
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   10680
      TabIndex        =   1
      Top             =   6000
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
      Left            =   720
      TabIndex        =   0
      Top             =   6000
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
      Caption0        =   "&Buscar"
      Tag0            =   "5"
      Visible0        =   0   'False
      ToolTipText0    =   "Buscar"
      UserControlWidth=   1200
   End
   Begin VB.Frame fraMovimientos 
      Caption         =   "Análisis de Movimientos"
      Height          =   5775
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   12855
      Begin VB.ComboBox cboMoneda 
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
         Left            =   7530
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   810
         Width           =   2325
      End
      Begin VB.CommandButton cmdCuenta 
         Caption         =   "..."
         Height          =   315
         Left            =   8910
         TabIndex        =   11
         ToolTipText     =   "Buscar Cuenta Contable"
         Top             =   360
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.TextBox txtCodCuenta 
         Height          =   315
         Left            =   7530
         TabIndex        =   10
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtAnalitica 
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
         Left            =   11460
         MaxLength       =   8
         TabIndex        =   9
         Top             =   360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtFile 
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
         Left            =   10770
         MaxLength       =   3
         TabIndex        =   8
         Top             =   360
         Visible         =   0   'False
         Width           =   615
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
         Left            =   1350
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   360
         Width           =   4125
      End
      Begin VB.CheckBox chkCuenta 
         Caption         =   "Cuenta Contable"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   5700
         TabIndex        =   6
         ToolTipText     =   "Marcar para filtrar por cuenta"
         Top             =   420
         Width           =   1815
      End
      Begin VB.CheckBox chkAnalitica 
         Caption         =   "Analítica"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   9600
         TabIndex        =   5
         ToolTipText     =   "Marcar para filtrar por analítica"
         Top             =   420
         Width           =   1455
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmMovimientoAnalitica.frx":0000
         Height          =   3975
         Left            =   240
         OleObjectBlob   =   "frmMovimientoAnalitica.frx":001A
         TabIndex        =   4
         Top             =   1260
         Width           =   12375
      End
      Begin MSComCtl2.DTPicker dtpFechaDesde 
         Height          =   315
         Left            =   1350
         TabIndex        =   12
         Top             =   810
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   177340417
         CurrentDate     =   38068
      End
      Begin MSComCtl2.DTPicker dtpFechaHasta 
         Height          =   315
         Left            =   4140
         TabIndex        =   13
         Top             =   810
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   177340417
         CurrentDate     =   38068
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Moneda Contable"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   6
         Left            =   5700
         TabIndex        =   22
         Top             =   870
         Width           =   1500
      End
      Begin VB.Label lblDescrip 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Contable"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   5
         Left            =   6360
         TabIndex        =   21
         Top             =   5400
         Width           =   1335
      End
      Begin VB.Label lblDescrip 
         BackStyle       =   0  'Transparent
         Caption         =   "Total ME"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   4
         Left            =   3240
         TabIndex        =   20
         Top             =   5400
         Width           =   1215
      End
      Begin VB.Label lblDescrip 
         BackStyle       =   0  'Transparent
         Caption         =   "Total MN"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   19
         Top             =   5400
         Width           =   1095
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Desde"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   18
         Top             =   870
         Width           =   615
      End
      Begin VB.Label lblMontoTotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         Height          =   285
         Index           =   0
         Left            =   1560
         TabIndex        =   17
         Top             =   5355
         Width           =   1365
      End
      Begin VB.Label lblMontoTotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         Height          =   285
         Index           =   1
         Left            =   4560
         TabIndex        =   16
         Top             =   5355
         Width           =   1365
      End
      Begin VB.Label lblMontoTotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         Height          =   285
         Index           =   2
         Left            =   7875
         TabIndex        =   15
         Top             =   5355
         Width           =   1365
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Fondo"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   14
         Top             =   420
         Width           =   615
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Hasta"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   3240
         TabIndex        =   3
         Top             =   870
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmMovimientoAnalitica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrFondo()      As String
Dim arrMoneda()     As String

Dim strCodFondo     As String, strCodMoneda     As String
Dim strEstado       As String, strSQL           As String
Dim adoConsulta     As ADODB.Recordset
Dim indSortAsc      As Boolean, indSortDesc     As Boolean

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

    Dim strFechaDesde   As String, strFechaHasta As String
    Dim datFecha        As Date
    Dim strSQL          As String
    Dim intRegistro     As Integer
    
    Set adoConsulta = New ADODB.Recordset
        
    strFechaDesde = Convertyyyymmdd(dtpFechaDesde.Value)
    datFecha = DateAdd("d", 1, dtpFechaHasta.Value)
    strFechaHasta = Convertyyyymmdd(datFecha)
        
    If chkCuenta.Value And chkAnalitica.Value Then
        strSQL = "SELECT FechaMovimiento,CodCuenta,CodFile,CodAnalitica,DescripMovimiento,IndDebeHaber," & _
            "MontoMovimiento,MontoMovimientoContable, M.CodSigno, ACDM.CodMoneda " & _
            "FROM AsientoContableDetalleMovimiento ACDM " & _
            "JOIN Moneda M on (ACDM.CodMoneda = M.CodMoneda)" & _
            "WHERE (FechaMovimiento >='" & strFechaDesde & "' AND FechaMovimiento <'" & strFechaHasta & "') AND " & _
            "CodCuenta='" & Trim(txtCodCuenta.Text) & "' AND CodFile='" & Trim(txtFile.Text) & "' AND CodAnalitica='" & Trim(txtAnalitica.Text) & "' AND " & _
            "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'  AND CodMonedaContable = '" & strCodMoneda & "' " & _
            "ORDER BY CodCuenta,CodFile,CodAnalitica,FechaMovimiento"
    ElseIf Not chkCuenta.Value And chkAnalitica.Value Then
        strSQL = "SELECT FechaMovimiento,CodCuenta,CodFile,CodAnalitica,DescripMovimiento,IndDebeHaber," & _
            "MontoMovimiento,MontoMovimientoContable, M.CodSigno, ACDM.CodMoneda " & _
            "FROM AsientoContableDetalleMovimiento ACDM " & _
            "JOIN Moneda M on (ACDM.CodMoneda = M.CodMoneda)" & _
            "WHERE (FechaMovimiento >='" & strFechaDesde & "' AND FechaMovimiento <'" & strFechaHasta & "') AND " & _
            "CodFile='" & Trim(txtFile.Text) & "' AND CodAnalitica='" & Trim(txtAnalitica.Text) & "' AND " & _
            "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'  AND CodMonedaContable = '" & strCodMoneda & "' " & _
            "ORDER BY CodCuenta,CodFile,CodAnalitica,FechaMovimiento"
    ElseIf Not chkAnalitica.Value And chkCuenta.Value Then
        strSQL = "SELECT FechaMovimiento,CodCuenta,CodFile,CodAnalitica,DescripMovimiento,IndDebeHaber," & _
            "MontoMovimiento,MontoMovimientoContable, M.CodSigno, ACDM.CodMoneda " & _
            "FROM AsientoContableDetalleMovimiento ACDM " & _
            "JOIN Moneda M on (ACDM.CodMoneda = M.CodMoneda)" & _
            "WHERE (FechaMovimiento >='" & strFechaDesde & "' AND FechaMovimiento <'" & strFechaHasta & "') AND " & _
            "CodCuenta='" & Trim(txtCodCuenta.Text) & "' AND CodFondo='" & strCodFondo & "' AND " & _
            "CodAdministradora='" & gstrCodAdministradora & "'  AND CodMonedaContable = '" & strCodMoneda & "' " & _
            "ORDER BY CodCuenta,CodFile,CodAnalitica,FechaMovimiento"
    Else
        strSQL = "SELECT FechaMovimiento,CodCuenta,CodFile,CodAnalitica,DescripMovimiento,IndDebeHaber," & _
            "MontoMovimiento,MontoMovimientoContable, M.CodSigno, ACDM.CodMoneda " & _
            "FROM AsientoContableDetalleMovimiento ACDM " & _
            "JOIN Moneda M on (ACDM.CodMoneda = M.CodMoneda)" & _
            "WHERE (FechaMovimiento >='" & strFechaDesde & "' AND FechaMovimiento <'" & strFechaHasta & "') AND " & _
            "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND CodMonedaContable = '" & strCodMoneda & "' " & _
            "ORDER BY CodCuenta,CodFile,CodAnalitica,FechaMovimiento"
    End If
    
    With adoConsulta
        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open strSQL
    End With

    tdgConsulta.DataSource = adoConsulta
        
    Dim dblMontoMN          As Double, dblMontoME           As Double
    Dim dblAcumuladoMN      As Double, dblAcumuladoME       As Double
    Dim dblMontoAcumuladoContable  As Double
    
    If adoConsulta.RecordCount = 0 Then Exit Sub
    
    With adoConsulta
        .MoveFirst
        
        dblMontoAcumuladoContable = 0
        dblAcumuladoMN = 0
        dblAcumuladoME = 0
        
        Do While Not .EOF
            If .Fields("CodMoneda") = "01" Then
                dblMontoMN = CDbl(.Fields("MontoMovimiento"))
                dblMontoME = 0
            Else
                dblMontoMN = 0
                dblMontoME = CDbl(.Fields("MontoMovimiento"))
            End If
            
            dblAcumuladoMN = dblAcumuladoMN + dblMontoMN
            dblAcumuladoME = dblAcumuladoME + dblMontoME
            dblMontoAcumuladoContable = dblMontoAcumuladoContable + CDbl(.Fields("MontoMovimientoContable"))
            .MoveNext
        Loop
    End With
        
    lblMontoTotal(0).Caption = CStr(dblAcumuladoMN)
    lblMontoTotal(1).Caption = CStr(dblAcumuladoME)
    lblMontoTotal(2).Caption = CStr(dblMontoAcumuladoContable)
End Sub

Public Sub Cancelar()

End Sub

Private Sub CargarReportes()

'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Movimiento por Analítica"
    
End Sub

Public Sub Eliminar()

End Sub


Public Sub Grabar()

End Sub

Public Sub Importar()

End Sub

Public Sub Imprimir()

    Call SubImprimir(1)
            
End Sub


Public Sub Modificar()

End Sub

Public Sub Salir()

    Unload Me
    
End Sub

Public Sub SubImprimir(Index As Integer)

'    Call LDoGrid(True)
'
'    Dim frmRpt As frmReportViewer
'    Dim aReportParamS(), aReportParamF(), aReportParamFn()
'
'    Set frmRpt = New frmReportViewer
'
'    ReDim aReportParamS(0)
'    ReDim aReportParamFn(6)
'    ReDim aReportParamF(6)
'
'    Select Case Index
'        Case 1
'            gstrNameRepo = "crCNMovimientoAnalitica"
'
'            aReportParamFn(0) = "Usuario"
'            aReportParamFn(1) = "CodFond"
'            aReportParamFn(2) = "Fondo"
'            aReportParamFn(3) = "FchDel"
'            aReportParamFn(4) = "FchAl"
'            aReportParamFn(5) = "Hora"
'            aReportParamFn(6) = "CiaName"
'
'            aReportParamF(0) = gstrLogin
'            aReportParamF(1) = s_CodFon
'            aReportParamF(2) = Left(cmb_FonMut & Space(40), 40)
'            aReportParamF(3) = Dat_FchCns(0).Value
'            aReportParamF(4) = Dat_FchCns(1).Value
'            aReportParamF(5) = Format(Time(), "hh:mm")
'            aReportParamF(6) = gstrCiaName
'    End Select
'
'    gstrSelFrml = ""
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

Private Sub cboFondo_Click()

    Dim adoRegistro As ADODB.Recordset
    
    strCodFondo = Valor_Caracter
    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(arrFondo(cboFondo.ListIndex))
    
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Fecha Vigente, Valor Cuota, Valor Cuota T-1, Moneda y Cantidad Inicial de Cuotas del Fondo ***
        .CommandText = "{ call up_ACSelDatosParametro(23,'" & strCodFondo & "','" & gstrCodAdministradora & "') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            gdatFechaActual = adoRegistro("FechaCuota")
            gdblTipoCambio = adoRegistro("ValorTipoCambio")
            gstrCodMoneda = adoRegistro("CodMoneda")
            dtpFechaDesde.Value = gdatFechaActual
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

Private Sub chkAnalitica_Click()

    If chkAnalitica.Value Then
        txtFile.Visible = True
        txtAnalitica.Visible = True
    Else
        txtFile.Visible = False
        txtAnalitica.Visible = False
    End If
    
End Sub

Private Sub chkCuenta_Click()

    If chkCuenta.Value Then
        txtCodCuenta.Visible = True
        cmdCuenta.Visible = True
    Else
        txtCodCuenta.Visible = False
        cmdCuenta.Visible = False
    End If
    
End Sub

Private Sub cmdCuenta_Click()

    gstrFormulario = "frmMovimientoAnalitica"
    frmBusquedaCuenta.Show vbModal
    
End Sub


Private Sub dtpFechaDesde_Change()
    If dtpFechaDesde.Value > dtpFechaHasta.Value Then
        dtpFechaDesde.Value = dtpFechaHasta.Value
    End If
End Sub

Private Sub dtpFechaHasta_Change()
    If dtpFechaDesde.Value > dtpFechaHasta.Value Then
       dtpFechaHasta.Value = dtpFechaDesde.Value
    End If
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
    
    '*** Fondos ***
    strSQL = "{ call up_ACSelDatosParametro(74,'" & gstrCodAdministradora & "','" & gstrCodFondoContable & "') }"
    CargarControlLista strSQL, cboFondo, arrFondo(), Valor_Caracter
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
    
    strSQL = "{ call up_ACSelDatos(2) }"
    CargarControlLista strSQL, cboMoneda, arrMoneda(), Valor_Caracter
    
    intRegistro = ObtenerItemLista(arrMoneda(), gstrCodMoneda)
    If intRegistro >= 0 Then cboMoneda.ListIndex = intRegistro
        
End Sub
Private Sub InicializarValores()

    strEstado = Reg_Defecto
    chkCuenta.Value = vbUnchecked
    chkAnalitica.Value = vbUnchecked
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
            
End Sub
Private Sub LDoGrid(blngraba As Boolean)
    
'    aGrdCnf(5).TitDes = "Tipo"
'    aGrdCnf(5).DatNom = "TIP_GENR"
'
'    aGrdCnf(7).TitDes = "Val.(S/.)"
'    aGrdCnf(7).DatNom = "VAL_MOVN"
'
'    aGrdCnf(8).TitDes = "Val.(US$)"
'    aGrdCnf(8).DatNom = "VAL_MOVX"
'
'    aGrdCnf(9).TitDes = "Contable"
'    aGrdCnf(9).DatNom = "VAL_CONT"
'
'    aGrdCnf(10).TitDes = "Saldo US$"
'    aGrdCnf(10).DatNom = "SIN_MONX"
'
'    aGrdCnf(11).TitDes = "Saldo Cont."
'    aGrdCnf(11).DatNom = "SIN_CONT"
    
'    strSQL = "SELECT FMMOVCON.COD_FOND, FMMOVCON.FCH_MOVI, FMMOVCON.COD_CTA, FMMOVCON.COD_FILE,FMMOVCON.COD_ANAL,FMMOVCON.TIP_GENR,"
'    strSQL = strSQL & "FMMOVCON.DSC_MOVI,FMMOVCON.VAL_MOVN,FMMOVCON.VAL_MOVX,FMMOVCON.VAL_CONT, FMSALDOS.SIN_MONX, FMSALDOS.SIN_CONT "
'    strSQL = strSQL & "FROM FMMOVCON, FMSALDOS"
'    strSQL = strSQL & " WHERE FMMOVCON.COD_FOND='" & s_CodFon & "'"
'    If chk_Cns(0).Value Then
'        strSQL = strSQL & " AND FMMOVCON.COD_CTA='" & txt_NroCta.Text & "' "
'    End If
'    If chk_Cns(1).Value Then
'        strSQL = strSQL & " AND FMMOVCON.COD_FILE='" & txt_CodFil.Text & "' "
'        strSQL = strSQL & " AND FMMOVCON.COD_ANAL='" & txt_CodAna.Text & "' "
'    End If
'    strSQL = strSQL & " AND (FCH_MOVI BETWEEN '" & FmtFec(Dat_FchCns(0).Value, "win", "yyyymmdd", Res) & "' AND '" & FmtFec(Dat_FchCns(1).Value, "win", "yyyymmdd", Res) & "')"
'    strSQL = strSQL & " AND FMMOVCON.COD_FOND = FMSALDOS.COD_FOND "
'    strSQL = strSQL & " AND FMMOVCON.COD_FILE = FMSALDOS.COD_FILE "
'    strSQL = strSQL & " AND FMMOVCON.COD_ANAL = FMSALDOS.COD_ANAL "
'    strSQL = strSQL & " AND FMMOVCON.COD_CTA  = FMSALDOS.COD_CTA "
'    strSQL = strSQL & " AND FMMOVCON.FCH_MOVI = FMSALDOS.FCH_SALD "
'    adoComm.CommandText = strSQL
'    Set adoresultaux2 = adoComm.Execute
'    If Not blngraba Then
'        Call LlenarGrid(Grd_Mov, adoresultaux2, aGrdCnf(), Adirreg())
'    Else
'        strSQL = "delete tmovanal"
'        adoComm.CommandText = strSQL
'        adoConn.Execute adoComm.CommandText
'        Do While Not adoresultaux2.EOF
'            If dblsalimovn <> adoresultaux2!sin_cont Then
'                dblsalimovn = adoresultaux2!sin_cont
'                dblacummovn = 0
'            End If
'
'            If dblsalimovx <> adoresultaux2!sin_monx Then
'                dblsalimovx = adoresultaux2!sin_monx
'                dblacummovx = 0
'            End If
'
'            dblacummovn = dblacummovn + adoresultaux2!VAL_CONT
'            dblsaldmovn = dblacummovn + dblsalimovn
'            dblacummovx = dblacummovx + adoresultaux2!VAL_MOVX
'            dblsaldmovx = dblacummovx + dblsalimovx
'
'            strSQL = "Insert Into tmovanal values ('"
'            strSQL = strSQL & adoresultaux2!FCH_MOVI & "', '"
'            strSQL = strSQL & adoresultaux2!COD_CTA & "', '"
'            strSQL = strSQL & adoresultaux2!COD_FILE & "', '"
'            strSQL = strSQL & adoresultaux2!COD_ANAL & "', '"
'            strSQL = strSQL & adoresultaux2!DSC_MOVI & "', "
'            strSQL = strSQL & adoresultaux2!VAL_MOVN & ", "
'            strSQL = strSQL & adoresultaux2!VAL_MOVX & ", "
'            strSQL = strSQL & adoresultaux2!VAL_CONT & ", "
'            strSQL = strSQL & dblsaldmovx & ", "
'            strSQL = strSQL & dblsaldmovn & " )"
'            adoComm.CommandText = strSQL
'            adoConn.Execute adoComm.CommandText
'
'            adoresultaux2.MoveNext
'        Loop
'    End If
'    adoresultaux2.Close: Set adoresultaux2 = Nothing
   
End Sub

Private Sub LGrdCalc()

'    Dim n_Row As Long, n_TotRow As Long, dblsalimovx As Double, strSQL As String
'    Dim n_TotMonn As Currency, n_TotMonx As Currency, n_TotMonc As Currency
'    Dim dblacummovn As Double, dblacummovx As Double, dblsalimovn As Double
'
'    n_TotMonn = 0: n_TotMonx = 0: n_TotMonc = 0
'    n_TotRow = Grd_Mov.Rows - 1
'    dblsalimovn = 0: dblsalimovx = 0
'    dblacummovn = 0: dblacummovx = 0
'
'    Grd_Mov.Row = 1
'    Grd_Mov.Col = 11
'    If Trim(Grd_Mov.Text) = "" Then
'        Exit Sub
'    End If
'
'    For n_Row = 1 To n_TotRow
'        Grd_Mov.Row = n_Row
'        Grd_Mov.Col = 11
'        If dblsalimovn <> CDbl(Grd_Mov.Text) Then
'            dblsalimovn = CDbl(Grd_Mov.Text)
'            dblacummovn = 0
'        End If
'        Grd_Mov.Col = 10
'        If dblsalimovx <> CDbl(Grd_Mov.Text) Then
'            dblsalimovx = CDbl(Grd_Mov.Text)
'            dblacummovx = 0
'        End If
'        Grd_Mov.Col = 9
'        dblacummovn = dblacummovn + CDbl(Grd_Mov.Text)
'        Grd_Mov.Col = 11
'        Grd_Mov.Text = Format(dblacummovn + dblsalimovn, "#,##0.00")
'        Grd_Mov.Col = 8
'        dblacummovx = dblacummovx + CDbl(Grd_Mov.Text)
'        Grd_Mov.Col = 10
'        Grd_Mov.Text = Format(dblacummovx + dblsalimovx, "#,##0.00")
'        Grd_Mov.Col = 7
'        If IsNumeric(Grd_Mov.Text) Then
'            n_TotMonn = n_TotMonn + CDbl(Grd_Mov.Text)
'        End If
'        Grd_Mov.Col = 8
'        If IsNumeric(Grd_Mov.Text) Then
'            n_TotMonx = n_TotMonx + CDbl(Grd_Mov.Text)
'        End If
'        Grd_Mov.Col = 9
'        If IsNumeric(Grd_Mov.Text) Then
'            n_TotMonc = n_TotMonc + CDbl(Grd_Mov.Text)
'        End If
'        If blngraba Then
'            strSQL = "insert into "
'        End If
'
'    Next
'    lbl_MtoTot(0).Caption = Format(CCur(n_TotMonn), "#,##0.00")
'    lbl_MtoTot(1).Caption = Format(CCur(n_TotMonx), "#,##0.00")
'    lbl_MtoTot(2).Caption = Format(CCur(n_TotMonc), "#,##0.00")
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call OcultarReportes
    Set frmMovimientoAnalitica = Nothing
    
End Sub

Private Sub lblMontoTotal_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblMontoTotal(Index), Decimales_Monto)
    
End Sub

Private Sub tdgConsulta_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)

    If ColIndex = 5 Then
        Call DarFormatoValor(Value, Decimales_Monto)
    End If
    
    If ColIndex = 6 Then
        Call DarFormatoValor(Value, Decimales_Monto)
    End If
    
    If ColIndex = 7 Then
        Call DarFormatoValor(Value, Decimales_Monto)
    End If
    
End Sub

Private Sub txtAnalitica_LostFocus()

    txtAnalitica.Text = Right(String(8, "0") & Trim(txtAnalitica.Text), 8)
    
End Sub

Private Sub txtCodCuenta_LostFocus()

    If Trim(txtCodCuenta.Text) = Valor_Caracter Then Exit Sub
    
    If Not ValidarCuentaContable(txtCodCuenta.Text, gstrCodAdministradora) Then
        MsgBox "Cuenta no existe...", vbCritical
        cmdCuenta.SetFocus
    End If
    
End Sub

Private Sub txtFile_LostFocus()

    txtFile.Text = Right(String(3, "0") & Trim(txtFile.Text), 3)
    
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
