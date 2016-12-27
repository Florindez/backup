VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{830D5A06-9B70-4F8C-98B6-7A19C4E7760B}#1.0#0"; "TAMControls.ocx"
Object = "{442CAE95-1D41-47B1-BE83-6995DA3CE254}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmFondoConceptoActivoFijo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Autorizaci�n de Activos Fijos"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   9600
   Begin VB.CommandButton cmdSeleccionar 
      Caption         =   "&Seleccionar"
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
      Left            =   360
      Picture         =   "frmFondoConceptoActivoFijo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4680
      Width           =   1200
   End
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   6480
      TabIndex        =   0
      Top             =   4680
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   1296
      Buttons         =   2
      Caption0        =   "&Imprimir"
      Tag0            =   "6"
      ToolTipText0    =   "Imprimir"
      Caption1        =   "&Salir"
      Tag1            =   "9"
      ToolTipText1    =   "Salir"
      UserControlWidth=   2700
   End
   Begin TAMControls2.ucBotonEdicion2 cmdOpcion 
      Height          =   735
      Left            =   1800
      TabIndex        =   2
      Top             =   4680
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
      Caption0        =   "&Buscar"
      Tag0            =   "5"
      ToolTipText0    =   "Buscar"
      UserControlWidth=   1200
   End
   Begin TabDlg.SSTab tabGastos 
      Height          =   4215
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   7435
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
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
      TabPicture(0)   =   "frmFondoConceptoActivoFijo.frx":05AB
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "tdgConsulta"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraGasto(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "frmFondoConceptoActivoFijo.frx":05C7
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "cmdAccion(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraGasto(2)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame fraGasto 
         Caption         =   "Selecci�n de Gastos"
         Height          =   2655
         Index           =   2
         Left            =   240
         TabIndex        =   14
         Top             =   600
         Width           =   8895
         Begin TrueOleDBGrid60.TDBGrid tdgGasto 
            Height          =   1575
            Left            =   600
            OleObjectBlob   =   "frmFondoConceptoActivoFijo.frx":05E3
            TabIndex        =   18
            Top             =   720
            Width           =   7695
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fondo"
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   16
            Top             =   360
            Width           =   450
         End
         Begin VB.Label lblDescripFondo 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   1
            Left            =   1320
            TabIndex        =   15
            Top             =   360
            Width           =   7335
         End
      End
      Begin VB.Frame fraGasto 
         Caption         =   "Criterios de B�squeda"
         Height          =   975
         Index           =   0
         Left            =   -74760
         TabIndex        =   11
         Top             =   480
         Width           =   8895
         Begin VB.ComboBox cboFondo 
            Height          =   315
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   360
            Width           =   5535
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "SAB"
            Height          =   195
            Index           =   0
            Left            =   960
            TabIndex        =   13
            Top             =   375
            Width           =   315
         End
      End
      Begin VB.Frame fraGasto 
         Caption         =   "Selecci�n de Gastos"
         Height          =   2655
         Index           =   1
         Left            =   -74760
         TabIndex        =   6
         Top             =   480
         Width           =   8895
         Begin TrueOleDBGrid60.TDBGrid tdgGastos 
            Bindings        =   "frmFondoConceptoActivoFijo.frx":2E38
            Height          =   1575
            Index           =   0
            Left            =   240
            OleObjectBlob   =   "frmFondoConceptoActivoFijo.frx":2E50
            TabIndex        =   7
            Top             =   840
            Width           =   8415
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "SAB"
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   9
            Top             =   360
            Width           =   315
         End
         Begin VB.Label lblDescripFondo 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   0
            Left            =   1320
            TabIndex        =   8
            Top             =   360
            Width           =   7335
         End
      End
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Index           =   0
         Left            =   -69120
         TabIndex        =   4
         Top             =   3240
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
      Begin TAMControls.ucBotonEdicion cmdAccion2 
         Height          =   390
         Left            =   -69120
         TabIndex        =   5
         Top             =   3240
         Visible         =   0   'False
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   688
         Buttons         =   2
         Caption0        =   "&Guardar"
         Tag0            =   "2"
         ToolTipText0    =   "Guardar"
         Caption1        =   "&Cancelar"
         Tag1            =   "8"
         ToolTipText1    =   "Cancelar"
         UserControlHeight=   390
         UserControlWidth=   2700
      End
      Begin MSAdodcLib.Adodc adoGastos2 
         Height          =   330
         Left            =   -73560
         Top             =   3240
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
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmFondoConceptoActivoFijo.frx":57D9
         Height          =   2295
         Left            =   -74760
         OleObjectBlob   =   "frmFondoConceptoActivoFijo.frx":57F3
         TabIndex        =   10
         Top             =   1590
         Width           =   8895
      End
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Index           =   1
         Left            =   6240
         TabIndex        =   17
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
   End
End
Attribute VB_Name = "frmFondoConceptoActivoFijo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrFondo()          As String

Dim strCodFondo         As String
Dim strEstado           As String, strSQL           As String
Dim adoConsulta As ADODB.Recordset
Dim adoGastos As ADODB.Recordset

Private Sub CargarActivoFijo()

    strSQL = "SELECT CodCuenta,RTRIM(DescripCuenta) DescripGasto " & _
        "FROM PlanContable " & _
        "WHERE CodAdministradora='" & gstrCodAdministradora & "' AND IndMovimiento='X' AND " & _
        "(CodCuenta LIKE '3[3-6]%') " & _
        "ORDER BY DescripGasto"
    
    Set adoGastos = New ADODB.Recordset

    With adoGastos
        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open strSQL
    End With
    
    tdgGasto.DataSource = adoGastos
    tdgGasto.Refresh
        
End Sub

Private Sub cboFondo_Click()

    strCodFondo = Valor_Caracter
    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(arrFondo(cboFondo.ListIndex))
    
    Call Buscar
    
End Sub

Private Sub cmdSeleccionar_Click()
    Call Modificar
End Sub

Private Sub Form_Load()

    Call InicializarValores
    Call CargarListas
    Call DarFormato
    Call Buscar
    
    Call ValidarPermisoUsoControl(Trim(gstrLoginUS), Me, Trim(App.Title) + Separador_Codigo_Objeto + _
    gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)
    
    CentrarForm Me
    
End Sub


Public Sub Buscar()
                
    strSQL = "SELECT FCG.CodCuenta,(RTRIM(DescripCuenta)) DescripGasto " & _
        "FROM FondoConceptoActivoFijo FCG JOIN PlanContable PCG ON(PCG.CodCuenta=FCG.CodCuenta AND PCG.CodAdministradora=FCG.CodAdministradora) " & _
        "WHERE CodFondo='" & strCodFondo & "' AND FCG.CodAdministradora='" & gstrCodAdministradora & "' " & _
        "ORDER BY DescripGasto"
                        
    strEstado = Reg_Defecto
    
    Set adoConsulta = New ADODB.Recordset
    
    With adoConsulta
        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open strSQL
    End With
    
    tdgConsulta.DataSource = adoConsulta
    tdgConsulta.Refresh
    
    If adoConsulta.RecordCount > 0 Then strEstado = Reg_Consulta

End Sub
Private Sub DarFormato()

    Dim intCont As Integer
    Dim elemento As Object
    
    For intCont = 0 To (lblDescrip.Count - 1)
        Call FormatoEtiqueta(lblDescrip(intCont), vbLeftJustify)
    Next
    
    For intCont = 0 To (fraGasto.Count - 1)
        Call FormatoMarco(fraGasto(intCont))
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
    CargarControlLista strSQL, cboFondo, arrFondo(), Valor_Caracter
    
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
        
End Sub
Private Sub InicializarValores()
                        
    '*** Valores Iniciales ***
    tabGastos.Tab = 0
    tabGastos.TabEnabled(1) = False
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgConsulta.Columns(1).Width = tdgConsulta.Width * 0.01 * 15
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion(1).FormularioActivo = Me
    
End Sub

Private Sub tabGastos_Click(PreviousTab As Integer)

    Select Case tabGastos.Tab
        Case 1
            If strEstado = Reg_Defecto Then tabGastos.Tab = 0
        
    End Select
    
End Sub

Public Sub Accion(nAccion As ButtonAction)
    
    Select Case nAccion
                
        Case vQuery
            Call Modificar
        Case vSearch
            Call Buscar
        Case vSave
            Call Grabar
        Case vCancel
            Call Cancelar
        Case vExit
            Call Salir
        Case vPrint
            Call SubImprimir(1)
        
    End Select
    
End Sub

Public Sub Salir()

    Unload Me
    
End Sub
Public Sub Cancelar()

    cmdOpcion.Visible = True
    With tabGastos
        .TabEnabled(0) = True
        .TabEnabled(1) = False
        .Tab = 0
    End With
    strEstado = Reg_Consulta
    
End Sub
Public Sub Grabar()

    Dim intContador         As Integer, intRegistro     As Integer
    Dim intAccion           As Integer, lngNumError     As Long
    Dim strNumDetalleFile   As String
    Dim mensaje As String
    If strEstado = Reg_Consulta Then Exit Sub
    
    On Error GoTo CtrlError
            
    If strEstado = Reg_Edicion Then
        Dim adoRegistro         As ADODB.Recordset
        
        frmMainMdi.stbMdi.Panels(3).Text = "Actualizar Concepto de Gastos por Fondo..."
        
        mensaje = Mensaje_Edicion
        
        If MsgBox(mensaje, vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption) <> vbYes Then Exit Sub
        
        Me.MousePointer = vbHourglass
        
        With adoComm
            Set adoRegistro = New ADODB.Recordset
            
            .CommandText = "DELETE FondoConceptoActivoFijo WHERE " & _
                "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
            adoConn.Execute .CommandText
        
            intContador = tdgGasto.SelBookmarks.Count - 1
            adoGastos.MoveFirst

            For intRegistro = 0 To intContador
                adoGastos.MoveFirst
                
                adoGastos.Move CLng(tdgGasto.SelBookmarks(intRegistro) - 1), 0
                tdgGasto.Refresh
                
                '*** Guardar gasto seleccionado ***
                .CommandText = "INSERT INTO FondoConceptoActivoFijo VALUES ('" & _
                    strCodFondo & "','" & gstrCodAdministradora & "','" & _
                    Trim(tdgGasto.Columns(0).Value) & "')"
                adoConn.Execute .CommandText
                
                '*** Obtener secuencial ***
                .CommandText = "SELECT COUNT(*) NumDetalle FROM InversionDetalleFile WHERE CodFile='030'"
                Set adoRegistro = .Execute
                
                If Not adoRegistro.EOF Then
                    If IsNull(adoRegistro("NumDetalle")) Then
                        strNumDetalleFile = "001"
                    Else
                        strNumDetalleFile = Format(CInt(adoRegistro("NumDetalle")) + 1, "000")
                    End If
                Else
                    strNumDetalleFile = "001"
                End If
                adoRegistro.Close
                    
                .CommandText = "SELECT IndVigente FROM InversionDetalleFile " & _
                    "WHERE CodFile='030' AND DescripDetalleFile='" & Trim(tdgGasto.Columns(0).Value) & "'"
                Set adoRegistro = .Execute
                
                If Not adoRegistro.EOF Then
                    If adoRegistro("IndVigente") = Valor_Caracter Then
                        .CommandText = "UPDATE InversionDetalleFile SET IndVigente='X' " & _
                            "WHERE CodFile='030' AND DescripDetalleFile=" & Trim(tdgGasto.Columns(0).Value) & "'"
                        adoConn.Execute .CommandText
                    End If
                Else
                    .CommandText = "INSERT INTO InversionDetalleFile VALUES('" & strNumDetalleFile & "','030','" & _
                        Trim(tdgGasto.Columns(0).Value) & "','','X','')"
                    adoConn.Execute .CommandText
                End If
                adoRegistro.Close: Set adoRegistro = Nothing
            Next
        End With
    
        Me.MousePointer = vbDefault
                        
        MsgBox Mensaje_Edicion_Exitosa, vbExclamation
        
        frmMainMdi.stbMdi.Panels(3).Text = "Acci�n"
        
        cmdOpcion.Visible = True
        With tabGastos
            .TabEnabled(0) = True
            .TabEnabled(1) = False
            .Tab = 0
        End With
        Call Buscar

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
Public Sub Modificar()

    If strCodFondo = Valor_Caracter Then
        MsgBox "No existen fondos definidos...", vbCritical, Me.Caption
        Exit Sub
    End If
    
    If strEstado = Reg_Consulta Or strEstado = Reg_Defecto Then
        strEstado = Reg_Edicion
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabGastos
            .TabEnabled(0) = False
            .TabEnabled(1) = True
            .Tab = 1
        End With
    End If
        
End Sub


Private Sub LlenarFormulario(strModo As String)

    Dim adoRecord   As ADODB.Recordset
    
    Select Case strModo
        Case Reg_Edicion
            lblDescripFondo(1).Caption = Trim(cboFondo.Text)
                                    
            Call CargarActivoFijo
        
            Set adoRecord = New ADODB.Recordset
                                    
            adoComm.CommandText = "SELECT CodCuenta FROM FondoConceptoActivoFijo " & _
                "WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
            Set adoRecord = adoComm.Execute
                        
            Do While Not adoRecord.EOF
                adoGastos.MoveFirst
                
                adoGastos.Find ("CodCuenta='" & adoRecord("CodCuenta") & "'")
                
                tdgGasto.SelBookmarks.Add adoGastos.Bookmark
                
                adoRecord.MoveNext
            Loop
            adoRecord.Close: Set adoRecord = Nothing
    
    End Select
    
End Sub


Public Sub SubImprimir(Index As Integer)

    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    Dim strFechaDesde           As String, strFechaHasta        As String

    If tabGastos.Tab = 1 Then Exit Sub

    Select Case Index
        Case 1
            gstrNameRepo = "AutorizacionActivoFijoGrilla"

            Set frmReporte = New frmVisorReporte

            ReDim aReportParamS(1)
            ReDim aReportParamFn(3)
            ReDim aReportParamF(3)

            aReportParamFn(0) = "Usuario"
            aReportParamFn(1) = "Hora"
            aReportParamFn(2) = "NombreEmpresa"
            aReportParamFn(3) = "Fondo"

            aReportParamF(0) = gstrLogin
            aReportParamF(1) = Format(Time(), "hh:mm:ss")
            aReportParamF(2) = gstrNombreEmpresa & Space(1)
            aReportParamF(3) = Trim(cboFondo.Text)

            aReportParamS(0) = strCodFondo
            aReportParamS(1) = gstrCodAdministradora

    End Select

    gstrSelFrml = Valor_Caracter
    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"

    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())

    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
    frmReporte.Show vbModal

    Set frmReporte = Nothing

    Screen.MousePointer = vbNormal

End Sub

Private Sub tdgConsulta_HeadClick(ByVal ColIndex As Integer)
    Static numColindex As Integer

    tdgConsulta.Splits(0).Columns(numColindex).HeadingStyle.ForegroundPicture = Null

    Call OrdenarDBGrid(ColIndex, adoConsulta, tdgConsulta)
    
    numColindex = ColIndex
End Sub


Private Sub tdgGasto_Click()

End Sub
