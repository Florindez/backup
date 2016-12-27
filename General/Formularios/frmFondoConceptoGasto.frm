VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmFondoConceptoGasto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Autorización de Gastos por Fondo"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5250
   ScaleWidth      =   9660
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   7920
      TabIndex        =   10
      Top             =   4440
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
      TabIndex        =   11
      Top             =   4440
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   1296
      Buttons         =   2
      Caption0        =   "&Seleccionar"
      Tag0            =   "3"
      ToolTipText0    =   "Seleccionar"
      Caption1        =   "&Buscar"
      Tag1            =   "5"
      ToolTipText1    =   "Buscar"
      UserControlWidth=   2700
   End
   Begin MSAdodcLib.Adodc adoConsulta 
      Height          =   330
      Left            =   4560
      Top             =   4440
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
      Caption         =   "adoConsulta"
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
   Begin TabDlg.SSTab tabGastos 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   7223
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
      TabPicture(0)   =   "frmFondoConceptoGasto.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "tdgConsulta"
      Tab(0).Control(1)=   "fraGasto(0)"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "frmFondoConceptoGasto.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraGasto(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "adoGastos"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdAccion"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   6240
         TabIndex        =   9
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
      Begin MSAdodcLib.Adodc adoGastos 
         Height          =   330
         Left            =   1440
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
      Begin VB.Frame fraGasto 
         Caption         =   "Selección de Gastos"
         Height          =   2655
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   8895
         Begin TrueOleDBGrid60.TDBGrid tdgGastos 
            Bindings        =   "frmFondoConceptoGasto.frx":0038
            Height          =   1575
            Left            =   240
            OleObjectBlob   =   "frmFondoConceptoGasto.frx":0050
            TabIndex        =   8
            Top             =   840
            Width           =   8415
         End
         Begin VB.Label lblDescripFondo 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1320
            TabIndex        =   7
            Top             =   360
            Width           =   7335
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fondo"
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   6
            Top             =   360
            Width           =   450
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmFondoConceptoGasto.frx":29D6
         Height          =   1935
         Left            =   -74760
         OleObjectBlob   =   "frmFondoConceptoGasto.frx":29F0
         TabIndex        =   2
         Top             =   1560
         Width           =   8895
      End
      Begin VB.Frame fraGasto 
         Caption         =   "Criterios de Búsqueda"
         Height          =   975
         Index           =   0
         Left            =   -74760
         TabIndex        =   1
         Top             =   480
         Width           =   8895
         Begin VB.ComboBox cboFondo 
            Height          =   315
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   360
            Width           =   5535
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fondo"
            Height          =   195
            Index           =   0
            Left            =   960
            TabIndex        =   3
            Top             =   375
            Width           =   450
         End
      End
   End
End
Attribute VB_Name = "frmFondoConceptoGasto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrFondo()          As String

Dim strCodFondo         As String
Dim strEstado           As String, strSQL           As String

Private Sub CargarGastos()

    strSQL = "SELECT CodCuenta,RTRIM(DescripCuenta) DescripGasto " & _
        "FROM PlanContable " & _
        "WHERE CodAdministradora='" & gstrCodAdministradora & "' AND IndMovimiento='X' AND " & _
        "CodCuenta LIKE '6%'" & _
        "ORDER BY DescripGasto"
'    strSQL = "SELECT distinct DC.CodCuenta,RTRIM(PC.DescripCuenta) DescripGasto " & _
'        "FROM PlanContable PC,DinamicaContable DC " & _
'        "WHERE DC.CodAdministradora='" & gstrCodAdministradora & "' AND DC.CodCuenta LIKE '6%' AND " & _
'        "PC.CodAdministradora = DC.CodAdministradora AND DC.CodFile = '099' and PC.CodCuenta = DC.CodCuenta AND PC.NumVersion = dbo.uf_CNObtenerPlanContableVigente() " & _
'        "ORDER BY DescripGasto"
    
    With adoGastos
        .ConnectionString = gstrConnectConsulta
        .RecordSource = strSQL
        .Refresh
    End With
        
    tdgGastos.Refresh
        
End Sub

Private Sub cboFondo_Click()

    strCodFondo = Valor_Caracter
    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(arrFondo(cboFondo.ListIndex))
    
    Call Buscar
    
End Sub


Private Sub Form_Activate()

    Call CargarReportes
    
End Sub


Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Listado de Gastos Permitidos del Fondo"
    
End Sub
Private Sub Form_Deactivate()

    Call OcultarReportes
    
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
                
'    strSQL = "SELECT FCG.CodCuenta,(RTRIM(DescripCuenta)) DescripGasto " & _
'        "FROM FondoConceptoGasto FCG JOIN PlanContable PCG ON(PCG.CodCuenta=FCG.CodCuenta AND PCG.CodAdministradora=FCG.CodAdministradora) " & _
'        "WHERE CodFondo='" & strCodFondo & "' AND FCG.CodAdministradora='" & gstrCodAdministradora & "' " & _
'        "ORDER BY DescripGasto"
                        
    strSQL = "{ call up_GNLstFondoConceptoGasto ('" & strCodFondo & "','" & gstrCodAdministradora & "')}"
                        
    strEstado = Reg_Defecto
    With adoConsulta
        .ConnectionString = gstrConnectConsulta
        .RecordSource = strSQL
        .Refresh
    End With
        
    tdgConsulta.Refresh
    
    If adoConsulta.Recordset.RecordCount > 0 Then strEstado = Reg_Consulta

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
    Set cmdAccion.FormularioActivo = Me
    
End Sub
Private Sub Form_Unload(Cancel As Integer)

    Call OcultarReportes
    Set frmFondoComision = Nothing
    
End Sub


Private Sub tabGastos_Click(PreviousTab As Integer)

    Select Case tabGastos.Tab
        Case 1
            If strEstado = Reg_Defecto Then tabGastos.Tab = 0
        
    End Select
    
End Sub

Public Sub Accion(nAccion As ButtonAction)
    
    Select Case nAccion
                
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
    
    If strEstado = Reg_Consulta Then Exit Sub
    
    On Error GoTo CtrlError
            
    If strEstado = Reg_Edicion Then
        Dim adoRegistro         As ADODB.Recordset
        
        frmMainMdi.stbMdi.Panels(3).Text = "Actualizar Concepto de Gastos por Fondo..."
        
        Me.MousePointer = vbHourglass
        
        With adoComm
            Set adoRegistro = New ADODB.Recordset
            
            .CommandText = "DELETE FondoConceptoGasto WHERE " & _
                "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
            adoConn.Execute .CommandText
        
            intContador = tdgGastos.SelBookmarks.Count - 1
            adoGastos.Recordset.MoveFirst

            For intRegistro = 0 To intContador
                adoGastos.Recordset.MoveFirst
                
                adoGastos.Recordset.Move CLng(tdgGastos.SelBookmarks(intRegistro) - 1), 0
                tdgGastos.Refresh
                
                '*** Guardar gasto seleccionado ***
                .CommandText = "INSERT INTO FondoConceptoGasto VALUES ('" & _
                    strCodFondo & "','" & gstrCodAdministradora & "','" & _
                    Trim(tdgGastos.Columns(0).Value) & "')"
                adoConn.Execute .CommandText
                
                '*** Obtener secuencial ***
                .CommandText = "SELECT MAX(CodDetalleFile) NumDetalle FROM InversionDetalleFile WHERE CodFile='099'"
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
                    "WHERE CodFile='099' AND DescripDetalleFile='" & Trim(tdgGastos.Columns(0).Value) & "'"
                Set adoRegistro = .Execute
                
                If Not adoRegistro.EOF Then
                    If adoRegistro("IndVigente") = Valor_Caracter Then
                        .CommandText = "UPDATE InversionDetalleFile SET IndVigente='X' " & _
                            "WHERE CodFile='099' AND DescripDetalleFile=" & Trim(tdgGastos.Columns(0).Value) & "'"
                        adoConn.Execute .CommandText
                    End If
                Else
                    .CommandText = "INSERT INTO InversionDetalleFile VALUES('" & strNumDetalleFile & "','099','" & _
                        Trim(tdgGastos.Columns(0).Value) & "','','X','')"
                    adoConn.Execute .CommandText
                End If
                adoRegistro.Close: Set adoRegistro = Nothing
            Next
        End With
    
        Me.MousePointer = vbDefault
                        
        MsgBox Mensaje_Edicion_Exitosa, vbExclamation
        
        frmMainMdi.stbMdi.Panels(3).Text = "Acción"
        
        cmdOpcion.Visible = True
        With tabGastos
            .TabEnabled(0) = True
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
Public Sub Imprimir()

    
    
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
            lblDescripFondo.Caption = Trim(cboFondo.Text)
                                    
            Call CargarGastos
        
            Set adoRecord = New ADODB.Recordset
                                    
            adoComm.CommandText = "SELECT CodCuenta FROM FondoConceptoGasto " & _
                "WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
            Set adoRecord = adoComm.Execute
                        
            Do While Not adoRecord.EOF
                adoGastos.Recordset.MoveFirst
                
                adoGastos.Recordset.Find ("CodCuenta='" & adoRecord("CodCuenta") & "'")
                
                If tdgGastos.EOF = False Then
                    tdgGastos.SelBookmarks.Add adoGastos.Recordset.Bookmark
                End If
                
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
            gstrNameRepo = "FondoConceptoGasto"

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

