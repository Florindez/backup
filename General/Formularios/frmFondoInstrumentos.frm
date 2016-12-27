VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{442CAE95-1D41-47B1-BE83-6995DA3CE254}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmFondoInstrumentos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inversiones del Fondo"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   6585
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   4800
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
      Left            =   480
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
   Begin TabDlg.SSTab tabInstrumentos 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6255
      _ExtentX        =   11033
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
      TabPicture(0)   =   "frmFondoInstrumentos.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraInversion(0)"
      Tab(0).Control(1)=   "tdgConsulta"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "frmFondoInstrumentos.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraInversion(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "adoInstrumentos"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdAccion"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   3120
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
      Begin MSAdodcLib.Adodc adoInstrumentos 
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
      Begin VB.Frame fraInversion 
         Caption         =   "Selección de Inversión"
         Height          =   2655
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   5775
         Begin TrueOleDBGrid60.TDBGrid tdgInstrumentos 
            Bindings        =   "frmFondoInstrumentos.frx":0038
            Height          =   1575
            Left            =   240
            OleObjectBlob   =   "frmFondoInstrumentos.frx":0056
            TabIndex        =   8
            Top             =   840
            Width           =   5295
         End
         Begin VB.Label lblDescripFondo 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1320
            TabIndex        =   7
            Top             =   360
            Width           =   4215
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
         Bindings        =   "frmFondoInstrumentos.frx":28A6
         Height          =   1935
         Left            =   -74760
         OleObjectBlob   =   "frmFondoInstrumentos.frx":28C0
         TabIndex        =   2
         Top             =   1560
         Width           =   5775
      End
      Begin VB.Frame fraInversion 
         Caption         =   "Criterios de Búsqueda"
         Height          =   975
         Index           =   0
         Left            =   -74760
         TabIndex        =   1
         Top             =   480
         Width           =   5775
         Begin VB.ComboBox cboFondo 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   360
            Width           =   4095
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fondo"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   3
            Top             =   375
            Width           =   450
         End
      End
   End
End
Attribute VB_Name = "frmFondoInstrumentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrFondo()          As String
Dim strCodFondo         As String
Dim strEstado           As String, strSQL           As String
Dim adoConsulta         As ADODB.Recordset
Dim indSortAsc          As Boolean, indSortDesc     As Boolean

Private Sub CargarInstrumentos()

    '*** Tipo de Instrumento ***
    strSQL = "SELECT CodFile,DescripFile FROM InversionFile WHERE IndInstrumento='X' AND IndVigente='X' ORDER BY DescripFile"
    
    With adoInstrumentos
        .ConnectionString = gstrConnectConsulta
        .RecordSource = strSQL
        .Refresh
    End With
        
    tdgInstrumentos.Refresh
        
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
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Listado de Inversiones Permitidas del Fondo"
    
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

    Set adoConsulta = New ADODB.Recordset
                
    strSQL = "SELECT FIF.CodFile,CodDetalleFile,DescripFile " & _
        "FROM FondoInversionFile FIF JOIN InversionFile IVF ON(IVF.CodFile=FIF.CodFile) " & _
        "WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' " & _
        "ORDER BY DescripFile"
                        
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

Private Sub DarFormato()

    Dim intCont As Integer
    Dim elemento As Object
    
    For intCont = 0 To (lblDescrip.Count - 1)
        Call FormatoEtiqueta(lblDescrip(intCont), vbLeftJustify)
    Next
    
    For intCont = 0 To (fraInversion.Count - 1)
        Call FormatoMarco(fraInversion(intCont))
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
    tabInstrumentos.Tab = 0
    tabInstrumentos.TabEnabled(1) = False
    
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgConsulta.Columns(0).Width = tdgConsulta.Width * 0.01 * 9
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me

End Sub
Private Sub Form_Unload(Cancel As Integer)

    Call OcultarReportes
    Set frmFondoComision = Nothing
    
End Sub


Private Sub tabInstrumentos_Click(PreviousTab As Integer)

    Select Case tabInstrumentos.Tab
        Case 1
            If strEstado = Reg_Defecto Then tabInstrumentos.Tab = 0
        
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
    With tabInstrumentos
        .TabEnabled(0) = True
        .TabEnabled(1) = False
        .Tab = 0
    End With
    strEstado = Reg_Consulta
    
End Sub
Public Sub Grabar()

    Dim intContador     As Integer, intRegistro     As Integer
    Dim intAccion       As Integer, lngNumError     As Long
    
    If strEstado = Reg_Consulta Then Exit Sub
    
    On Error GoTo CtrlError
            
    If strEstado = Reg_Edicion Then
        
        frmMainMdi.stbMdi.Panels(3).Text = "Actualizar Instrumentos por Fondo..."
        
        Me.MousePointer = vbHourglass
        
        With adoComm
            .CommandText = "DELETE FondoInversionFile WHERE " & _
                "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
            adoConn.Execute .CommandText
        
            intContador = tdgInstrumentos.SelBookmarks.Count - 1
            adoInstrumentos.Recordset.MoveFirst

            For intRegistro = 0 To intContador
                adoInstrumentos.Recordset.MoveFirst
                
                adoInstrumentos.Recordset.Move CLng(tdgInstrumentos.SelBookmarks(intRegistro) - 1), 0
                'tdgInstrumentos.Row = tdgInstrumentos.SelBookmarks(intRegistro) - 1
                tdgInstrumentos.Refresh
                
                .CommandText = "INSERT INTO FondoInversionFile VALUES ('" & _
                    strCodFondo & "','" & gstrCodAdministradora & "','" & _
                    Trim(tdgInstrumentos.Columns(0).Value) & "','000')"
                adoConn.Execute .CommandText
                        
            Next
        End With
    
        Me.MousePointer = vbDefault
                        
        MsgBox Mensaje_Edicion_Exitosa, vbExclamation
        
        frmMainMdi.stbMdi.Panels(3).Text = "Acción"
        
        cmdOpcion.Visible = True
        With tabInstrumentos
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
        With tabInstrumentos
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
                                    
            Call CargarInstrumentos
        
            Set adoRecord = New ADODB.Recordset
                                    
            adoComm.CommandText = "SELECT CodFile FROM FondoInversionFile " & _
                "WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
            Set adoRecord = adoComm.Execute
                        
            Do While Not adoRecord.EOF
                adoInstrumentos.Recordset.MoveFirst
                
                adoInstrumentos.Recordset.Find ("CodFile='" & adoRecord("CodFile") & "'")
                
                tdgInstrumentos.SelBookmarks.Add adoInstrumentos.Recordset.Bookmark
                
                adoRecord.MoveNext
            Loop
            adoRecord.Close: Set adoRecord = Nothing
    
    End Select
    
End Sub


Public Sub SubImprimir(Index As Integer)

    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    Dim strFechaDesde           As String, strFechaHasta        As String

    If tabInstrumentos.Tab = 1 Then Exit Sub

    Select Case Index
        Case 1
            gstrNameRepo = "FondoInstrumento"

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

