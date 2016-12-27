VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmClasificacionInversion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clasificación de Instrumentos"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5460
   ScaleWidth      =   10830
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   8760
      TabIndex        =   2
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
      Left            =   840
      TabIndex        =   1
      Top             =   4680
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
      Caption0        =   "&Modificar"
      Tag0            =   "3"
      ToolTipText0    =   "Modificar"
      UserControlWidth=   1200
   End
   Begin MSAdodcLib.Adodc adoConsulta 
      Height          =   330
      Left            =   4560
      Top             =   4800
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
   Begin TabDlg.SSTab tabClasificacion 
      Height          =   4335
      Left            =   165
      TabIndex        =   5
      Top             =   240
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   7646
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
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
      TabPicture(0)   =   "frmClasificacionInversion.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tdgConsulta"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraClasificacion"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "frmClasificacionInversion.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraDetalle"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "adoInstrumentos"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdAccion"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   -67920
         TabIndex        =   7
         Top             =   3480
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
         Left            =   -72360
         Top             =   3600
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
      Begin VB.Frame fraClasificacion 
         Height          =   855
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   9975
         Begin VB.ComboBox cboFondo 
            Height          =   315
            Left            =   2280
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   360
            Width           =   6735
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fondo"
            Height          =   195
            Index           =   4
            Left            =   840
            TabIndex        =   14
            Top             =   375
            Width           =   450
         End
      End
      Begin VB.Frame fraDetalle 
         Height          =   2895
         Left            =   -74760
         TabIndex        =   6
         Top             =   480
         Width           =   9975
         Begin VB.ComboBox cboSubRiesgoLP 
            Height          =   315
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   2160
            Width           =   2175
         End
         Begin VB.ComboBox cboSubRiesgoCP 
            Height          =   315
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1800
            Width           =   2175
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Clasificación"
            Height          =   195
            Index           =   1
            Left            =   480
            TabIndex        =   17
            Top             =   1440
            Width           =   885
         End
         Begin VB.Label lblInstrumento 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2160
            TabIndex        =   16
            Top             =   840
            Width           =   7215
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Instrumento"
            Height          =   195
            Index           =   5
            Left            =   480
            TabIndex        =   15
            Top             =   840
            Width           =   825
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Máxima"
            Height          =   195
            Index           =   3
            Left            =   480
            TabIndex        =   12
            Top             =   2280
            Width           =   540
         End
         Begin VB.Label lblFondo 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2160
            TabIndex        =   10
            Top             =   480
            Width           =   7215
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Minima"
            Height          =   195
            Index           =   2
            Left            =   480
            TabIndex        =   9
            Top             =   1875
            Width           =   495
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fondo"
            Height          =   195
            Index           =   0
            Left            =   480
            TabIndex        =   8
            Top             =   495
            Width           =   450
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmClasificacionInversion.frx":0038
         Height          =   2535
         Left            =   240
         OleObjectBlob   =   "frmClasificacionInversion.frx":0052
         TabIndex        =   13
         Top             =   1440
         Width           =   9975
      End
   End
End
Attribute VB_Name = "frmClasificacionInversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrFile()               As String, arrSubRiesgoCP()     As String
Dim arrSubRiesgoLP()        As String, arrFondo()           As String
Dim arrInstrumento()        As String
Dim strCodFile              As String, strCodSubRiesgoCP    As String
Dim strCodSubRiesgoLP       As String, strCodFondo          As String
Dim strCodLimite            As String, strCodRiesgoCP       As String
Dim strCodRiesgoLP          As String
Dim strEstado               As String, strSQL               As String

Private Sub cboFondo_Click()

    strCodFondo = Valor_Caracter
    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(arrFondo(cboFondo.ListIndex))
    
    Call Buscar
    
End Sub

Private Sub cboSubRiesgoCP_Click()

    strCodSubRiesgoCP = Valor_Caracter
    If cboSubRiesgoCP.ListIndex < 0 Then Exit Sub
    
    strCodSubRiesgoCP = Trim(arrSubRiesgoCP(cboSubRiesgoCP.ListIndex))
    
End Sub

Private Sub cboSubRiesgoLP_Click()

    strCodSubRiesgoLP = Valor_Caracter
    If cboSubRiesgoLP.ListIndex < 0 Then Exit Sub
    
    strCodSubRiesgoLP = Trim(arrSubRiesgoLP(cboSubRiesgoLP.ListIndex))
    
End Sub

Private Sub Form_Activate()

    Call CargarReportes
    
End Sub

Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Clasificación de Instrumentos"
    
End Sub

Public Sub SubImprimir(Index As Integer)

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

Public Sub Adicionar()

    If strCodFondo = Valor_Caracter Then
        MsgBox "No existen fondos definidos...", vbCritical, Me.Caption
        Exit Sub
    End If
    
    frmMainMdi.stbMdi.Panels(3).Text = "Adicionar Instrumentos..."
                    
    strEstado = Reg_Adicion
    LlenarFormulario strEstado
    cmdOpcion.Visible = False
    With tabClasificacion
        .TabEnabled(0) = False
        .Tab = 1
    End With
    
End Sub

Public Sub Salir()

    Unload Me
    
End Sub

Public Sub Cancelar()

    cmdOpcion.Visible = True
    With tabClasificacion
        .TabEnabled(0) = True
        .Tab = 0
    End With
    strEstado = Reg_Consulta
    
End Sub

Public Sub Grabar()

    Dim intContador     As Integer, intRegistro     As Integer
    Dim intAccion       As Integer, lngNumError     As Long
    
    If strEstado = Reg_Consulta Then Exit Sub
    
    On Error GoTo CtrlError
            
    If strEstado = Reg_Adicion Then
        
        frmMainMdi.stbMdi.Panels(3).Text = "Adicionar Instrumento de Inversión..."
        
        Me.MousePointer = vbHourglass
        
        With adoComm
            
                
'                .CommandText = "INSERT INTO FondoInversionFile VALUES ('" & _
'                    strCodFondo & "','" & gstrCodAdministradora & "','" & _
'                    'Trim(tdgInstrumentos.Columns(0).Value) & "','000','02','02','" & _
'                    strCodLimite & "','" & strCodRiesgoCP & "','" & strCodSubRiesgoCP & "','" & _
'                    strCodRiesgoLP & "','" & strCodSubRiesgoLP & "')"
'                adoConn.Execute .CommandText
                        
            
                    
        End With
    
        Me.MousePointer = vbDefault
                        
        MsgBox Mensaje_Adicion_Exitosa, vbExclamation
        
        frmMainMdi.stbMdi.Panels(3).Text = "Acción"
        
        cmdOpcion.Visible = True
        With tabClasificacion
            .TabEnabled(0) = True
            .Tab = 0
        End With
        Call Buscar

    End If
    
    If strEstado = Reg_Edicion Then
        
        frmMainMdi.stbMdi.Panels(3).Text = "Actualizar Instrumento de Inversión..."
        
        Me.MousePointer = vbHourglass
        
        With adoComm
'            .CommandText = "DELETE FondoInversionFile WHERE " & _
'                "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
'                "CodReglamento='02' AND CodEstructura='02' AND " & _
'                "CodLimite='" & tdgConsulta.Columns(3).Value & "'"
'            adoConn.Execute .CommandText
        
            
                    
        End With
    
        Me.MousePointer = vbDefault
                        
        MsgBox Mensaje_Edicion_Exitosa, vbExclamation
        
        frmMainMdi.stbMdi.Panels(3).Text = "Acción"
        
        cmdOpcion.Visible = True
        With tabClasificacion
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

Public Sub Buscar()
                
    strSQL = "SELECT DISTINCT LRED.CodLimite,CodRiesgoCP,CodSubRiesgoCP,CodRiesgoLP,CodSubRiesgoLP,DescripLimite " & _
        "FROM LimiteReglamentoEstructuraDetalle LRED LEFT JOIN InversionFileClasificacion IFC " & _
        "ON(IFC.CodEstructura=LRED.CodEstructura AND IFC.CodLimite=LRED.CodLimite AND CodReglamento='02' AND CodAdministradora='" & gstrCodAdministradora & "') " & _
        "WHERE LRED.CodEstructura='02' AND IndTitulo<>'X' " & _
        "ORDER BY DescripLimite"

    strEstado = Reg_Defecto
    With adoConsulta
        .ConnectionString = gstrConnectConsulta
        .RecordSource = strSQL
        .Refresh
    End With

    tdgConsulta.Refresh

    If adoConsulta.Recordset.RecordCount > 0 Then strEstado = Reg_Consulta

End Sub

Public Sub Modificar()
        
    If strEstado = Reg_Consulta Then
        strEstado = Reg_Edicion
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabClasificacion
            .TabEnabled(0) = False
            .Tab = 1
        End With
    End If
        
End Sub

Private Sub LlenarFormulario(strModo As String)

    Dim adoRecord       As ADODB.Recordset
    Dim intRegistro     As Integer
    
    Select Case strModo
        Case Reg_Edicion
            lblFondo.Caption = Trim(cboFondo.Text)
            lblInstrumento.Caption = tdgConsulta.Columns(0).Value
            
                                    
            
                        
    End Select
    
End Sub

Private Sub Form_Deactivate()

    Call OcultarReportes
    
End Sub

Private Sub Form_Load()

    Call InicializarValores
    Call CargarListas
    Call DarFormato
    Call Buscar
    
    CentrarForm Me
    
End Sub

Private Sub DarFormato()

    Dim intCont As Integer
    
    For intCont = 0 To (lblDescrip.Count - 1)
        Call FormatoEtiqueta(lblDescrip(intCont), vbLeftJustify)
    Next
            
End Sub

Private Sub CargarListas()
            
    '*** Fondos ***
    strSQL = "{ call up_ACSelDatosParametro(74,'" & gstrCodAdministradora & "','" & gstrCodFondoContable & "') }"
    CargarControlLista strSQL, cboFondo, arrFondo(), Valor_Caracter
    
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
        
    '*** SubRiesgos ***
    strSQL = "SELECT (CodRiesgo + EquivalenciaRiesgo) CODIGO, CodSubRiesgo DESCRIP FROM ClasificacionRiesgoDetalle WHERE CodCategoria='01'"
    CargarControlLista strSQL, cboSubRiesgoCP, arrSubRiesgoCP(), Sel_Defecto
    
    If cboSubRiesgoCP.ListCount > 0 Then cboSubRiesgoCP.ListIndex = 0
    
    strSQL = "SELECT (CodRiesgo + EquivalenciaRiesgo) CODIGO, CodSubRiesgo DESCRIP FROM ClasificacionRiesgoDetalle WHERE CodCategoria='02'"
    CargarControlLista strSQL, cboSubRiesgoLP, arrSubRiesgoLP(), Sel_Defecto
    
    If cboSubRiesgoLP.ListCount > 0 Then cboSubRiesgoLP.ListIndex = 0
        
End Sub

Private Sub InicializarValores()
                        
    '*** Valores Iniciales ***
    tabClasificacion.Tab = 0
    
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgConsulta.Columns(0).Width = tdgConsulta.Width * 0.01 * 60
    tdgConsulta.Columns(1).Width = tdgConsulta.Width * 0.01 * 14
    tdgConsulta.Columns(2).Width = tdgConsulta.Width * 0.01 * 14
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    
End Sub

Private Sub tabClasificacion_Click(PreviousTab As Integer)

    Select Case tabClasificacion.Tab
        Case 1
            If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vModify)
            If strEstado = Reg_Defecto Then tabClasificacion.Tab = 0
        
    End Select
    
End Sub

Private Sub tdgConsulta_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)

    If ColIndex = 2 Then
        Call DarFormatoValor(Value, Decimales_Tasa2)
    End If
    
    If ColIndex = 3 Then
        Call DarFormatoValor(Value, Decimales_Tasa2)
    End If
    
End Sub

