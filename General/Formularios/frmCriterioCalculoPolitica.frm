VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmCriterioCalculoPolitica 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Criterio Comparación para la Política de Inversión"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4905
   ScaleWidth      =   9525
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   7680
      TabIndex        =   3
      Top             =   4080
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
      Left            =   720
      TabIndex        =   2
      Top             =   4080
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
   Begin MSAdodcLib.Adodc adoConsulta 
      Height          =   330
      Left            =   4560
      Top             =   4080
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
   Begin TabDlg.SSTab tabCriterio 
      Height          =   3735
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   6588
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
      TabPicture(0)   =   "frmCriterioCalculoPolitica.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tdgConsulta"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraCriterio"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "frmCriterioCalculoPolitica.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraDetalle"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdAccion"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   -69120
         TabIndex        =   5
         Top             =   2880
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
      Begin VB.Frame fraCriterio 
         Caption         =   "Criterios de Búsqueda"
         Height          =   1215
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   8535
         Begin VB.ComboBox cboLimite 
            Height          =   315
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   720
            Width           =   4815
         End
         Begin VB.ComboBox cboTipoReglamento 
            Height          =   315
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   360
            Width           =   4815
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Limite"
            Height          =   195
            Index           =   4
            Left            =   480
            TabIndex        =   16
            Top             =   720
            Width           =   405
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Reglamento"
            Height          =   195
            Index           =   0
            Left            =   480
            TabIndex        =   15
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame fraDetalle 
         Height          =   2295
         Left            =   -74760
         TabIndex        =   8
         Top             =   480
         Width           =   8535
         Begin VB.ComboBox cboCriterio 
            Height          =   315
            Left            =   2400
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1680
            Width           =   5655
         End
         Begin VB.Label lblLimite 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2400
            TabIndex        =   18
            Top             =   1320
            Width           =   5655
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Limite"
            Height          =   195
            Index           =   5
            Left            =   360
            TabIndex        =   17
            Top             =   1320
            Width           =   405
         End
         Begin VB.Label lblTipoReglamento 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2400
            TabIndex        =   13
            Top             =   480
            Width           =   5655
         End
         Begin VB.Label lblDescripFondo 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2400
            TabIndex        =   12
            Top             =   930
            Width           =   5655
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Criterio Comparación"
            Height          =   195
            Index           =   3
            Left            =   360
            TabIndex        =   11
            Top             =   1635
            Width           =   1455
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fondo"
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   10
            Top             =   956
            Width           =   450
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Reglamento"
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   9
            Top             =   500
            Width           =   1215
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmCriterioCalculoPolitica.frx":0038
         Height          =   1455
         Left            =   240
         OleObjectBlob   =   "frmCriterioCalculoPolitica.frx":0052
         TabIndex        =   7
         Top             =   1800
         Width           =   8535
      End
   End
End
Attribute VB_Name = "frmCriterioCalculoPolitica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrTipoReglamento()     As String, arrCriterio()        As String
Dim arrEstructura()         As String

Dim strCodTipoReglamento    As String, strCodCriterio       As String
Dim strCodFondo             As String, strCodEstructura     As String
Dim strEstado               As String, strSQL               As String

Private Sub cboCriterio_Click()

    strCodCriterio = Valor_Caracter
    If cboCriterio.ListIndex < 0 Then Exit Sub
    
    strCodCriterio = Trim(arrCriterio(cboCriterio.ListIndex))
    
End Sub


Private Sub cboLimite_Click()

    strCodEstructura = Valor_Caracter
    If cboLimite.ListIndex < 0 Then Exit Sub
    
    strCodEstructura = Trim(arrEstructura(cboLimite.ListIndex))
    
    Call Buscar
    
End Sub


Private Sub cboTipoReglamento_Click()

    strCodTipoReglamento = Valor_Caracter
    If cboTipoReglamento.ListIndex < 0 Then Exit Sub
    
    strCodTipoReglamento = Trim(arrTipoReglamento(cboTipoReglamento.ListIndex))
    
    strSQL = "SELECT CodEstructura CODIGO, DescripEstructura DESCRIP FROM LimiteReglamentoEstructura " & _
        "WHERE CodReglamento='" & strCodTipoReglamento & "' ORDER BY DescripEstructura"
    CargarControlLista strSQL, cboLimite, arrEstructura(), Valor_Caracter
    
    If cboLimite.ListCount > 0 Then cboLimite.ListIndex = 0
            
End Sub


Private Sub Form_Activate()

    Call CargarReportes
    
End Sub

Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Criterio Comparación"
    
End Sub
Public Sub SubImprimir(Index As Integer)

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
    With tabCriterio
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
            
    If strEstado = Reg_Edicion Then
        
        frmMainMdi.stbMdi.Panels(3).Text = "Actualizar Criterio Comparación para la Política de Inversión..."
        
        Me.MousePointer = vbHourglass
        
        With adoComm
            .CommandText = "UPDATE LimiteReglamento SET " & _
                "CodCriterio='" & strCodCriterio & "' " & _
                "WHERE CodReglamento='" & strCodTipoReglamento & "' AND CodEstructura='" & strCodEstructura & "' AND " & _
                "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
            adoConn.Execute .CommandText
        End With
    
        Me.MousePointer = vbDefault
                        
        MsgBox Mensaje_Edicion_Exitosa, vbExclamation
        
        frmMainMdi.stbMdi.Panels(3).Text = "Acción"
        
        cmdOpcion.Visible = True
        With tabCriterio
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
                
    strSQL = "SELECT LR.CodFondo,LR.CodCriterio,DescripCriterio,isnull(DescripFondo,'TODOS') DescripFondo " & _
        "FROM LimiteReglamento LR LEFT JOIN CriterioBaseReglamento CBR ON(CBR.CodCriterio=LR.CodCriterio) " & _
        "LEFT JOIN Fondo ON(Fondo.CodFondo=LR.CodFondo) " & _
        "WHERE CodReglamento='" & strCodTipoReglamento & "' AND CodEstructura='" & strCodEstructura & "' " & _
        "ORDER BY DescripFondo"
                        
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
        With tabCriterio
            .TabEnabled(0) = False
            .Tab = 1
        End With
    End If
        
End Sub
Private Sub LlenarFormulario(strModo As String)

    Dim intRegistro     As Integer
    
    Select Case strModo
        Case Reg_Edicion
            lblTipoReglamento.Caption = Trim(cboTipoReglamento.Text)
            lblDescripFondo.Caption = tdgConsulta.Columns(0).Value
            strCodFondo = tdgConsulta.Columns(2).Value
            lblLimite.Caption = Trim(cboLimite.Text)
            
            If cboCriterio.ListCount > 0 Then cboCriterio.ListIndex = 0
            
            intRegistro = ObtenerItemLista(arrCriterio(), tdgConsulta.Columns(3).Value)
            If intRegistro >= 0 Then cboCriterio.ListIndex = intRegistro

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
    
    '*** Reglamentos ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPREG' AND ValorParametro='X' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboTipoReglamento, arrTipoReglamento(), Valor_Caracter
    
    If cboTipoReglamento.ListCount > 0 Then cboTipoReglamento.ListIndex = 0
    
    '*** Criterios Comparación ***
    strSQL = "SELECT CodCriterio CODIGO,DescripCriterio DESCRIP FROM CriterioBaseReglamento ORDER BY DescripCriterio"
    CargarControlLista strSQL, cboCriterio, arrCriterio(), Sel_Defecto
    
End Sub
Private Sub InicializarValores()
                        
    '*** Valores Iniciales ***
    tabCriterio.Tab = 0
    
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgConsulta.Columns(2).Width = tdgConsulta.Width * 0.01 * 50
    tdgConsulta.Columns(3).Width = tdgConsulta.Width * 0.01 * 40
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    
End Sub

Private Sub tabCriterio_Click(PreviousTab As Integer)

    Select Case tabCriterio.Tab
        Case 1
            If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vModify)
            If strEstado = Reg_Defecto Then tabCriterio.Tab = 0
        
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









