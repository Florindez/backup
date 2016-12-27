VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{442CAE95-1D41-47B1-BE83-6995DA3CE254}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmFormulaMant 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Formulas"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   5880
      TabIndex        =   17
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
      Left            =   240
      TabIndex        =   18
      Top             =   4680
      Width           =   4200
      _ExtentX        =   7408
      _ExtentY        =   1296
      Buttons         =   3
      Caption0        =   "&Nuevo"
      Tag0            =   "0"
      ToolTipText0    =   "Nuevo"
      Caption1        =   "&Modificar"
      Tag1            =   "3"
      ToolTipText1    =   "Modificar"
      Caption2        =   "&Buscar"
      Tag2            =   "5"
      ToolTipText2    =   "Buscar"
      UserControlWidth=   4200
   End
   Begin TabDlg.SSTab tabFormula 
      Height          =   4515
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   7964
      _Version        =   393216
      Style           =   1
      Tabs            =   2
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
      TabPicture(0)   =   "frmFormulaMant.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tdgConsulta"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "frmFormulaMant.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraDetalle"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdAccion"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   -70680
         TabIndex        =   1
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
      Begin VB.Frame fraDetalle 
         Height          =   2895
         Left            =   -74640
         TabIndex        =   2
         Top             =   480
         Width           =   6735
         Begin VB.CommandButton cmdCondicion 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   6270
            TabIndex        =   14
            ToolTipText     =   "Buscar Proveedor"
            Top             =   2400
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CheckBox chkCondicion 
            Caption         =   "Tiene Condición"
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
            Height          =   315
            Left            =   1650
            TabIndex        =   13
            Top             =   2040
            Width           =   3315
         End
         Begin VB.CommandButton cmdFormula 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   6270
            TabIndex        =   11
            ToolTipText     =   "Buscar Proveedor"
            Top             =   1230
            Width           =   375
         End
         Begin VB.TextBox txtDescripFormula 
            Height          =   315
            Left            =   1650
            MaxLength       =   60
            TabIndex        =   4
            Top             =   808
            Width           =   4605
         End
         Begin VB.ComboBox cboDatosFormula 
            Height          =   315
            ItemData        =   "frmFormulaMant.frx":0038
            Left            =   1650
            List            =   "frmFormulaMant.frx":003F
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1650
            Width           =   3255
         End
         Begin VB.Label lblCondicionEtiqueta 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Condición"
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
            Left            =   360
            TabIndex        =   16
            Top             =   2430
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label lblCondicion 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1650
            TabIndex        =   15
            Top             =   2400
            Visible         =   0   'False
            Width           =   4590
         End
         Begin VB.Label lblFormula 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1650
            TabIndex        =   12
            Top             =   1230
            Width           =   4590
         End
         Begin VB.Label lblCodFormula 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1650
            TabIndex        =   9
            Top             =   405
            Width           =   2055
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código"
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
            Index           =   1
            Left            =   360
            TabIndex        =   8
            Top             =   420
            Width           =   585
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción"
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
            Left            =   360
            TabIndex        =   7
            Top             =   825
            Width           =   1005
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Formula"
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
            TabIndex        =   6
            Top             =   1260
            Width           =   675
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Datos"
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
            Index           =   4
            Left            =   360
            TabIndex        =   5
            Top             =   1680
            Width           =   510
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmFormulaMant.frx":0056
         Height          =   3585
         Left            =   330
         OleObjectBlob   =   "frmFormulaMant.frx":0070
         TabIndex        =   10
         Top             =   600
         Width           =   6735
      End
   End
End
Attribute VB_Name = "frmFormulaMant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrEstado()             As String, strCodEstado         As String
Dim strEstado               As String, strSQL               As String
Dim arrDatosFormula()       As String, strDatosCodFormula   As String
Dim adoConsulta             As ADODB.Recordset
Dim indSortAsc              As Boolean, indSortDesc         As Boolean

Private Sub cboDatosFormula_Click()
    strDatosCodFormula = Valor_Caracter
    If cboDatosFormula.ListIndex < 0 Then Exit Sub
    
    strDatosCodFormula = Trim(arrDatosFormula(cboDatosFormula.ListIndex))
End Sub

'Private Sub cboEstado_Click()
'
'    strCodEstado = Valor_Caracter
'
'    If cboEstado.ListIndex < 0 Then Exit Sub
'
'    strCodEstado = arrEstado(cboEstado.ListIndex)
'
'End Sub

Private Sub chkCondicion_Click()
lblCondicionEtiqueta.Visible = chkCondicion.Value
lblCondicion.Visible = chkCondicion.Value
cmdCondicion.Visible = chkCondicion.Value
End Sub

Private Sub cmdCondicion_Click()
Dim strRespuesta As String

strRespuesta = lblCondicion.Caption
frmFormulas.mostrarForm strRespuesta, "01", strDatosCodFormula
lblCondicion.Caption = strRespuesta
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

    Dim c As Object
    Dim elemento As Object

    For Each c In Me.Controls
        
        If TypeOf c Is Label Then
            Call FormatoEtiqueta(c, vbLeftJustify)
        End If
    Next
    
    For Each elemento In Me.Controls
    
        If TypeOf elemento Is TDBGrid Then
            Call FormatoGrilla(elemento)
        End If
    
    Next
            
End Sub

Public Sub Buscar()

    Set adoConsulta = New ADODB.Recordset
    
    strSQL = "SELECT CodFormula, DescripFormula, FormulaMonto, indCondicion, FormulaCondicion, CodFormulaDatos " & _
        "FROM Formula"
        
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
Private Sub CargarReportes()

    'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Listado"
    'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    
End Sub
Private Sub CargarListas()
        
      
    '*** Datos Formula
    strSQL = "SELECT CodFormulaDatos CODIGO,DescripFormulaDatos DESCRIP From FormulaDatos ORDER BY DescripFormulaDatos"
    CargarControlLista strSQL, cboDatosFormula, arrDatosFormula(), Valor_Caracter
            
End Sub

Public Sub Accion(nAccion As ButtonAction)
    
    Select Case nAccion
        
        Case vNew
            Call Adicionar
        Case vModify
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
    
    frmMainMdi.stbMdi.Panels(3).Text = "Adicionar Monedas..."
                    
    strEstado = Reg_Adicion
    LlenarFormulario strEstado
    cmdOpcion.Visible = False
    With tabFormula
        .TabEnabled(0) = False
        .Tab = 1
        .TabEnabled(1) = True
    End With
    
End Sub

Public Sub Cancelar()

    cmdOpcion.Visible = True
    With tabFormula
        .TabEnabled(0) = True
        .Tab = 0
        .TabEnabled(1) = False
    End With
    Call Buscar
    
End Sub

Public Sub Eliminar()

End Sub

Public Sub Grabar()

'    Dim adoresult           As ADODB.Recordset, adoRec      As ADODB.Recordset
    Dim intAccion           As Integer, lngNumError         As Long
    Dim strOperacion        As String
    
    If strEstado = Reg_Consulta Then Exit Sub
    
    On Error GoTo CtrlError
    
    If strEstado = Reg_Adicion Or strEstado = Reg_Edicion Then
        If TodoOK() Then
            Me.MousePointer = vbHourglass
            
            strOperacion = "I"
            If strEstado = Reg_Edicion Then strOperacion = "U"
            
            '*** Guardar ***
            With adoComm
                .CommandText = "{ call up_GNManFormula('" & _
                    lblCodFormula.Caption & "','" & Trim(txtDescripFormula.Text) & "','" & _
                    Trim(lblFormula.Caption) & "','" & strDatosCodFormula & "','" & _
                    IIf(chkCondicion.Value, Valor_Indicador, "") & "','" & Trim(lblCondicion.Caption) & "','" & strOperacion & "') }"
                adoConn.Execute .CommandText
            End With
            
            Me.MousePointer = vbDefault
                        
            If strEstado = Reg_Adicion Then
                MsgBox Mensaje_Adicion_Exitosa, vbExclamation
            Else
                MsgBox Mensaje_Edicion_Exitosa, vbExclamation
            End If
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
                        
            cmdOpcion.Visible = True
            With tabFormula
                .TabEnabled(0) = True
                .Tab = 0
                .TabEnabled(1) = False
            End With
            Call Buscar
        End If
    End If
    
'    If strEstado = Reg_Edicion Then
'        If TodoOK() Then
'            Me.MousePointer = vbHourglass
'
'            '*** Guardar ***
'            With adoComm
'                .CommandText = "{ call up_GNManMoneda('" & _
'                    strCodMoneda & "','" & Trim(txtDescripMoneda.Text) & "','" & _
'                    Trim(txtSimbolo.Text) & "','" & strCodMonedaCambio & "','" & _
'                    Trim(txtCodConasev.Text) & "','" & Trim(txtSigno.Text) & "','" & strCodEstado & "','U') }"
'                adoConn.Execute .CommandText
'            End With
'
'            Me.MousePointer = vbDefault
'
'            MsgBox Mensaje_Edicion_Exitosa, vbExclamation
'
'            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
'
'            cmdOpcion.Visible = True
'            With tabMoneda
'                .TabEnabled(0) = True
'                .Tab = 0
'            End With
'            Call Buscar
'        End If
'    End If
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

Private Function TodoOK() As Boolean
        
    TodoOK = False
        
    If Trim(txtDescripFormula.Text) = Valor_Caracter Then
        MsgBox "Debe indicar la descripción", vbCritical
        txtDescripFormula.SetFocus
        Exit Function
    End If
    
    If Trim(lblFormula.Caption) = Valor_Caracter Then
        MsgBox "Debe indicar la formula", vbCritical
        Exit Function
    End If
    
    If Trim(lblCondicion.Caption) = Valor_Caracter And chkCondicion.Value Then
        MsgBox "Debe indicar la formula de la condición", vbCritical
        Exit Function
    End If
    
    '*** Si todo paso OK ***
    TodoOK = True
  
End Function
Public Sub Imprimir()
    
    Call SubImprimir(1)
    
End Sub

Public Sub Modificar()

    If strEstado = Reg_Consulta Then
        strEstado = Reg_Edicion
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabFormula
            .TabEnabled(0) = False
            .Tab = 1
            .TabEnabled(1) = True
        End With
        
    End If
    
End Sub

Private Sub LlenarFormulario(strModo As String)

    Dim adoRegistro     As ADODB.Recordset
    Dim intRegistro     As Integer
    
    Select Case strModo
        Case Reg_Adicion
            
            lblCodFormula.Caption = Valor_Caracter
            txtDescripFormula.Text = Valor_Caracter
            lblFormula.Caption = Valor_Caracter
            lblCondicion.Caption = Valor_Caracter
            cboDatosFormula.ListIndex = -1
            txtDescripFormula.SetFocus
        
        Case Reg_Edicion
        
            lblCodFormula.Caption = Trim(tdgConsulta.Columns(0).Value)

            txtDescripFormula.Text = Trim(tdgConsulta.Columns(1).Value)
            lblFormula.Caption = Trim(tdgConsulta.Columns(2).Value)

            intRegistro = ObtenerItemLista(arrDatosFormula(), "" & Trim(tdgConsulta.Columns(5).Value))
            If intRegistro >= 0 Then cboDatosFormula.ListIndex = intRegistro
            
            chkCondicion.Value = IIf(Trim(tdgConsulta.Columns(3).Value) = Valor_Indicador, 1, 0)
            lblCondicion.Caption = Trim(tdgConsulta.Columns(4).Value)

            txtDescripFormula.SetFocus

    End Select
    
End Sub

Public Sub Salir()

    Unload Me
    
End Sub

Public Sub SubImprimir(Index As Integer)
    
End Sub
Private Sub InicializarValores()

    '*** Valores Iniciales ***
    strEstado = Reg_Defecto
    tabFormula.Tab = 0
    tabFormula.TabEnabled(1) = False
    
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgConsulta.Columns(0).Width = tdgConsulta.Width * 0.01 * 12
    tdgConsulta.Columns(1).Width = tdgConsulta.Width * 0.01 * 32
    tdgConsulta.Columns(2).Width = tdgConsulta.Width * 0.01 * 14
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call OcultarReportes
    Set frmMoneda = Nothing
    
End Sub

Private Sub tabFormula_Click(PreviousTab As Integer)

    Select Case tabFormula.Tab
        Case 1
            If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vModify)
            If strEstado = Reg_Defecto Then tabFormula.Tab = 0
        
    End Select
    
End Sub

Private Sub cmdFormula_Click()
Dim strRespuesta As String

strRespuesta = lblFormula.Caption
frmFormulas.mostrarForm strRespuesta, "01", strDatosCodFormula
lblFormula.Caption = strRespuesta

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
