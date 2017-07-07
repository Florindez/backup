VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmAsesores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asesores internos y externos"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9015
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   9015
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab tabInstitucion 
      Height          =   6645
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   11721
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   4
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
      TabPicture(0)   =   "frmAsesores.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "tdgConsulta"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Datos Básicos"
      TabPicture(1)   =   "frmAsesores.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraDatos"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame fraDatos 
         Height          =   6075
         Left            =   270
         TabIndex        =   2
         Top             =   390
         Width           =   8295
         Begin VB.CommandButton cmdAgregar 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7620
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Agregar detalle"
            Top             =   3240
            Width           =   375
         End
         Begin VB.CommandButton cmdQuitar 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7620
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Quitar detalle"
            Top             =   3690
            Width           =   375
         End
         Begin VB.TextBox txtTasaComision 
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
            Left            =   1770
            TabIndex        =   17
            Top             =   2820
            Width           =   2220
         End
         Begin VB.TextBox txtApellidoMat 
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
            Left            =   1770
            MaxLength       =   200
            TabIndex        =   13
            Top             =   1620
            Width           =   3270
         End
         Begin VB.TextBox txtApellidoPat 
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
            Left            =   1770
            MaxLength       =   200
            TabIndex        =   12
            Top             =   1230
            Width           =   3270
         End
         Begin VB.TextBox txtRuc 
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
            Left            =   -3120
            TabIndex        =   7
            Top             =   5400
            Visible         =   0   'False
            Width           =   390
         End
         Begin VB.TextBox txtNombre 
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
            Left            =   1770
            MaxLength       =   200
            TabIndex        =   6
            Top             =   825
            Width           =   3270
         End
         Begin VB.ComboBox cboTipoIdentidad 
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
            Left            =   1770
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   2025
            Width           =   3270
         End
         Begin VB.TextBox txtNumIdentidad 
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
            Left            =   1770
            TabIndex        =   4
            Top             =   2415
            Width           =   2220
         End
         Begin VB.ComboBox cboTipoAsesor 
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
            Left            =   1770
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   390
            Width           =   2550
         End
         Begin TrueOleDBGrid60.TDBGrid tdgClientes 
            Bindings        =   "frmAsesores.frx":0038
            Height          =   2595
            Left            =   240
            OleObjectBlob   =   "frmAsesores.frx":0052
            TabIndex        =   18
            Top             =   3270
            Width           =   7275
         End
         Begin VB.Label lblInstitucion 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tasa de Comisión"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   4
            Left            =   270
            TabIndex        =   16
            Top             =   2910
            Width           =   1335
         End
         Begin VB.Label lblInstitucion 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Apellido Materno"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   10
            Left            =   270
            TabIndex        =   15
            Top             =   1710
            Width           =   1455
         End
         Begin VB.Label lblInstitucion 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Apellido Paterno"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   270
            TabIndex        =   14
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label lblInstitucion 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Asesor"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   270
            TabIndex        =   11
            Top             =   480
            Width           =   840
         End
         Begin VB.Label lblInstitucion 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombres"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   270
            TabIndex        =   10
            Top             =   900
            Width           =   630
         End
         Begin VB.Label lblInstitucion 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo ID."
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   3
            Left            =   270
            TabIndex        =   9
            Top             =   2100
            Width           =   1215
         End
         Begin VB.Label lblInstitucion 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nro.ID."
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   32
            Left            =   270
            TabIndex        =   8
            Top             =   2490
            Width           =   1215
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmAsesores.frx":28BA
         Height          =   4425
         Left            =   -74730
         OleObjectBlob   =   "frmAsesores.frx":28D4
         TabIndex        =   1
         Top             =   510
         Width           =   7245
      End
   End
   Begin TAMControls2.ucBotonEdicion2 cmdOpcion 
      Height          =   735
      Left            =   270
      TabIndex        =   21
      Top             =   6840
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   1296
      Buttons         =   2
      Caption0        =   "&Nuevo"
      Tag0            =   "0"
      ToolTipText0    =   "Nuevo"
      Caption1        =   "&Modificar"
      Tag1            =   "3"
      ToolTipText1    =   "Modificar"
      UserControlWidth=   2700
   End
   Begin TAMControls2.ucBotonEdicion2 cmdAccion 
      Height          =   735
      Left            =   3150
      TabIndex        =   22
      Top             =   6840
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
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   7560
      TabIndex        =   23
      Top             =   6840
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
      Caption0        =   "&Salir"
      Tag0            =   "9"
      ToolTipText0    =   "Salir"
      UserControlWidth=   1200
   End
End
Attribute VB_Name = "frmAsesores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrTipoAsesor() As String
Dim arrTipoIdentidad() As String
Dim strsql As String
Dim strEstado As String


Dim adoConsulta As ADODB.Recordset


Private Sub Form_Load()

    'Call InicializarValores
    Call CargarListas
    Call Buscar
    
'    Call ValidarPermisoUsoControl(Trim(gstrLoginUS), Me, Trim(App.Title) + Separador_Codigo_Objeto + _
'    gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)
    
    CentrarForm Me

    Set cmdAccion.FormularioActivo = Me
    Set cmdOpcion.FormularioActivo = Me
    
    
    
End Sub

Private Sub CargarListas()

    Dim strsql As String
    
    strsql = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPASE' ORDER BY CodParametro"
    CargarControlLista strsql, cboTipoAsesor, arrTipoAsesor(), ""
    If cboTipoAsesor.ListCount > 0 Then cboTipoAsesor.ListIndex = 0
    
    strsql = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPIDE' AND ValorParametro ='" & Codigo_Persona_Natural & "' ORDER BY DescripParametro"
    CargarControlLista strsql, cboTipoIdentidad, arrTipoIdentidad(), Sel_Defecto
    If cboTipoIdentidad.ListCount > 0 Then cboTipoIdentidad.ListIndex = 0

    
    
    
End Sub

Public Sub Buscar()

    Dim adoresultAux1 As ADODB.Recordset
    Set adoConsulta = New ADODB.Recordset
                                                                                    
    Me.MousePointer = vbHourglass
            
    strsql = "SELECT CodPersona,DescripPersona FROM InstitucionPersona WHERE TipoPersona='" & Codigo_Tipo_Persona_Contratante & "' ORDER BY DescripPersona"
        
    strEstado = Reg_Defecto
    With adoConsulta
        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open strsql
    End With
        
    tdgConsulta.DataSource = adoConsulta
    
    If adoConsulta.RecordCount > 0 Then strEstado = Reg_Consulta
    
    Me.MousePointer = vbDefault
            
End Sub

