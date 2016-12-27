VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inversionistas"
   ClientHeight    =   8415
   ClientLeft      =   975
   ClientTop       =   915
   ClientWidth     =   11730
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00000002&
   Icon            =   "frmCliente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8415
   ScaleWidth      =   11730
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   9360
      TabIndex        =   93
      Top             =   7560
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
      Caption0        =   "&Salir"
      Tag0            =   "9"
      Visible0        =   0   'False
      ToolTipText0    =   "Salir"
      UserControlWidth=   1200
   End
   Begin TAMControls2.ucBotonEdicion2 cmdAccion 
      Height          =   735
      Left            =   6420
      TabIndex        =   92
      Top             =   7560
      Visible         =   0   'False
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   1296
      Buttons         =   2
      Caption0        =   "&Guardar"
      Tag0            =   "2"
      Visible0        =   0   'False
      ToolTipText0    =   "Guardar"
      Caption1        =   "&Cancelar"
      Tag1            =   "8"
      Visible1        =   0   'False
      ToolTipText1    =   "Cancelar"
      UserControlWidth=   2700
   End
   Begin TAMControls2.ucBotonEdicion2 cmdOpcion 
      Height          =   735
      Left            =   480
      TabIndex        =   91
      Top             =   7560
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   1296
      Buttons         =   4
      Caption0        =   "&Nuevo"
      Tag0            =   "0"
      Visible0        =   0   'False
      ToolTipText0    =   "Nuevo"
      Caption1        =   "&Modificar"
      Tag1            =   "3"
      Visible1        =   0   'False
      ToolTipText1    =   "Modificar"
      Caption2        =   "&Eliminar"
      Tag2            =   "4"
      Visible2        =   0   'False
      ToolTipText2    =   "Eliminar"
      Caption3        =   "&Buscar"
      Tag3            =   "5"
      Visible3        =   0   'False
      ToolTipText3    =   "Buscar"
      UserControlWidth=   5700
   End
   Begin MSAdodcLib.Adodc adoDependientes 
      Height          =   330
      Left            =   8430
      Top             =   540
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
      Caption         =   "adoDependientes"
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
   Begin TabDlg.SSTab tabCliente 
      Height          =   7425
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   13097
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      Tab             =   6
      TabsPerRow      =   7
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
      TabPicture(0)   =   "frmCliente.frx":1CFA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "tdgConsulta"
      Tab(0).Control(1)=   "fraCliente(0)"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Datos Generales"
      TabPicture(1)   =   "frmCliente.frx":1D16
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraCliente(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Datos Domiciliarios"
      TabPicture(2)   =   "frmCliente.frx":1D32
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraCliente(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Datos Laborales y Bancarios"
      TabPicture(3)   =   "frmCliente.frx":1D4E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraCliente(3)"
      Tab(3).Control(1)=   "Frame2"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Datos Representantes"
      TabPicture(4)   =   "frmCliente.frx":1D6A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fraCliente(4)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Datos Vinculados"
      TabPicture(5)   =   "frmCliente.frx":1D86
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame1"
      Tab(5).Control(1)=   "lblDescrip(33)"
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "PEPs"
      TabPicture(6)   =   "frmCliente.frx":1DA2
      Tab(6).ControlEnabled=   -1  'True
      Tab(6).Control(0)=   "fraPep"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).ControlCount=   1
      Begin VB.Frame fraPep 
         Caption         =   "PEPs"
         Height          =   6015
         Left            =   210
         TabIndex        =   163
         Top             =   600
         Width           =   10695
         Begin VB.CommandButton btnPEPEliminar 
            Caption         =   "&Eliminar"
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
            Left            =   9120
            TabIndex        =   90
            Top             =   5160
            Width           =   1395
         End
         Begin VB.CommandButton btnPEPEditar 
            Caption         =   "&Modificar"
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
            Left            =   9120
            TabIndex        =   89
            Top             =   4440
            Width           =   1395
         End
         Begin VB.CommandButton btnPEPAgregar 
            Caption         =   "&Agregar"
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
            Left            =   9120
            TabIndex        =   87
            Top             =   3720
            Width           =   1395
         End
         Begin VB.CheckBox chkAdminRecPubNo 
            Alignment       =   1  'Right Justify
            Caption         =   "No"
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
            Height          =   255
            Left            =   9480
            TabIndex        =   84
            Top             =   1920
            Width           =   615
         End
         Begin VB.CheckBox chkAdminRecPubSi 
            Alignment       =   1  'Right Justify
            Caption         =   "Si"
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
            Left            =   8520
            TabIndex        =   83
            Top             =   1935
            Width           =   615
         End
         Begin VB.TextBox txtCargoDesemPEPS 
            Height          =   300
            Left            =   5160
            TabIndex        =   82
            Top             =   1920
            Width           =   2775
         End
         Begin VB.TextBox txtInstitucionPEPS 
            Height          =   300
            Left            =   600
            TabIndex        =   81
            Top             =   1920
            Width           =   4500
         End
         Begin VB.CheckBox chkPepNo 
            Alignment       =   1  'Right Justify
            Caption         =   "No"
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
            Height          =   375
            Left            =   9600
            TabIndex        =   80
            Top             =   480
            Width           =   615
         End
         Begin VB.CheckBox chkPepSi 
            Alignment       =   1  'Right Justify
            Caption         =   "Si"
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
            Height          =   375
            Left            =   8520
            TabIndex        =   79
            Top             =   480
            Width           =   615
         End
         Begin MSComCtl2.DTPicker dtpPepFechaDesde 
            Height          =   285
            Left            =   6240
            TabIndex        =   85
            Top             =   2880
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   503
            _Version        =   393216
            Format          =   175439873
            CurrentDate     =   38069
         End
         Begin MSComCtl2.DTPicker dtpPepFechaHasta 
            Height          =   285
            Left            =   8520
            TabIndex        =   86
            Top             =   2880
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   503
            _Version        =   393216
            Format          =   175439873
            CurrentDate     =   38069
         End
         Begin TrueOleDBGrid60.TDBGrid TDBPEP 
            Height          =   1815
            Left            =   600
            OleObjectBlob   =   "frmCliente.frx":1DBE
            TabIndex        =   88
            Top             =   3720
            Width           =   8295
         End
         Begin VB.Label lblDescInstitucion 
            Caption         =   $"frmCliente.frx":5E39
            ForeColor       =   &H00800000&
            Height          =   1095
            Left            =   600
            TabIndex        =   172
            Top             =   2400
            Width           =   4485
         End
         Begin VB.Label lblPep 
            Caption         =   $"frmCliente.frx":5F40
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
            Height          =   615
            Left            =   600
            TabIndex        =   171
            Top             =   480
            Width           =   7335
         End
         Begin VB.Label lblPepDesde 
            Caption         =   "Desde"
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
            Height          =   285
            Index           =   57
            Left            =   5640
            TabIndex        =   170
            Top             =   2880
            Width           =   855
         End
         Begin VB.Label lblPepHasta 
            Caption         =   "Hasta"
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
            Height          =   285
            Index           =   56
            Left            =   7920
            TabIndex        =   169
            Top             =   2880
            Width           =   855
         End
         Begin VB.Label lblFechasDesem 
            Caption         =   "Fechas en las que desempeño o desempeña el cargo"
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
            Height          =   255
            Left            =   5640
            TabIndex        =   168
            Top             =   2520
            Width           =   4935
         End
         Begin VB.Label lblAdministra 
            Caption         =   "Administra Recursos Públicos"
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
            Height          =   255
            Left            =   8040
            TabIndex        =   167
            Top             =   1560
            Width           =   2535
         End
         Begin VB.Label lblCargo 
            Caption         =   "Cargo que desempeña"
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
            Height          =   375
            Left            =   5160
            TabIndex        =   166
            Top             =   1560
            Width           =   2175
         End
         Begin VB.Label lblInstitucion 
            Caption         =   "Institución"
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
            Height          =   255
            Left            =   600
            TabIndex        =   165
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label lblInfo 
            Caption         =   "Si su respuesta es SI será necesario completar la informacion requerida lineas abajo."
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   56
            Left            =   600
            TabIndex        =   164
            Top             =   1080
            Width           =   6495
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Cuentas Bancarias"
         Height          =   2655
         Left            =   -74790
         TabIndex        =   155
         Top             =   4590
         Width           =   10680
         Begin VB.TextBox txtCCI 
            Height          =   315
            Left            =   6000
            MaxLength       =   20
            TabIndex        =   56
            Top             =   540
            Width           =   2235
         End
         Begin VB.CommandButton btnAgregarCtaCte 
            Caption         =   "&Agregar"
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
            Left            =   9060
            TabIndex        =   58
            Top             =   960
            Width           =   1395
         End
         Begin VB.CommandButton btnEditarCtaCte 
            Caption         =   "&Modificar"
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
            Left            =   9060
            TabIndex        =   60
            Top             =   1530
            Width           =   1395
         End
         Begin VB.CommandButton btnEliminarCtaCte 
            Caption         =   "&Eliminar"
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
            Left            =   9060
            TabIndex        =   61
            Top             =   2100
            Width           =   1395
         End
         Begin VB.TextBox txtCtaCte 
            Height          =   315
            Left            =   8310
            TabIndex        =   57
            Top             =   540
            Width           =   2115
         End
         Begin VB.ComboBox cboMoneda 
            Height          =   315
            ItemData        =   "frmCliente.frx":5FEA
            Left            =   4500
            List            =   "frmCliente.frx":5FEC
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   540
            Width           =   1485
         End
         Begin VB.ComboBox cboTipoCta 
            Height          =   315
            Left            =   2820
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   540
            Width           =   1665
         End
         Begin VB.ComboBox cboBancos 
            Height          =   315
            Left            =   210
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   540
            Width           =   2595
         End
         Begin TrueOleDBGrid60.TDBGrid TDBCtasCtes 
            Bindings        =   "frmCliente.frx":5FEE
            Height          =   1515
            Left            =   210
            OleObjectBlob   =   "frmCliente.frx":600C
            TabIndex        =   59
            Top             =   975
            Width           =   8775
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Banco"
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
            Height          =   225
            Index           =   37
            Left            =   210
            TabIndex        =   160
            Top             =   270
            Width           =   1065
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Tipo de Cuenta"
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
            Height          =   225
            Index           =   38
            Left            =   2820
            TabIndex        =   159
            Top             =   270
            Width           =   1575
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Moneda"
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
            Height          =   225
            Index           =   39
            Left            =   4530
            TabIndex        =   158
            Top             =   270
            Width           =   1125
         End
         Begin VB.Label lblDescrip 
            Caption         =   "# CCI"
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
            Height          =   225
            Index           =   40
            Left            =   6000
            TabIndex        =   157
            Top             =   270
            Width           =   825
         End
         Begin VB.Label lblDescrip 
            Caption         =   "# Cuenta"
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
            Height          =   225
            Index           =   41
            Left            =   8310
            TabIndex        =   156
            Top             =   270
            Width           =   825
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Familiares Dependientes (Esposa, Hijos)"
         Height          =   6615
         Left            =   -74760
         TabIndex        =   139
         Top             =   780
         Width           =   10725
         Begin VB.ComboBox cboVinculacion 
            Height          =   315
            Left            =   7860
            Style           =   2  'Dropdown List
            TabIndex        =   74
            Top             =   630
            Width           =   2520
         End
         Begin VB.CommandButton btnDepEliminar 
            Caption         =   "&Eliminar"
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
            Left            =   8970
            TabIndex        =   78
            Top             =   2460
            Width           =   1395
         End
         Begin VB.CommandButton btnDepEditar 
            Caption         =   "&Modificar"
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
            Left            =   8970
            TabIndex        =   77
            Top             =   1770
            Width           =   1395
         End
         Begin VB.CommandButton btnDepAgregar 
            Caption         =   "&Agregar"
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
            Left            =   8970
            TabIndex        =   75
            Top             =   1050
            Width           =   1395
         End
         Begin VB.TextBox txtDocumentoLaboral 
            Height          =   315
            Left            =   5790
            TabIndex        =   73
            Top             =   630
            Width           =   1995
         End
         Begin VB.ComboBox cboTipoDocLaboral 
            Height          =   315
            Left            =   3450
            Style           =   2  'Dropdown List
            TabIndex        =   72
            Top             =   630
            Width           =   2280
         End
         Begin VB.TextBox txtNombresLaboral 
            Height          =   315
            Left            =   120
            TabIndex        =   71
            Top             =   630
            Width           =   3195
         End
         Begin TrueOleDBGrid60.TDBGrid TDBDependientes 
            Bindings        =   "frmCliente.frx":9D4C
            Height          =   1785
            Left            =   120
            OleObjectBlob   =   "frmCliente.frx":9D6A
            TabIndex        =   76
            Top             =   1050
            Width           =   8775
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Tipo de Vinculación"
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
            Height          =   225
            Index           =   49
            Left            =   7860
            TabIndex        =   147
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Nro Documento"
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
            Height          =   225
            Index           =   48
            Left            =   5790
            TabIndex        =   146
            Top             =   360
            Width           =   1515
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Tipo de Documento"
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
            Height          =   225
            Index           =   47
            Left            =   3450
            TabIndex        =   145
            Top             =   360
            Width           =   1875
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Nombres y Apellidos"
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
            Height          =   225
            Index           =   46
            Left            =   120
            TabIndex        =   144
            Top             =   360
            Width           =   3105
         End
      End
      Begin VB.Frame fraCliente 
         Caption         =   "Centro de Trabajo"
         Height          =   3975
         Index           =   3
         Left            =   -74790
         TabIndex        =   124
         Top             =   600
         Width           =   10680
         Begin VB.TextBox txtOcupCompletar 
            Height          =   315
            Left            =   6540
            MaxLength       =   25
            TabIndex        =   41
            Top             =   1110
            Width           =   3840
         End
         Begin VB.ComboBox cboOcupacion 
            Height          =   315
            Left            =   1700
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   1100
            Width           =   3180
         End
         Begin VB.TextBox txtCelularTrabajo 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Left            =   1700
            MaxLength       =   15
            TabIndex        =   52
            Top             =   3520
            Visible         =   0   'False
            Width           =   3960
         End
         Begin VB.TextBox txtEMailTrabajo 
            Height          =   315
            Left            =   7000
            MaxLength       =   45
            TabIndex        =   49
            Top             =   2470
            Width           =   3360
         End
         Begin VB.TextBox txtTelefonoTrabajo 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Left            =   7000
            MaxLength       =   15
            TabIndex        =   50
            Top             =   2820
            Width           =   3360
         End
         Begin VB.TextBox txtFaxTrabajo 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Left            =   7000
            MaxLength       =   15
            TabIndex        =   51
            Top             =   3170
            Width           =   3360
         End
         Begin VB.ComboBox cboDepartamentoTrabajo 
            Height          =   315
            Left            =   1700
            Style           =   2  'Dropdown List
            TabIndex        =   45
            Top             =   2470
            Width           =   3960
         End
         Begin VB.ComboBox cboProvinciaTrabajo 
            Height          =   315
            Left            =   1700
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Top             =   2820
            Width           =   3960
         End
         Begin VB.ComboBox cboPaisTrabajo 
            Height          =   315
            ItemData        =   "frmCliente.frx":D4D2
            Left            =   1700
            List            =   "frmCliente.frx":D4D9
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   2120
            Width           =   3960
         End
         Begin VB.ComboBox cboDistritoTrabajo 
            Height          =   315
            Left            =   1700
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   3170
            Width           =   3960
         End
         Begin VB.TextBox txtNombreEmpresa 
            Height          =   315
            Left            =   1700
            MaxLength       =   25
            TabIndex        =   38
            Top             =   750
            Width           =   3180
         End
         Begin VB.TextBox txtCargo 
            Height          =   315
            Left            =   7000
            MaxLength       =   25
            TabIndex        =   48
            Top             =   2120
            Width           =   3360
         End
         Begin VB.TextBox txtDireccionTrabajo1 
            Height          =   315
            Left            =   1700
            MaxLength       =   45
            TabIndex        =   42
            Top             =   1450
            Width           =   7035
         End
         Begin VB.TextBox txtDireccionTrabajo2 
            Height          =   315
            Left            =   1700
            MaxLength       =   45
            TabIndex        =   43
            Top             =   1770
            Width           =   7035
         End
         Begin VB.OptionButton optTipoTrabajador 
            Caption         =   "Dependiente"
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
            Height          =   345
            Index           =   0
            Left            =   400
            TabIndex        =   36
            Top             =   390
            Width           =   1545
         End
         Begin VB.OptionButton optTipoTrabajador 
            Caption         =   "Independiente"
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
            Height          =   345
            Index           =   1
            Left            =   2130
            TabIndex        =   37
            Top             =   390
            Width           =   1695
         End
         Begin VB.TextBox txtRUCEmpresa 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Left            =   6540
            MaxLength       =   11
            TabIndex        =   39
            Top             =   750
            Width           =   2160
         End
         Begin VB.Label lblCompletar 
            Caption         =   "Complementar"
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
            Height          =   225
            Index           =   51
            Left            =   5220
            TabIndex        =   149
            Top             =   1160
            Width           =   1215
         End
         Begin VB.Label lblcelulartrabajo 
            Caption         =   "Celular"
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
            Height          =   225
            Index           =   33
            Left            =   400
            TabIndex        =   137
            Top             =   3580
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.Label lblDescrip 
            Caption         =   "E-Mail"
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
            Height          =   225
            Index           =   28
            Left            =   5800
            TabIndex        =   136
            Top             =   2530
            Width           =   675
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Teléfono"
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
            Height          =   225
            Index           =   27
            Left            =   5800
            TabIndex        =   135
            Top             =   2880
            Width           =   915
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Fax"
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
            Height          =   225
            Index           =   26
            Left            =   5800
            TabIndex        =   134
            Top             =   3230
            Width           =   795
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Departamento"
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
            Height          =   225
            Index           =   25
            Left            =   400
            TabIndex        =   133
            Top             =   2530
            Width           =   1365
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Provincia"
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
            Height          =   225
            Index           =   24
            Left            =   400
            TabIndex        =   132
            Top             =   2880
            Width           =   1035
         End
         Begin VB.Label lblDescrip 
            Caption         =   "País"
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
            Height          =   225
            Index           =   23
            Left            =   400
            TabIndex        =   131
            Top             =   2180
            Width           =   750
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Dirección"
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
            Height          =   225
            Index           =   22
            Left            =   400
            TabIndex        =   130
            Top             =   1510
            Width           =   1215
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Distrito"
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
            Height          =   240
            Index           =   21
            Left            =   400
            TabIndex        =   129
            Top             =   3230
            Width           =   870
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Empresa"
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
            Height          =   225
            Index           =   20
            Left            =   400
            TabIndex        =   128
            Top             =   810
            Width           =   1065
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Ocupación"
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
            Height          =   225
            Index           =   19
            Left            =   400
            TabIndex        =   127
            Top             =   1160
            Width           =   960
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Cargo"
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
            Index           =   18
            Left            =   5800
            TabIndex        =   126
            Top             =   2180
            Width           =   645
         End
         Begin VB.Label lblDescrip 
            Caption         =   "RUC"
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
            Height          =   225
            Index           =   32
            Left            =   5220
            TabIndex        =   125
            Top             =   810
            Width           =   705
         End
      End
      Begin VB.Frame fraCliente 
         Caption         =   "Representantes Legales"
         Height          =   6615
         Index           =   4
         Left            =   -74790
         TabIndex        =   123
         Top             =   600
         Width           =   10695
         Begin VB.ComboBox cboVinculacionLegales 
            Height          =   315
            Left            =   7890
            Style           =   2  'Dropdown List
            TabIndex        =   66
            Top             =   600
            Width           =   2520
         End
         Begin VB.TextBox txtApellidosLegales 
            Height          =   315
            Left            =   1950
            TabIndex        =   63
            Top             =   600
            Width           =   1815
         End
         Begin MSAdodcLib.Adodc adoRepresentantes 
            Height          =   330
            Left            =   8070
            Top             =   4050
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
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
         Begin VB.TextBox txtNombresLegales 
            Height          =   315
            Left            =   150
            TabIndex        =   62
            Top             =   600
            Width           =   1755
         End
         Begin VB.ComboBox cboTipoDocLegales 
            Height          =   315
            Left            =   3810
            Style           =   2  'Dropdown List
            TabIndex        =   64
            Top             =   600
            Width           =   2130
         End
         Begin VB.TextBox txtDocumentoLegales 
            Height          =   315
            Left            =   6030
            TabIndex        =   65
            Top             =   600
            Width           =   1785
         End
         Begin VB.CommandButton btnRepAgregar 
            Caption         =   "&Agregar"
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
            Left            =   9060
            TabIndex        =   67
            Top             =   1140
            Width           =   1395
         End
         Begin VB.CommandButton btnRepEditar 
            Caption         =   "&Modificar"
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
            Left            =   9060
            TabIndex        =   69
            Top             =   1920
            Width           =   1395
         End
         Begin VB.CommandButton btnRepEliminar 
            Caption         =   "&Eliminar"
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
            Left            =   9060
            TabIndex        =   70
            Top             =   2700
            Width           =   1395
         End
         Begin TrueOleDBGrid60.TDBGrid TDBRepresentantes 
            Bindings        =   "frmCliente.frx":D4ED
            Height          =   1935
            Left            =   150
            OleObjectBlob   =   "frmCliente.frx":D50B
            TabIndex        =   68
            Top             =   1170
            Width           =   8775
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Apellidos"
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
            Height          =   225
            Index           =   50
            Left            =   1980
            TabIndex        =   148
            Top             =   360
            Width           =   1635
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Tipo de Vinculación"
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
            Height          =   225
            Index           =   45
            Left            =   7890
            TabIndex        =   143
            Top             =   360
            Width           =   2025
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Nro Documento"
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
            Height          =   225
            Index           =   44
            Left            =   6030
            TabIndex        =   142
            Top             =   360
            Width           =   1515
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Tipo de Documento"
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
            Height          =   225
            Index           =   43
            Left            =   3810
            TabIndex        =   141
            Top             =   360
            Width           =   1875
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Nombres"
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
            Height          =   225
            Index           =   42
            Left            =   150
            TabIndex        =   140
            Top             =   360
            Width           =   1635
         End
      End
      Begin VB.Frame fraCliente 
         Caption         =   "Domicilio"
         Height          =   6255
         Index           =   2
         Left            =   -74790
         TabIndex        =   107
         Top             =   600
         Width           =   10680
         Begin VB.TextBox txtFaxDomicilio 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Left            =   2400
            MaxLength       =   15
            TabIndex        =   29
            Top             =   4410
            Width           =   3360
         End
         Begin VB.TextBox txtCelularDomicilio 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Left            =   2400
            MaxLength       =   15
            TabIndex        =   28
            Top             =   4410
            Width           =   3360
         End
         Begin VB.ComboBox cboEnvioInformacion 
            Height          =   315
            Left            =   2400
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   3060
            Width           =   4000
         End
         Begin VB.TextBox txtWeb 
            Height          =   315
            Left            =   2400
            MaxLength       =   150
            TabIndex        =   30
            Top             =   4860
            Width           =   3360
         End
         Begin VB.TextBox txtDireccionDomicilio1 
            Height          =   315
            Left            =   2400
            MaxLength       =   45
            TabIndex        =   23
            Top             =   510
            Width           =   6500
         End
         Begin VB.TextBox txtDireccionDomicilio2 
            Height          =   315
            Left            =   2400
            MaxLength       =   45
            TabIndex        =   24
            Top             =   810
            Width           =   6500
         End
         Begin VB.CommandButton cmdDefault 
            Caption         =   "&Alternativo"
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
            Left            =   9120
            TabIndex        =   25
            ToolTipText     =   "Asignar la dirección de la Administradora"
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtTelefonoDomicilio 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Left            =   2400
            MaxLength       =   15
            TabIndex        =   27
            Top             =   3960
            Width           =   3360
         End
         Begin VB.TextBox txtEMailDomicilio 
            Height          =   315
            Left            =   2400
            MaxLength       =   45
            TabIndex        =   26
            Top             =   3510
            Width           =   3360
         End
         Begin VB.ComboBox cboPais 
            Height          =   315
            ItemData        =   "frmCliente.frx":11871
            Left            =   2400
            List            =   "frmCliente.frx":11878
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   1260
            Width           =   4000
         End
         Begin VB.ComboBox cboProvincia 
            Height          =   315
            Left            =   2400
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   2160
            Width           =   4000
         End
         Begin VB.ComboBox cboDepartamento 
            Height          =   315
            Left            =   2400
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   1710
            Width           =   4000
         End
         Begin VB.ComboBox cboDistrito 
            Height          =   315
            Left            =   2400
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   2610
            Width           =   4000
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Fax"
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
            Height          =   225
            Index           =   12
            Left            =   500
            TabIndex        =   162
            Top             =   4470
            Width           =   705
         End
         Begin VB.Label lblcomentariodev 
            Caption         =   "Aqui debajo hay otro campo, 1 es txtFaxDomicilio y el otro txtCelularDomicilio"
            Height          =   375
            Left            =   5970
            TabIndex        =   161
            Top             =   4380
            Visible         =   0   'False
            Width           =   3195
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Envío de Información"
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
            Height          =   285
            Index           =   55
            Left            =   500
            TabIndex        =   154
            Top             =   3120
            Width           =   2040
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Web"
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
            Height          =   225
            Index           =   52
            Left            =   500
            TabIndex        =   151
            Top             =   4920
            Width           =   705
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Teléfono"
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
            Height          =   225
            Index           =   11
            Left            =   500
            TabIndex        =   114
            Top             =   4020
            Width           =   795
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Email"
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
            Height          =   225
            Index           =   10
            Left            =   500
            TabIndex        =   113
            Top             =   3570
            Width           =   705
         End
         Begin VB.Label lblDescrip 
            Caption         =   "País"
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
            Height          =   225
            Index           =   17
            Left            =   500
            TabIndex        =   112
            Top             =   1320
            Width           =   750
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Provincia"
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
            Height          =   225
            Index           =   16
            Left            =   500
            TabIndex        =   111
            Top             =   2220
            Width           =   1035
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Departamento"
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
            Height          =   225
            Index           =   15
            Left            =   500
            TabIndex        =   110
            Top             =   1770
            Width           =   1245
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Distrito"
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
            Height          =   225
            Index           =   14
            Left            =   500
            TabIndex        =   109
            Top             =   2670
            Width           =   870
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Dirección"
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
            Height          =   220
            Index           =   13
            Left            =   500
            TabIndex        =   108
            Top             =   570
            Width           =   1065
         End
      End
      Begin VB.Frame fraCliente 
         Caption         =   "Definición"
         Height          =   6255
         Index           =   1
         Left            =   -74790
         TabIndex        =   95
         Top             =   600
         Width           =   10680
         Begin VB.ComboBox cboPaisResidencia 
            Height          =   315
            ItemData        =   "frmCliente.frx":11885
            Left            =   7900
            List            =   "frmCliente.frx":1188C
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   4110
            Width           =   2400
         End
         Begin VB.ComboBox cboPaisNacimiento 
            Height          =   315
            ItemData        =   "frmCliente.frx":11899
            Left            =   7900
            List            =   "frmCliente.frx":118A0
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   2310
            Width           =   2400
         End
         Begin VB.TextBox txtOtroDocumento 
            Height          =   315
            Left            =   2250
            MaxLength       =   50
            TabIndex        =   9
            Top             =   1860
            Width           =   3360
         End
         Begin VB.TextBox txtPartidaElectronica 
            Height          =   315
            Left            =   7900
            MaxLength       =   75
            TabIndex        =   22
            Top             =   5010
            Width           =   2370
         End
         Begin VB.TextBox txtObjetivoSocial 
            Height          =   315
            Left            =   2250
            MaxLength       =   75
            TabIndex        =   15
            Top             =   4560
            Width           =   3360
         End
         Begin VB.ComboBox cboSexo 
            Height          =   315
            Left            =   7900
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   2760
            Width           =   2385
         End
         Begin VB.ComboBox cboClasePersona 
            Height          =   315
            Left            =   2250
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   510
            Width           =   3360
         End
         Begin VB.TextBox txtNumIdentidad 
            Height          =   315
            Left            =   2250
            MaxLength       =   15
            TabIndex        =   10
            Top             =   2310
            Width           =   3360
         End
         Begin VB.ComboBox cboTipoDocumento 
            Height          =   315
            Left            =   2250
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1410
            Width           =   5940
         End
         Begin VB.TextBox txtNombres 
            Height          =   315
            Left            =   2250
            MaxLength       =   25
            TabIndex        =   13
            Top             =   3660
            Width           =   3360
         End
         Begin VB.TextBox txtApellidoMaterno 
            Height          =   315
            Left            =   2250
            MaxLength       =   25
            TabIndex        =   12
            Top             =   3210
            Width           =   3360
         End
         Begin VB.ComboBox cboNacionalidad 
            Height          =   315
            Left            =   7920
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   3660
            Width           =   2385
         End
         Begin VB.ComboBox cboEstadoCivil 
            Height          =   315
            Left            =   7900
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   3210
            Width           =   2385
         End
         Begin VB.TextBox txtRazonSocial 
            Height          =   315
            Left            =   2250
            MaxLength       =   75
            TabIndex        =   14
            Top             =   4110
            Width           =   3360
         End
         Begin VB.TextBox txtApellidoPaterno 
            Height          =   315
            Left            =   2250
            MaxLength       =   25
            TabIndex        =   11
            Top             =   2760
            Width           =   3360
         End
         Begin MSComCtl2.DTPicker dtpFechaNacimiento 
            Height          =   315
            Left            =   7900
            TabIndex        =   8
            Top             =   1860
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   556
            _Version        =   393216
            Format          =   175439873
            CurrentDate     =   38069
         End
         Begin MSComCtl2.DTPicker dtpFechaConstitucion 
            Height          =   315
            Left            =   7900
            TabIndex        =   21
            Top             =   4560
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   556
            _Version        =   393216
            Format          =   175439873
            CurrentDate     =   38069
         End
         Begin VB.Label lblDescrip 
            Caption         =   "País de Residencia"
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
            Height          =   225
            Index           =   54
            Left            =   5895
            TabIndex        =   153
            Top             =   4170
            Width           =   1860
         End
         Begin VB.Label lblDescrip 
            Caption         =   "País de Nacimiento"
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
            Height          =   225
            Index           =   53
            Left            =   5895
            TabIndex        =   152
            Top             =   2370
            Width           =   1860
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Nro Documento"
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
            Height          =   225
            Index           =   51
            Left            =   500
            TabIndex        =   150
            Top             =   2370
            Width           =   1740
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Objetivo Social"
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
            Height          =   225
            Index           =   36
            Left            =   500
            TabIndex        =   122
            Top             =   4620
            Width           =   1470
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Partida Electrónica"
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
            Height          =   225
            Index           =   35
            Left            =   5895
            TabIndex        =   121
            Top             =   5070
            Width           =   1830
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Fecha de Constitución"
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
            Height          =   225
            Index           =   34
            Left            =   5895
            TabIndex        =   120
            Top             =   4620
            Width           =   1860
         End
         Begin VB.Label lblCodigoCliente 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   7905
            TabIndex        =   118
            Top             =   960
            Width           =   2385
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Sexo"
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
            Height          =   225
            Index           =   31
            Left            =   5895
            TabIndex        =   117
            Top             =   2820
            Width           =   930
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Razón Social"
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
            Height          =   225
            Index           =   9
            Left            =   500
            TabIndex        =   116
            Top             =   4170
            Width           =   1290
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Clase Persona"
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
            Height          =   225
            Index           =   30
            Left            =   500
            TabIndex        =   115
            Top             =   570
            Width           =   1500
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Código Cliente"
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
            Height          =   225
            Index           =   8
            Left            =   5895
            TabIndex        =   106
            Top             =   1020
            Width           =   1500
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Nombres"
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
            Height          =   225
            Index           =   5
            Left            =   500
            TabIndex        =   105
            Top             =   3720
            Width           =   930
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Apellido Materno"
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
            Height          =   225
            Index           =   4
            Left            =   500
            TabIndex        =   104
            Top             =   3270
            Width           =   1710
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Apellido Paterno"
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
            Height          =   225
            Index           =   3
            Left            =   500
            TabIndex        =   103
            Top             =   2820
            Width           =   1545
         End
         Begin VB.Label lblFechaIngreso 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "dd/mm/yyyy"
            Height          =   300
            Left            =   7900
            TabIndex        =   102
            Top             =   510
            Width           =   2385
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Fecha de Ingreso"
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
            Height          =   225
            Index           =   2
            Left            =   5900
            TabIndex        =   101
            Top             =   570
            Width           =   1500
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Otro Documento"
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
            Height          =   225
            Index           =   1
            Left            =   500
            TabIndex        =   100
            Top             =   1920
            Width           =   1740
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Tipo Documento"
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
            Height          =   225
            Index           =   0
            Left            =   500
            TabIndex        =   99
            Top             =   1470
            Width           =   1500
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Nacionalidad"
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
            Height          =   225
            Index           =   7
            Left            =   5895
            TabIndex        =   98
            Top             =   3720
            Width           =   1500
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Estado Civil"
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
            Height          =   225
            Index           =   6
            Left            =   5895
            TabIndex        =   97
            Top             =   3270
            Width           =   1500
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Fecha de Nacimiento"
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
            Height          =   225
            Index           =   29
            Left            =   5895
            TabIndex        =   96
            Top             =   1920
            Width           =   1860
         End
      End
      Begin VB.Frame fraCliente 
         Caption         =   "Criterios de búsqueda"
         ForeColor       =   &H00000000&
         Height          =   1200
         Index           =   0
         Left            =   -74790
         TabIndex        =   94
         Top             =   600
         Width           =   10680
         Begin VB.OptionButton optCriterios 
            Caption         =   "Num. Documento"
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
            Height          =   225
            Index           =   0
            Left            =   6240
            TabIndex        =   1
            Top             =   570
            Width           =   1830
         End
         Begin VB.OptionButton optCriterios 
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
            Height          =   225
            Index           =   1
            Left            =   405
            TabIndex        =   3
            Top             =   570
            Value           =   -1  'True
            Width           =   1425
         End
         Begin VB.TextBox txtNumDocumento 
            Enabled         =   0   'False
            Height          =   285
            Left            =   8280
            MaxLength       =   15
            TabIndex        =   2
            Top             =   540
            Width           =   1935
         End
         Begin VB.TextBox txtDescripCliente 
            Height          =   285
            Left            =   2160
            MaxLength       =   75
            TabIndex        =   4
            Top             =   540
            Width           =   3495
         End
         Begin VB.Label lblContador 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   6750
            TabIndex        =   119
            Top             =   900
            Width           =   3495
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmCliente.frx":118AD
         Height          =   4845
         Left            =   -74790
         OleObjectBlob   =   "frmCliente.frx":118C7
         TabIndex        =   5
         Top             =   2100
         Width           =   10680
      End
      Begin VB.Label lblDescrip 
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
         Height          =   225
         Index           =   33
         Left            =   -74550
         TabIndex        =   138
         Top             =   900
         Width           =   4815
      End
   End
End
Attribute VB_Name = "frmCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrClasePersona()           As String, arrTipoDocumento()           As String
Dim arrSexo()                   As String, arrEstadoCivil()             As String
Dim arrNacionalidad()           As String, arrPais()                    As String
Dim arrDepartamento()           As String, arrProvincia()               As String
Dim arrDistrito()               As String, arrPaisTrabajo()             As String
Dim arrDepartamentoTrabajo()    As String, arrProvinciaTrabajo()        As String
Dim arrDistritoTrabajo()        As String, arrVinculacion()             As String
Dim arrVinculacionLegales()     As String
Dim arrTipoDocRepresentante()   As String, arrTipoDocVinculado()        As String

Dim strCodClasePersona          As String, strCodTipoDocumento          As String
Dim strCodSexo                  As String, strCodEstadoCivil            As String
Dim strCodNacionalidad          As String, strCodPais                   As String
Dim strCodDepartamento          As String, strCodProvincia              As String
Dim strCodDistrito              As String, strCodPaisTrabajo            As String
Dim strCodDepartamentoTrabajo   As String, strCodProvinciaTrabajo       As String
Dim strCodDistritoTrabajo       As String, strCodCliente                As String
Dim strCodMoneda                As String, strCodTipoTrab               As String
Dim arrOcupacion()              As String, arrMoneda()                  As String
Dim arrTipoCta()                As String, arrBancos()                  As String
Dim strEstado                   As String, strCodOcupacion              As String
Dim strCodPaisNacimiento        As String, strCodPaisResidencia         As String
Dim arrPaisNacimiento()         As String, arrPaisResidencia()          As String
Dim arrEnvioInformacion()       As String, strCodEnvioInformacion       As String
Dim strPEPS                     As String, strInstitucionPEPS           As String
Dim strCargoPEPS                As String, strRecursosPublicos          As String
Dim intCodIndti As Integer
Dim adoRegistroAux  As ADODB.Recordset, adoRegistroAuxCtaCte, adoRegistroPEP, adoRegistroAuxPEP
Dim adoConsulta                 As ADODB.Recordset
Dim indSortAsc                  As Boolean, indSortDesc                 As Boolean

Public Sub Adicionar()
            
    frmMainMdi.stbMdi.Panels(3).Text = "Adicionar inversionista..."
    
    intCodIndti = Valor_Numero
    
    strEstado = Reg_Adicion
    LlenarFormulario strEstado
    cmdOpcion.Visible = False
    cmdAccion.Visible = True
    With tabCliente
        .TabEnabled(0) = True
        .TabEnabled(1) = True
        .TabEnabled(2) = True
        .TabEnabled(3) = True
        .TabEnabled(4) = True
        .TabEnabled(5) = True
        .TabEnabled(6) = True
        .Tab = 1
    End With
    
    chkPepNo.Value = 1
    
    
    'Call Deshabilita
            
End Sub

Private Sub LlenarFormulario(strModo As String)


    Dim strSql As String
    Dim intRegistro As Integer
    
    Select Case strModo
        Case Reg_Adicion
        
            cboClasePersona.Enabled = True ' HMC

            cboClasePersona.ListIndex = -1
            intRegistro = ObtenerItemLista(arrClasePersona(), Codigo_Persona_Natural)
            If intRegistro >= Valor_Numero Then cboClasePersona.ListIndex = intRegistro
            
            cboTipoDocumento.ListIndex = -1
            If cboTipoDocumento.ListCount > 0 Then cboTipoDocumento.ListIndex = Valor_Numero
              
            txtNumIdentidad.Text = Valor_Caracter
            txtApellidoPaterno.Text = Valor_Caracter
            txtApellidoMaterno.Text = Valor_Caracter
            txtNombres.Text = Valor_Caracter
            txtRazonSocial.Text = Valor_Caracter
            txtWeb.Text = Valor_Caracter
            
            
            lblFechaIngreso.Caption = CStr(gdatFechaActual)
            'lblCodigoCliente.Caption = Valor_Caracter
            lblCodigoCliente.Caption = Format(strCodTipoDocumento & Trim(txtNumIdentidad.Text) & strCodClasePersona, "00000000000000000000")
            dtpFechaNacimiento.Value = gdatFechaActual
            
            cboSexo.ListIndex = -1
            If cboSexo.ListCount > 0 Then cboSexo.ListIndex = Valor_Numero
            
            cboEstadoCivil.ListIndex = -1
            If cboEstadoCivil.ListCount > 0 Then cboEstadoCivil.ListIndex = Valor_Numero
            
            cboNacionalidad.ListIndex = -1
            If cboNacionalidad.ListCount > 0 Then cboNacionalidad.ListIndex = Valor_Numero
            
            cboPaisNacimiento.ListIndex = -1
            If cboPaisNacimiento.ListCount > 0 Then cboPaisNacimiento.ListIndex = Valor_Numero
            
            cboPaisResidencia.ListIndex = -1
            If cboPaisResidencia.ListCount > 0 Then cboPaisResidencia.ListIndex = Valor_Numero
            
            txtDireccionDomicilio1.Text = Valor_Caracter
            txtDireccionDomicilio2.Text = Valor_Caracter
            txtEMailDomicilio.Text = Valor_Caracter
            txtTelefonoDomicilio.Text = Valor_Caracter
            txtFaxDomicilio.Text = Valor_Caracter
            txtCelularDomicilio.Text = Valor_Caracter
            
            cboPais.ListIndex = -1
            If cboPais.ListCount > 0 Then cboPais.ListIndex = Valor_Numero
            
            txtNombreEmpresa.Text = Valor_Caracter
            'txtOcupacion.Text = Valor_Caracter
            cboOcupacion.ListIndex = Valor_Numero
            txtCargo.Text = Valor_Caracter
            txtDireccionTrabajo1.Text = Valor_Caracter
            txtDireccionTrabajo2.Text = Valor_Caracter
            txtEMailTrabajo.Text = Valor_Caracter
            txtTelefonoTrabajo.Text = Valor_Caracter
            txtFaxTrabajo.Text = Valor_Caracter
            
            cboPaisTrabajo.ListIndex = -1
            If cboPaisTrabajo.ListCount > 0 Then cboPaisTrabajo.ListIndex = Valor_Numero
            
            cboEnvioInformacion.ListIndex = -1
            If cboEnvioInformacion.ListCount > 0 Then cboEnvioInformacion.ListIndex = Valor_Numero
                                    
            cboTipoDocumento.SetFocus
            
            Call ConfiguraRecordsetPep
            Call ConfiguraRecordsetAuxiliar
            Call ConfiguraRecordsetAuxiliarCtaCte
            TDBCtasCtes.DataSource = adoRegistroAuxCtaCte
            TDBPEP.DataSource = adoRegistroAuxPEP
            If strCodClasePersona = Codigo_Persona_Natural Then
                TDBDependientes.DataSource = adoRegistroAux
            End If
            If strCodClasePersona = Codigo_Persona_Juridica Then
                TDBRepresentantes.DataSource = adoRegistroAux
            End If
            
        Case Reg_Edicion
            Dim adoRegistro As ADODB.Recordset

            Set adoRegistro = New ADODB.Recordset

            strCodCliente = Trim(tdgConsulta.Columns(0))
            
            adoComm.CommandText = "{ call up_ACSelDatosParametro(7,'" & strCodCliente & "') }"
            Set adoRegistro = adoComm.Execute

            If Not adoRegistro.EOF Then
            
                cboClasePersona.Enabled = False ' HMC
            
                lblFechaIngreso.Caption = adoRegistro("FechaIngreso")
                lblCodigoCliente.Caption = adoRegistro("CodUnico")
                
                intRegistro = ObtenerItemLista(arrClasePersona(), adoRegistro("ClaseCliente"))
                If intRegistro >= Valor_Numero Then cboClasePersona.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrTipoDocumento(), adoRegistro("TipoIdentidad"))
                If intRegistro >= Valor_Numero Then cboTipoDocumento.ListIndex = intRegistro
                
                    intRegistro = ObtenerItemLista(arrPaisNacimiento(), adoRegistro("PaisNacimiento"))
                    If intRegistro >= Valor_Numero Then cboPaisNacimiento.ListIndex = intRegistro
              
                
                intRegistro = ObtenerItemLista(arrPaisResidencia(), adoRegistro("PaisResidencia"))
                If intRegistro >= Valor_Numero Then cboPaisResidencia.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrEnvioInformacion(), adoRegistro("EnvioInformacion"))
                If intRegistro >= Valor_Numero Then cboEnvioInformacion.ListIndex = intRegistro
                
                txtNumIdentidad.Text = Trim(adoRegistro("NumIdentidad"))
                txtApellidoPaterno.Text = Trim(adoRegistro("ApellidoPaterno"))
                txtApellidoMaterno.Text = Trim(adoRegistro("ApellidoMaterno"))
                txtNombres.Text = Trim(adoRegistro("Nombres"))
                txtRazonSocial.Text = Trim(adoRegistro("RazonSocial"))
                
                dtpFechaNacimiento.Value = adoRegistro("FechaNacimiento")
                dtpFechaConstitucion.Value = adoRegistro("FechaConstitucion")
                txtObjetivoSocial.Text = Trim(adoRegistro("ObjetivoSocial"))
                txtPartidaElectronica.Text = Trim(adoRegistro("PartidaElectronica"))
                
                intRegistro = ObtenerItemLista(arrPaisNacimiento(), adoRegistro("PaisNacimiento"))
                If intRegistro >= Valor_Numero Then cboPaisNacimiento.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrPaisResidencia(), adoRegistro("PaisResidencia"))
                If intRegistro >= Valor_Numero Then cboPaisResidencia.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrSexo(), adoRegistro("SexoCliente"))
                If intRegistro >= Valor_Numero Then cboSexo.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrEstadoCivil(), adoRegistro("EstadoCivil"))
                If intRegistro >= Valor_Numero Then cboEstadoCivil.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrNacionalidad(), adoRegistro("CodNacionalidad"))
                If intRegistro >= Valor_Numero Then cboNacionalidad.ListIndex = intRegistro
                                
                txtDireccionDomicilio1.Text = Trim(adoRegistro("DireccionCliente1"))
                txtDireccionDomicilio2.Text = Trim(adoRegistro("DireccionCliente2"))
                txtEMailDomicilio.Text = Trim(adoRegistro("CorreoCliente"))
                txtTelefonoDomicilio.Text = Trim(adoRegistro("NumTelefono"))
                txtFaxDomicilio.Text = Trim(adoRegistro("NumFax"))
                txtWeb.Text = Trim(adoRegistro("Web"))
                txtCelularDomicilio.Text = Trim(adoRegistro("NumCel"))
                
                intRegistro = ObtenerItemLista(arrPais(), adoRegistro("CodPais"))
                If intRegistro >= Valor_Numero Then cboPais.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrDepartamento(), adoRegistro("CodDepartamento"))
                If intRegistro >= Valor_Numero Then cboDepartamento.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrProvincia(), adoRegistro("CodProvincia"))
                If intRegistro >= Valor_Numero Then cboProvincia.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrDistrito(), adoRegistro("CodDistrito"))
                If intRegistro >= Valor_Numero Then cboDistrito.ListIndex = intRegistro
                
                txtNombreEmpresa.Text = Trim(adoRegistro("CentroTrabajo"))
                txtRUCEmpresa.Text = Trim(adoRegistro("RUCTrabajo"))
                txtOcupCompletar.Text = Trim(adoRegistro("ComplementarOcupacion"))
    
                If Trim(adoRegistro("TipoTrabajador")) = "D" Then
                    optTipoTrabajador(0).Value = True
                    optTipoTrabajador_Click (0)
                    strCodTipoTrab = "D"
                End If
                If Trim(adoRegistro("TipoTrabajador")) = "I" Then
                    optTipoTrabajador(1).Value = True
                    strCodTipoTrab = "I"
                End If
                
                intRegistro = ObtenerItemLista(arrOcupacion(), adoRegistro("OcupacionCliente"))
                If intRegistro >= Valor_Numero Then cboOcupacion.ListIndex = intRegistro
                'cboOcupacion.Text = Trim(adoRegistro("OcupacionCliente"))
                
                txtCargo.Text = Trim(adoRegistro("CargoCliente"))
                txtDireccionTrabajo1.Text = Trim(adoRegistro("DireccionTrabajo1"))
                txtDireccionTrabajo2.Text = Trim(adoRegistro("DireccionTrabajo2"))
                txtEMailTrabajo.Text = Trim(adoRegistro("CorreoTrabajo"))
                txtTelefonoTrabajo.Text = Trim(adoRegistro("NumTelefonoTrabajo"))
                txtFaxTrabajo.Text = Trim(adoRegistro("NumFaxTrabajo"))
                txtCelularTrabajo.Text = Trim(adoRegistro("NumCelTrabajo"))
                
                intRegistro = ObtenerItemLista(arrPaisTrabajo(), adoRegistro("CodPaisTrabajo"))
                If intRegistro >= Valor_Numero Then cboPaisTrabajo.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrDepartamentoTrabajo(), adoRegistro("CodDepartamentoTrabajo"))
                If intRegistro >= Valor_Numero Then cboDepartamentoTrabajo.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrProvinciaTrabajo(), adoRegistro("CodProvinciaTrabajo"))
                If intRegistro >= Valor_Numero Then cboProvinciaTrabajo.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrDistritoTrabajo(), adoRegistro("CodDistritoTrabajo"))
                If intRegistro >= Valor_Numero Then cboDistritoTrabajo.ListIndex = intRegistro
                
                Call CargarDetallePEPS
                Call CargarDetalleGrillaCtaCte
                Call CargarDetalleGrilla
               
                
            End If
            adoRegistro.Close: Set adoRegistro = Nothing
    
    End Select
    
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

Public Sub Buscar()
        
    Dim strSql As String
    
    Set adoConsulta = New ADODB.Recordset
                                                                                    
    Me.MousePointer = vbHourglass
                
    strEstado = Reg_Defecto
        
    strSql = "SELECT CodUnico,DescripParametro TipoIdentidad,NumIdentidad,DescripCliente,FechaIngreso "
    strSql = strSql & "FROM Cliente JOIN AuxiliarParametro ON(AuxiliarParametro.CodParametro=Cliente.TipoIdentidad AND AuxiliarParametro.CodTipoParametro='TIPIDE') "
    If Trim(txtNumDocumento.Text) <> "" And optCriterios(0).Value Then
        strSql = strSql & "WHERE NumIdentidad LIKE '" & Trim(txtNumDocumento.Text) & "%'"
    ElseIf Trim(txtDescripCliente.Text) <> "" And optCriterios(1).Value Then
        strSql = strSql & "WHERE DescripCliente LIKE '%" & Trim(txtDescripCliente.Text) & "%'"
    End If
    
    strSql = strSql + " and IndEstado = '01' "
    
    With adoConsulta
        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open strSql
    End With
    
    tdgConsulta.DataSource = adoConsulta
    
    Me.lblContador.Caption = "Se encontraron " & adoConsulta.RecordCount & " registros."
    
    If adoConsulta.RecordCount > 0 Then strEstado = Reg_Consulta
        
    
    Me.MousePointer = vbDefault
                                    
End Sub

Public Sub Cancelar()

    cmdOpcion.Visible = True
    cmdAccion.Visible = False
    With tabCliente
        .TabEnabled(0) = True
        .TabEnabled(1) = False
        .TabEnabled(2) = False
        .TabEnabled(3) = False
        .TabEnabled(4) = False
        .TabEnabled(5) = False
        .TabEnabled(6) = False
        .Tab = Valor_Numero
    End With
    Call Buscar
    
End Sub

Private Sub Deshabilita()

    fraCliente(1).Enabled = False
    fraCliente(2).Enabled = False
    fraCliente(3).Enabled = False
    
End Sub

Private Sub Habilita()

    fraCliente(1).Enabled = True
    fraCliente(2).Enabled = True
    fraCliente(3).Enabled = True
    
End Sub

Public Sub Imprimir()

End Sub

Public Sub SubImprimir(index As Integer)

    Dim strSeleccionRegistro    As String
    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()

    If tabCliente.Tab = 1 Then Exit Sub
    
    Select Case index
        Case 1
            gstrNameRepo = "Cliente"
                        
            '*** Lista de Inversionistas por rango de fecha ***
            strSeleccionRegistro = "{Cliente.FechaIngreso} IN 'Fch1' TO 'Fch2'"
            gstrSelFrml = strSeleccionRegistro
            frmRangoFecha.Show vbModal
                
            If gstrSelFrml <> "0" Then
            
            Set frmReporte = New frmVisorReporte

            ReDim aReportParamS(1)
            ReDim aReportParamFn(4)
            ReDim aReportParamF(4)
                        
            aReportParamFn(0) = "Usuario"
            aReportParamFn(1) = "Hora"
            aReportParamFn(2) = "NombreEmpresa"
            aReportParamFn(3) = "FechaDel"
            aReportParamFn(4) = "FechaAl"
            
            aReportParamF(0) = gstrLogin
            aReportParamF(1) = Format(Time(), "hh:mm:ss")
            aReportParamF(2) = gstrNombreEmpresa & Space(1)
            aReportParamF(3) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10)
            aReportParamF(4) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10)
                        
            aReportParamS(0) = Convertyyyymmdd(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10))
            aReportParamS(1) = Convertyyyymmdd(DateAdd("d", 1, CVDate(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10))))
            End If
            
        Case 2
            gstrNameRepo = "ListaClientesNaturales"
                        
            '*** Lista de Clientes por rango de fecha ***
            strSeleccionRegistro = "{Cliente.FechaIngreso} IN 'Fch1' TO 'Fch2'"
            gstrSelFrml = strSeleccionRegistro
            'frmRangoFecha.Show vbModal
                
            If gstrSelFrml <> "0" Then
            
                Set frmReporte = New frmVisorReporte
    
                ReDim aReportParamS(0)
                ReDim aReportParamFn(3)
                ReDim aReportParamF(3)
                            
                aReportParamFn(0) = "Usuario"
                aReportParamFn(1) = "Hora"
                aReportParamFn(2) = "NombreEmpresa"
                aReportParamFn(3) = "Fecha"
                
                aReportParamF(0) = gstrLogin
                aReportParamF(1) = Format(Time(), "hh:mm:ss")
                aReportParamF(2) = gstrNombreEmpresa & Space(1)
                aReportParamF(3) = "20900815"
                            
                aReportParamS(0) = "20900815"
            End If
            
        Case 3
            gstrNameRepo = "ListaClientesJuridicos"
                        
            '*** Lista de Clientes por rango de fecha ***
            strSeleccionRegistro = "{Cliente.FechaIngreso} IN 'Fch1' TO 'Fch2'"
            gstrSelFrml = strSeleccionRegistro
            'frmRangoFecha.Show vbModal
                
            If gstrSelFrml <> "0" Then
            
            Set frmReporte = New frmVisorReporte

            ReDim aReportParamS(0)
            ReDim aReportParamFn(3)
            ReDim aReportParamF(3)
                        
            aReportParamFn(0) = "Usuario"
            aReportParamFn(1) = "Hora"
            aReportParamFn(2) = "NombreEmpresa"
            aReportParamFn(3) = "Fecha"
            
            aReportParamF(0) = gstrLogin
            aReportParamF(1) = Format(Time(), "hh:mm:ss")
            aReportParamF(2) = gstrNombreEmpresa & Space(1)
            aReportParamF(3) = "20900815"
             
            aReportParamS(0) = "20900815"
            End If
            
                  
        Case 4
            gstrNameRepo = "ClientePEPS"
                        
            '*** Lista de Clientes por rango de fecha ***
            strSeleccionRegistro = "{Cliente.FechaIngreso} IN 'Fch1' TO 'Fch2'"
            gstrSelFrml = strSeleccionRegistro
            
            frmRangoFecha.Show vbModal
                
            If gstrSelFrml <> "0" Then
            
                Set frmReporte = New frmVisorReporte
    
                ReDim aReportParamS(1)
                ReDim aReportParamFn(4)
                ReDim aReportParamF(4)
                            
                aReportParamFn(0) = "Usuario"
                aReportParamFn(1) = "Hora"
                aReportParamFn(2) = "NombreEmpresa"
                aReportParamFn(3) = "FechaDel"
                aReportParamFn(4) = "FechaAl"
                
                aReportParamF(0) = gstrLogin
                aReportParamF(1) = Format(Time(), "hh:mm:ss")
                aReportParamF(2) = gstrNombreEmpresa & Space(1)
                aReportParamF(3) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10)
                aReportParamF(4) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10)
                            
                aReportParamS(0) = Convertyyyymmdd(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10))
                aReportParamS(1) = Convertyyyymmdd(DateAdd("d", 1, CVDate(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10))))
            End If
            
            
    End Select

    If gstrSelFrml = "0" Then Exit Sub
    
    gstrSelFrml = Valor_Caracter
    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"

    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())

    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
    frmReporte.Show vbModal

    Set frmReporte = Nothing

    Screen.MousePointer = vbNormal
    
End Sub

Public Sub Modificar()

    If strEstado = Reg_Consulta Then
    
        intCodIndti = Valor_Numero
    
        strEstado = Reg_Edicion
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        cmdAccion.Visible = True
        With tabCliente
            .TabEnabled(0) = False
            .TabEnabled(1) = True
            .TabEnabled(2) = True
            .TabEnabled(3) = True
            .TabEnabled(4) = True
            .TabEnabled(5) = True
            .TabEnabled(6) = True
            .Tab = 1
        End With
        'Call Habilita
    End If
    
End Sub

Public Sub Eliminar()
    Dim strSql As String
    
    Me.MousePointer = vbHourglass
    If strEstado = Reg_Consulta Or strEstado = Reg_Edicion Then
        If MsgBox(Mensaje_Eliminacion, vbQuestion + vbYesNo, gstrNombreEmpresa) = vbYes Then
            
            adoComm.CommandText = "Update Cliente set IndEstado='02' where CodUnico = '" & tdgConsulta.Columns("CodUnico") & "'"
        
            adoConn.Execute adoComm.CommandText
            
            Buscar

        End If
    End If
    Me.MousePointer = vbDefault
    
'    If strEstado = Reg_Consulta Or strEstado = Reg_Edicion Then
'        If MsgBox(Mensaje_Eliminacion, vbQuestion + vbYesNo, gstrNombreEmpresa) = vbYes Then
'            adoComm.CommandText = "Sp_INF_SelectMantPart '27', '" & Trim(aBookTipoDoc(cboTipoDocumento.ListIndex - 1)) & "', '" & Trim(txtNumIdentidad.Text) & "'"
'            adoConn.Execute adoComm.CommandText
'
'            txtNumIdentidad.Enabled = True: cboTipoDocumento.Enabled = True
'        End If
'    End If

End Sub

Public Sub Grabar()

    Dim objCuentasCorrientesXML  As DOMDocument60
    Dim objVinculadosRepresentantesXML  As DOMDocument60
    Dim objPEPXML As DOMDocument60
    Dim strCuentasCorrientesXML As String
    Dim strVinculadosRepresentantesXML As String
    Dim strPEPXML As String
    
    Dim intRegistro As Integer
    Dim adoError As ADODB.Error
    Dim strErrMsg As String
    Dim intAccion As Long
    Dim lngNumError As Long
    Dim ClienteSensible As String
    Dim ClienteSensibleOFAC As String
    Dim ClienteSensiblePEPS As String
    Dim ClienteSensible1267 As String
                    
    'VALIDANDO CLIENTES SENSIBLES
    Dim adoRegistroVal As ADODB.Recordset

    Set adoRegistroVal = New ADODB.Recordset

    strCodCliente = Trim(tdgConsulta.Columns(0))
            
    'Listas OFAC
    adoComm.CommandText = "Select * from SDN where SDN_NAME = '" & Trim(txtNombres.Text) & " " & Trim(txtApellidoPaterno.Text) & " " & Trim(txtApellidoMaterno.Text) & "' or SDN_NAME='" & Trim(txtRazonSocial.Text) & "'"
    Set adoRegistroVal = adoComm.Execute
    If adoRegistroVal.RecordCount > 0 Then
        ClienteSensibleOFAC = Valor_Indicador
        ClienteSensible = Valor_Indicador
    End If
    adoRegistroVal.Close: Set adoRegistroVal = Nothing
    
    adoComm.CommandText = "Select * from ALT where ALT_NAME = '" & Trim(txtNombres.Text) & " " & Trim(txtApellidoPaterno.Text) & " " & Trim(txtApellidoMaterno.Text) & "' or ALT_NAME='" & Trim(txtRazonSocial.Text) & "'"
    Set adoRegistroVal = adoComm.Execute
    If adoRegistroVal.RecordCount > 0 Then
        ClienteSensibleOFAC = Valor_Indicador
        ClienteSensible = Valor_Indicador
    End If
    adoRegistroVal.Close: Set adoRegistroVal = Nothing
     
    If ClienteSensibleOFAC <> "" Then
        MsgBox "El Inversionista se encuentra en las Listas OFAC", vbCritical, "Inversionista Sensible"
    End If
    
    'Listas PEPS
    adoComm.CommandText = "Select * from PEPS where NOMBRE = '" & Trim(txtNombres.Text) & " " & Trim(txtApellidoPaterno.Text) & " " & Trim(txtApellidoMaterno.Text) & "' or NOMBRE='" & Trim(txtRazonSocial.Text) & "' or DOC = '" & Trim(txtNumIdentidad.Text) & "'"
    Set adoRegistroVal = adoComm.Execute
    If adoRegistroVal.RecordCount > 0 Then
        ClienteSensiblePEPS = Valor_Indicador
        ClienteSensible = Valor_Indicador
    End If
    adoRegistroVal.Close: Set adoRegistroVal = Nothing
     
    If ClienteSensiblePEPS <> "" Then
        MsgBox "El Inversionista se encuentra en las Listas PEPS", vbCritical, "Inversionista Sensible"
    End If
    
    'Listas Resolucion 1267
    adoComm.CommandText = "Select * from INDIVIDUALALIAS where ALIAS_NAME = '" & Trim(txtNombres.Text) & " " & Trim(txtApellidoPaterno.Text) & " " & Trim(txtApellidoMaterno.Text) & "' or ALIAS_NAME='" & Trim(txtRazonSocial.Text) & "'"
    Set adoRegistroVal = adoComm.Execute
    If adoRegistroVal.RecordCount > 0 Then
        ClienteSensible1267 = Valor_Indicador
        ClienteSensible = Valor_Indicador
    End If
    adoRegistroVal.Close: Set adoRegistroVal = Nothing
    
    adoComm.CommandText = "Select * from INDIVIDUALESOFAC where (FIRST_NAME+' '+SECOND_NAME+' '+THIRD_NAME+' '+FOURTH_NAME) = '" & Trim(txtNombres.Text) & " " & Trim(txtApellidoPaterno.Text) & " " & Trim(txtApellidoMaterno.Text) & "' or (FIRST_NAME+' '+SECOND_NAME+' '+THIRD_NAME+' '+FOURTH_NAME)='" & Trim(txtRazonSocial.Text) & "'"
    Set adoRegistroVal = adoComm.Execute
    If adoRegistroVal.RecordCount > 0 Then
        ClienteSensible1267 = Valor_Indicador
        ClienteSensible = Valor_Indicador
    End If
    adoRegistroVal.Close: Set adoRegistroVal = Nothing
     
    If ClienteSensible1267 <> "" Then
        MsgBox "El Inversionista se encuentra en las Listas  de Resolución 1267", vbCritical, "Inversionista Sensible"
    End If
    ''''''''''''''''''''''''''''''''''''''''''''
    
    If strEstado = Reg_Defecto Then Exit Sub
    
    If Not TodoOK() Then Exit Sub
    
    On Error GoTo CtrlError
    
    Call XMLADORecordset(objCuentasCorrientesXML, "ClienteBancarios", "CuentaCorriente", adoRegistroAuxCtaCte, strErrMsg)
    strCuentasCorrientesXML = objCuentasCorrientesXML.xml
    
    Call XMLADORecordset(objPEPXML, "ClientePEP", "PEP", adoRegistroAuxPEP, strErrMsg)
    strPEPXML = objPEPXML.xml
    
    If strCodClasePersona = Codigo_Persona_Natural Then
        Call XMLADORecordset(objVinculadosRepresentantesXML, "ClienteDependientes", "Vinculado", adoRegistroAux, strErrMsg)
        strVinculadosRepresentantesXML = objVinculadosRepresentantesXML.xml
    End If
    If strCodClasePersona = Codigo_Persona_Juridica Then
        Call XMLADORecordset(objVinculadosRepresentantesXML, "ClienteRepresentantes", "Representante", adoRegistroAux, strErrMsg)
        strVinculadosRepresentantesXML = objVinculadosRepresentantesXML.xml
    End If
    
    If chkPepSi.Value = 1 Then
        strInstitucionPEPS = Trim(txtInstitucionPEPS.Text)
        strCargoPEPS = Trim(txtCargoDesemPEPS.Text)
        ClienteSensible = Valor_Indicador
        If chkAdminRecPubSi.Value = 1 Then
            strRecursosPublicos = Valor_Indicador
        Else
            strRecursosPublicos = Valor_Caracter
        End If
    Else
         ClienteSensible = Valor_Caracter
        strInstitucionPEPS = Valor_Caracter
        strCargoPEPS = Valor_Caracter
        strRecursosPublicos = Valor_Caracter
    End If
    

    If strEstado = Reg_Adicion Then
        Me.MousePointer = vbHourglass
        
        '*** Guardar Inversionista ***
        With adoComm
        
            .CommandText = "{ call up_PRManCliente('"
            .CommandText = .CommandText & Trim(lblCodigoCliente.Caption) & "','"
            .CommandText = .CommandText & strCodTipoDocumento & "','" & Trim(txtOtroDocumento.Text) & "','"
            .CommandText = .CommandText & Trim(txtNumIdentidad.Text) & "','"
            If strCodClasePersona = Codigo_Persona_Juridica Then
                .CommandText = .CommandText & Trim(txtNumIdentidad.Text) & "','"
            Else
                .CommandText = .CommandText & "','"
            End If
            .CommandText = .CommandText & Trim(txtApellidoPaterno.Text) & "','"
            .CommandText = .CommandText & Trim(txtApellidoMaterno.Text) & "','"
            .CommandText = .CommandText & Trim(txtNombres.Text) & "','"
            If strCodClasePersona = Codigo_Persona_Juridica Then
                .CommandText = .CommandText & Trim(txtRazonSocial.Text) & "','"
                .CommandText = .CommandText & Trim(txtRazonSocial.Text) & "','"
            Else
                .CommandText = .CommandText & Trim(txtApellidoPaterno.Text) & Space(1) & Trim(txtApellidoMaterno.Text) & Space(1) & Trim(txtNombres.Text) & "','"
                .CommandText = .CommandText & "','"
            End If
            .CommandText = .CommandText & Estado_Activo & "','"
            .CommandText = .CommandText & strCodClasePersona & "','"
            .CommandText = .CommandText & strCodSexo & "','"
            .CommandText = .CommandText & strCodEstadoCivil & "','"
            .CommandText = .CommandText & Convertyyyymmdd(dtpFechaNacimiento.Value) & "','"
            .CommandText = .CommandText & strCodPaisNacimiento & "','"
            .CommandText = .CommandText & strCodNacionalidad & "','"
            .CommandText = .CommandText & strCodPaisResidencia & "','"
            .CommandText = .CommandText & Convertyyyymmdd(dtpFechaConstitucion.Value) & "','"
            .CommandText = .CommandText & Trim(txtObjetivoSocial.Text) & "','"
            .CommandText = .CommandText & Trim(txtPartidaElectronica.Text) & "','"
            .CommandText = .CommandText & Convertyyyymmdd(CVDate(lblFechaIngreso.Caption)) & "','"
            .CommandText = .CommandText & strCodTipoTrab & "','"
            .CommandText = .CommandText & Trim(txtRUCEmpresa.Text) & "','"
            .CommandText = .CommandText & Trim(txtNombreEmpresa.Text) & "','"
            '.CommandText = .CommandText & Trim(txtOcupacion.Text) & "','"
            .CommandText = .CommandText & strCodOcupacion & "','"
            .CommandText = .CommandText & Trim(txtOcupCompletar.Text) & "','"
            .CommandText = .CommandText & Trim(txtCargo.Text) & "','"
            .CommandText = .CommandText & Trim(txtDireccionDomicilio1.Text) & "','"
            .CommandText = .CommandText & Trim(txtDireccionDomicilio2.Text) & "','"
            .CommandText = .CommandText & Trim(txtEMailDomicilio.Text) & "','"
            .CommandText = .CommandText & strCodDepartamento & "','"
            .CommandText = .CommandText & strCodProvincia & "','"
            .CommandText = .CommandText & strCodDistrito & "','"
            .CommandText = .CommandText & strCodPais & "','"
            .CommandText = .CommandText & Trim(txtTelefonoDomicilio.Text) & "','"
            .CommandText = .CommandText & Trim(txtFaxDomicilio.Text) & "','"
            .CommandText = .CommandText & Trim(txtCelularDomicilio.Text) & "','"
            .CommandText = .CommandText & Trim(txtWeb.Text) & "','"
            .CommandText = .CommandText & Trim(strCodEnvioInformacion) & "','"
            .CommandText = .CommandText & Trim(txtDireccionTrabajo1.Text) & "','"
            .CommandText = .CommandText & Trim(txtDireccionTrabajo2.Text) & "','"
            .CommandText = .CommandText & Trim(txtEMailTrabajo.Text) & "','"
            .CommandText = .CommandText & strCodDepartamentoTrabajo & "','"
            .CommandText = .CommandText & strCodProvinciaTrabajo & "','"
            .CommandText = .CommandText & strCodDistritoTrabajo & "','"
            .CommandText = .CommandText & strCodPaisTrabajo & "','"
            .CommandText = .CommandText & Trim(txtTelefonoTrabajo.Text) & "','"
            .CommandText = .CommandText & Trim(txtFaxTrabajo.Text) & "','"
            .CommandText = .CommandText & Trim(txtCelularTrabajo.Text) & "','"
            .CommandText = .CommandText & "','"
            .CommandText = .CommandText & gstrLogin & "','"
            .CommandText = .CommandText & Convertyyyymmdd(gdatFechaActual) & "','"
            .CommandText = .CommandText & gstrLogin & "','"
            .CommandText = .CommandText & Convertyyyymmdd(gdatFechaActual) & "','"
            .CommandText = .CommandText & ClienteSensible & "','"
            .CommandText = .CommandText & ClienteSensibleOFAC & "','"
            .CommandText = .CommandText & strPEPS & "','"
'            .CommandText = .CommandText & strInstitucionPEPS & "','"
'            .CommandText = .CommandText & strCargoPEPS & "','"
'            .CommandText = .CommandText & strRecursosPublicos & "','"
'            .CommandText = .CommandText & Convertyyyymmdd(dtpPepFechaDesde) & "','"
'            .CommandText = .CommandText & Convertyyyymmdd(dtpPepFechaHasta) & "','"
            .CommandText = .CommandText & ClienteSensible1267 & "','"
            
            '.CommandText = .CommandText & "I') }"
            .CommandText = .CommandText & strCuentasCorrientesXML & "','" & strVinculadosRepresentantesXML & "','" & strPEPXML & "','I') }"
            'MsgBox .CommandText, vbCritical
            adoConn.Execute .CommandText
                                        
        End With

        txtNumDocumento.Text = Trim(txtNumIdentidad.Text)

    End If
    
    If strEstado = Reg_Edicion Then
            
        If MsgBox(Mensaje_Edicion, vbQuestion + vbYesNo, gstrNombreEmpresa) = vbNo Then Exit Sub
        
        Me.MousePointer = vbHourglass
                                                        
        With adoComm
            '*** Actualizar Inversionista ***
            .CommandText = "{ call up_PRManCliente('" & _
                Trim(lblCodigoCliente.Caption) & "','" & strCodTipoDocumento & "','" & Trim(txtOtroDocumento.Text) & "','" & _
                Trim(txtNumIdentidad.Text) & "','"
            If strCodClasePersona = Codigo_Persona_Juridica Then
                .CommandText = .CommandText & Trim(txtNumIdentidad.Text) & "','"
            Else
                .CommandText = .CommandText & "','"
            End If
            .CommandText = .CommandText & Trim(txtApellidoPaterno.Text) & "','" & _
                Trim(txtApellidoMaterno.Text) & "','" & Trim(txtNombres.Text) & "','"
            If strCodClasePersona = Codigo_Persona_Juridica Then
                .CommandText = .CommandText & Trim(txtRazonSocial.Text) & "','"
                .CommandText = .CommandText & Trim(txtRazonSocial.Text) & "','"
            Else
                .CommandText = .CommandText & Trim(txtApellidoPaterno.Text) & Space(1) & Trim(txtApellidoMaterno.Text) & Space(1) & Trim(txtNombres.Text) & "','"
                .CommandText = .CommandText & "','"
            End If
            
            
            .CommandText = .CommandText & Estado_Activo & "','" & _
                strCodClasePersona & "','" & strCodSexo & "','" & _
                strCodEstadoCivil & "','" & Convertyyyymmdd(dtpFechaNacimiento.Value) & "','" & strCodPaisNacimiento & "','" & _
                strCodNacionalidad & "','" & strCodPaisResidencia & "','" & Convertyyyymmdd(dtpFechaConstitucion.Value) & "','" & _
                Trim(txtObjetivoSocial.Text) & "','" & Trim(txtPartidaElectronica.Text) & "','" & _
                Convertyyyymmdd(CVDate(lblFechaIngreso.Caption)) & "','" & strCodTipoTrab & "','" & Trim(txtRUCEmpresa.Text) & "','" & _
                Trim(txtNombreEmpresa.Text) & "','" & strCodOcupacion & "','" & Trim(txtOcupCompletar.Text) & "','" & _
                Trim(txtCargo.Text) & "','" & Trim(txtDireccionDomicilio1.Text) & "','" & _
                Trim(txtDireccionDomicilio2.Text) & "','" & Trim(txtEMailDomicilio.Text) & "','" & _
                strCodDepartamento & "','" & strCodProvincia & "','" & _
                strCodDistrito & "','" & strCodPais & "','" & _
                Trim(txtTelefonoDomicilio.Text) & "','" & Trim(txtFaxDomicilio.Text) & "','" & Trim(txtCelularDomicilio.Text) & "','" & Trim(txtWeb.Text) & "','" & Trim(strCodEnvioInformacion) & "','" & _
                Trim(txtDireccionTrabajo1.Text) & "','" & Trim(txtDireccionTrabajo2.Text) & "','" & _
                Trim(txtEMailTrabajo.Text) & "','" & strCodDepartamentoTrabajo & "','" & _
                strCodProvinciaTrabajo & "','" & strCodDistritoTrabajo & "','" & _
                strCodPaisTrabajo & "','" & Trim(txtTelefonoTrabajo.Text) & "','" & _
                Trim(txtFaxTrabajo.Text) & "','" & Trim(txtCelularTrabajo.Text) & "','','" & _
                gstrLogin & "','" & Convertyyyymmdd(gdatFechaActual) & "','" & _
                gstrLogin & "','" & Convertyyyymmdd(gdatFechaActual) & "','" & _
                ClienteSensible & "','" & ClienteSensibleOFAC & "','" & _
                strPEPS & "','" & _
                ClienteSensible1267 & "','" & _
                strCuentasCorrientesXML & "','" & strVinculadosRepresentantesXML & "','" & strPEPXML & "','U') }"
                
'                 & strInstitucionPEPS & "' ,'" & strCargoPEPS & "' ,'" & strRecursosPublicos & "','" & _
'                 Convertyyyymmdd(dtpPepFechaDesde) & "','" & Convertyyyymmdd(dtpPepFechaHasta) & "','" & _

                'MsgBox .CommandText, vbCritical
            adoConn.Execute .CommandText
            
            '*** Actualizar Contrato ***
            .CommandText = "UPDATE ParticipeContrato SET " & _
                "ApellidoPaterno='" & Trim(txtApellidoPaterno.Text) & "'," & _
                "ApellidoMaterno='" & Trim(txtApellidoMaterno.Text) & "'," & _
                "Nombres='" & Trim(txtNombres.Text) & "',"
            If strCodClasePersona = Codigo_Persona_Juridica Then
                .CommandText = .CommandText & "RazonSocial='" & Trim(txtRazonSocial.Text) & "'," & _
                    "DescripParticipe='" & Trim(txtRazonSocial.Text) & "',"
            Else
                .CommandText = .CommandText & "RazonSocial='',"
            End If
            .CommandText = .CommandText & "SexoParticipe='" & strCodSexo & "'," & _
                "EstadoCivil='" & strCodEstadoCivil & "'," & _
                "FechaNacimiento='" & Convertyyyymmdd(dtpFechaNacimiento.Value) & "'," & _
                "CodNacionalidad='" & strCodNacionalidad & "'," & _
                "NumTelefono='" & Trim(txtTelefonoDomicilio.Text) & "'," & _
                "NumFax='" & Trim(txtFaxDomicilio.Text) & "' " & _
                "WHERE CodUnico='" & strCodCliente & "'"
            adoConn.Execute .CommandText
            
            '*** Actualizar Dirección Postal del Contrato ***
            .CommandText = "UPDATE ParticipeContrato SET " & _
                "DireccionPostal1='" & Trim(txtDireccionDomicilio1.Text) & "'," & _
                "DireccionPostal2='" & Trim(txtDireccionDomicilio2.Text) & "'," & _
                "CodPais='" & strCodPais & "'," & _
                "CodDepartamento='" & strCodDepartamento & "'," & _
                "CodProvincia='" & strCodProvincia & "'," & _
                "CodDistrito='" & strCodDistrito & "' " & _
                "WHERE CodUnico='" & strCodCliente & "' AND TipoCorreoPostal='" & Codigo_Dirección_Domicilio & "'"
            adoConn.Execute .CommandText
            
            .CommandText = "UPDATE ParticipeContrato SET " & _
                "DireccionPostal1='" & Trim(txtDireccionTrabajo1.Text) & "'," & _
                "DireccionPostal2='" & Trim(txtDireccionTrabajo2.Text) & "'," & _
                "CodPais='" & strCodPais & "'," & _
                "CodDepartamento='" & strCodDepartamento & "'," & _
                "CodProvincia='" & strCodProvincia & "'," & _
                "CodDistrito='" & strCodDistrito & "' " & _
                "WHERE CodUnico='" & strCodCliente & "' AND TipoCorreoPostal='" & Codigo_Dirección_Trabajo & "'"
            adoConn.Execute .CommandText
                                
            '*** Actualizar Código Unico si hubo cambio ***
            If strCodCliente <> Trim(lblCodigoCliente.Caption) Then
                '*** Cliente ***
                .CommandText = "UPDATE Cliente SET " & _
                    "CodUnico='" & Trim(lblCodigoCliente.Caption) & "'," & _
                    "TipoIdentidad='" & strCodTipoDocumento & "'," & _
                    "NumIdentidad='" & Trim(txtNumIdentidad.Text) & "'"
                If strCodClasePersona = Codigo_Persona_Juridica Then
                    .CommandText = .CommandText & ",NumRuc='" & Trim(txtNumIdentidad.Text) & "'"
                Else
                    .CommandText = .CommandText & ",NumRuc=''"
                End If
                .CommandText = .CommandText & " WHERE CodUnico='" & strCodCliente & "'"
                
                adoConn.Execute .CommandText
                
                '*** Contrato ***
                .CommandText = "UPDATE ParticipeContrato SET " & _
                    "CodUnico='" & Trim(lblCodigoCliente.Caption) & "'," & _
                    "TipoIdentidad='" & strCodTipoDocumento & "'," & _
                    "NumIdentidad='" & Trim(txtNumIdentidad.Text) & "' " & _
                    "WHERE CodUnico='" & strCodCliente & "'"
                adoConn.Execute .CommandText
                
                '*** Detalle del Contrato ***
                .CommandText = "UPDATE ParticipeContratoDetalle SET " & _
                    "CodCliente='" & Trim(lblCodigoCliente.Caption) & "'," & _
                    "TipoIdentidad='" & strCodTipoDocumento & "'," & _
                    "NumIdentidad='" & Trim(txtNumIdentidad.Text) & "' " & _
                    "WHERE CodCliente='" & strCodCliente & "'"
                adoConn.Execute .CommandText
            End If
                            
            '*** Actualizar la Descripción del Partícipe ***
            If strCodClasePersona = Codigo_Persona_Natural Then
                Dim adoRegistro As ADODB.Recordset
                
                Set adoRegistro = New ADODB.Recordset
                
                .CommandText = "SELECT CodParticipe,TipoMancomuno FROM ParticipeContrato " & _
                    "WHERE CodUnico='" & Trim(lblCodigoCliente.Caption) & "'"
                Set adoRegistro = .Execute
                
                If Not adoRegistro.EOF Then
                    .CommandText = "{ call up_PRActMancomuno('" & Trim(adoRegistro("CodParticipe")) & "') }"
                   
'                        ,'" & _ Trim (lblCodigoCliente.Caption) & "','" & Trim(adoRegistro("TipoMancomuno")) & "') }"
                    adoConn.Execute .CommandText
                    
                    adoRegistro.Close: Set adoRegistro = Nothing
                Else
                    adoRegistro.Close
                    
                    .CommandText = "SELECT CodParticipe,TipoMancomuno FROM ParticipeContratoDetalle " & _
                        "WHERE CodCliente='" & Trim(lblCodigoCliente.Caption) & "'"
                    Set adoRegistro = .Execute
                    
                    Do While Not adoRegistro.EOF
                        .CommandText = "{ call up_PRActMancomuno('" & Trim(adoRegistro("CodParticipe")) & "','" & _
                            Trim(lblCodigoCliente.Caption) & "','" & Trim(adoRegistro("TipoMancomuno")) & "') }"
                        adoConn.Execute .CommandText
                    
                        adoRegistro.MoveNext
                    Loop
                    adoRegistro.Close: Set adoRegistro = Nothing
                End If
            End If
        End With

    End If

    Me.MousePointer = vbDefault
    MsgBox Mensaje_Edicion_Exitosa, vbExclamation
    frmMainMdi.stbMdi.Panels(3).Text = "Acción"

    chkPepNo.Value = 1
    
    cmdOpcion.Visible = True
    cmdAccion.Visible = False
    With tabCliente
        .TabEnabled(0) = True
        .Tab = Valor_Numero
    End With
    Call Buscar
    
    Exit Sub

CtrlError:
    If adoConn.Errors.Count > 0 Then
        For Each adoError In adoConn.Errors
            strErrMsg = strErrMsg & adoError.Description & " (" & adoError.NativeError & ") " & Chr(13)
        Next
        Me.MousePointer = vbDefault
        MsgBox strErrMsg, vbCritical + vbOKOnly, Me.Caption
    Else
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
    
        End If

End Sub

Public Sub Salir()

    Unload Me
    
End Sub



Private Sub btnAgregarCtaCte_Click()

    Dim adoRegistroCtaCte As ADODB.Recordset
    Dim intSecuencial As Integer
    Dim dblBookmark As Double
    
    If TodoOkCtaCte() Then
            adoRegistroAuxCtaCte.AddNew
            'adoRegistroAuxCtaCte.Fields("CodCliente") = Trim(strCodCliente)
            adoRegistroAuxCtaCte.Fields("Banco") = Trim(arrBancos(cboBancos.ListIndex))
            adoRegistroAuxCtaCte.Fields("BancoDesc") = Trim(cboBancos.List(cboBancos.ListIndex))
            adoRegistroAuxCtaCte.Fields("TipoCtaCte") = Trim(arrTipoCta(cboTipoCta.ListIndex))
            adoRegistroAuxCtaCte.Fields("TipoCtaCteDesc") = Trim(cboTipoCta.List(cboTipoCta.ListIndex))
            adoRegistroAuxCtaCte.Fields("TipoMoneda") = Trim(arrMoneda(cboMoneda.ListIndex))
            adoRegistroAuxCtaCte.Fields("TipoMonedaDesc") = Trim(cboMoneda.List(cboMoneda.ListIndex))
            adoRegistroAuxCtaCte.Fields("CCI") = Trim(txtCCI.Text)
            adoRegistroAuxCtaCte.Fields("NroCtaCte") = Trim(txtCtaCte.Text)
             
            adoRegistroAuxCtaCte.Update
            
            dblBookmark = adoRegistroAuxCtaCte.Bookmark
            
            TDBCtasCtes.Refresh
              
            
            btnEliminarCtaCte.Enabled = True
            
            Call LimpiarDatosCtaCte
    End If
End Sub

Private Sub btnDepAgregar_Click()

    Dim adoRegistro As ADODB.Recordset
    Dim intSecuencial As Integer
    Dim dblBookmark As Double
    
    If TodoOkListas() Then
            adoRegistroAux.AddNew
            'adoRegistroAux.Fields("CodCliente") = strCodCliente
            adoRegistroAux.Fields("NombresApellidos") = Trim(txtNombresLaboral.Text)
            adoRegistroAux.Fields("TipoDocumento") = arrTipoDocVinculado(cboTipoDocLaboral.ListIndex)
            adoRegistroAux.Fields("TipoDocDesc") = cboTipoDocLaboral.List(cboTipoDocLaboral.ListIndex)
            adoRegistroAux.Fields("NroDocumento") = Trim(txtDocumentoLaboral.Text)
            adoRegistroAux.Fields("Vinculacion") = arrVinculacion(cboVinculacion.ListIndex)
            adoRegistroAux.Fields("VincDesc") = cboVinculacion.List(cboVinculacion.ListIndex)
        
            adoRegistroAux.Update
            
            dblBookmark = adoRegistroAux.Bookmark
            
            TDBDependientes.Refresh
            
            'Call NumerarRegistros
            
            'adoRegistroAux.Bookmark = dblBookmark
            
            
            btnDepEliminar.Enabled = True
            
            Call LimpiarDatosDep
    End If
        
End Sub

Private Sub btnDepEditar_Click()
    If TodoOkListas() Then
            adoRegistroAux.Fields("NombresApellidos") = Trim(txtNombresLaboral.Text)
            adoRegistroAux.Fields("TipoDocumento") = arrTipoDocVinculado(cboTipoDocLaboral.ListIndex)
            adoRegistroAux.Fields("TipoDocDesc") = cboTipoDocLaboral.List(cboTipoDocLaboral.ListIndex)
            adoRegistroAux.Fields("NroDocumento") = Trim(txtDocumentoLaboral.Text)
            adoRegistroAux.Fields("Vinculacion") = arrVinculacion(cboVinculacion.ListIndex)
            adoRegistroAux.Fields("VincDesc") = cboVinculacion.List(cboVinculacion.ListIndex)
  
    End If
End Sub

Private Sub btnDepEliminar_Click()
    Dim dblBookmark As Double
    
        If adoRegistroAux.RecordCount > 0 Then
        
            dblBookmark = adoRegistroAux.Bookmark
        
            adoRegistroAux.Delete adAffectCurrent
            
            If adoRegistroAux.EOF Then
                adoRegistroAux.MovePrevious
                TDBDependientes.MovePrevious
            End If
                
            adoRegistroAux.Update
            
            If adoRegistroAux.RecordCount = Valor_Numero Then btnDepEliminar.Enabled = False
    
            If adoRegistroAux.RecordCount > 0 And Not adoRegistroAux.BOF And Not adoRegistroAux.EOF And dblBookmark > 1 Then adoRegistroAux.Bookmark = dblBookmark - 1
            
            TDBDependientes.Refresh
        
        End If
End Sub

Private Sub btnEditarCtaCte_Click()
    If TodoOkCtaCte() Then
            'adoRegistroAuxCtaCte.Fields("CodCliente") = Trim(strCodCliente)
            adoRegistroAuxCtaCte.Fields("Banco") = Trim(arrBancos(cboBancos.ListIndex))
            adoRegistroAuxCtaCte.Fields("BancoDesc") = Trim(cboBancos.List(cboBancos.ListIndex))
            adoRegistroAuxCtaCte.Fields("TipoCtaCte") = Trim(arrTipoCta(cboTipoCta.ListIndex))
            adoRegistroAuxCtaCte.Fields("TipoCtaCteDesc") = Trim(cboTipoCta.List(cboTipoCta.ListIndex))
            adoRegistroAuxCtaCte.Fields("TipoMoneda") = Trim(arrMoneda(cboMoneda.ListIndex))
            adoRegistroAuxCtaCte.Fields("TipoMonedaDesc") = Trim(cboMoneda.List(cboMoneda.ListIndex))
            adoRegistroAuxCtaCte.Fields("CCI") = Trim(txtCCI.Text)
            adoRegistroAuxCtaCte.Fields("NroCtaCte") = Trim(txtCtaCte.Text)
            
            Call LimpiarDatosCtaCte '20141128_JJCC
    End If
End Sub

Private Sub btnEliminarCtaCte_Click()

    Dim dblBookmark As Double
    
    
        If adoRegistroAuxCtaCte.RecordCount > 0 Then
        
            dblBookmark = adoRegistroAuxCtaCte.Bookmark
        
            adoRegistroAuxCtaCte.Delete adAffectCurrent
            
            If adoRegistroAuxCtaCte.EOF Then
                adoRegistroAuxCtaCte.MovePrevious
                TDBCtasCtes.MovePrevious
            End If
                
            adoRegistroAuxCtaCte.Update
            
            If adoRegistroAuxCtaCte.RecordCount = Valor_Numero Then btnEliminarCtaCte.Enabled = False
    
            If adoRegistroAuxCtaCte.RecordCount > 0 And Not adoRegistroAuxCtaCte.BOF And Not adoRegistroAuxCtaCte.EOF And dblBookmark > 1 Then adoRegistroAuxCtaCte.Bookmark = dblBookmark - 1
            
            TDBCtasCtes.Refresh
        
        End If
    
    
End Sub

Private Sub btnPEPAgregar_Click()

    Dim adoRegistroPEP As ADODB.Recordset
    Dim intSecuencial As Integer
    Dim dblBookmark As Double
    Dim strSql As String
    
            
    If TodoOKPEP() Then
    
            intCodIndti = intCodIndti + 1
                        
            adoRegistroAuxPEP.AddNew
            
            adoRegistroAuxPEP.Fields("CodUnico") = lblCodigoCliente.Caption
            adoRegistroAuxPEP.Fields("CodInstitucionPEPS") = intCodIndti
            adoRegistroAuxPEP.Fields("InstitucionPEPS") = Trim(txtInstitucionPEPS.Text)
            adoRegistroAuxPEP.Fields("CargoInstitucionPEPS") = Trim(txtCargoDesemPEPS.Text)
            
            If (chkAdminRecPubSi.Value = 1) Then
            adoRegistroAuxPEP.Fields("RecursosPublicosPEPS") = "Si"
            Else
            adoRegistroAuxPEP.Fields("RecursosPublicosPEPS") = "No"
            End If
            
            adoRegistroAuxPEP.Fields("FecDesdePEPS") = dtpPepFechaDesde.Value
            adoRegistroAuxPEP.Fields("FecHastaPEPS") = dtpPepFechaHasta.Value
             
            adoRegistroAuxPEP.Update
            
            dblBookmark = adoRegistroAuxPEP.Bookmark
            
            TDBPEP.Refresh
            
            btnPEPEliminar.Enabled = True
            
            TDBPEP.DataSource = adoRegistroAuxPEP
            
            Call LimpiarDatosPEP
    End If
End Sub

Private Sub btnPEPEditar_Click()
    If TodoOKPEP() Then
            adoRegistroAuxPEP.Fields("InstitucionPEPS") = Trim(txtInstitucionPEPS.Text)
            adoRegistroAuxPEP.Fields("CargoInstitucionPEPS") = Trim(txtCargoDesemPEPS.Text)
            If (chkAdminRecPubSi.Value = 1) Then adoRegistroAuxPEP.Fields("RecursosPublicosPEPS") = "Si"
            If (chkAdminRecPubNo.Value = 1) Then adoRegistroAuxPEP.Fields("RecursosPublicosPEPS") = "No"
            adoRegistroAuxPEP.Fields("FecDesdePEPS") = dtpPepFechaDesde.Value
            adoRegistroAuxPEP.Fields("FecHastaPEPS") = dtpPepFechaHasta.Value
            
            Call LimpiarDatosPEP '20141128_JJCC
    End If
End Sub

Private Sub btnPEPEliminar_Click()
    Dim dblBookmark As Double
    
    If TDBPEP.SelBookmarks.Count >= 1 Then
    
        If adoRegistroAuxPEP.RecordCount > 0 Then
        
            dblBookmark = adoRegistroAuxPEP.Bookmark
        
            adoRegistroAuxPEP.Delete adAffectCurrent
            
            If adoRegistroAuxPEP.EOF Then
                adoRegistroAuxPEP.MovePrevious
                TDBPEP.MovePrevious
            End If
                
            adoRegistroAuxPEP.Update
            
            If adoRegistroAuxPEP.RecordCount = Valor_Numero Then btnPEPEliminar.Enabled = False
    
            If adoRegistroAuxPEP.RecordCount > 0 And Not adoRegistroAuxPEP.BOF And Not adoRegistroAuxPEP.EOF And dblBookmark > 1 Then
            
            adoRegistroAuxPEP.Bookmark = dblBookmark - 1
            
            
            
            End If
            
            If adoRegistroAuxPEP.RecordCount >= 1 Then
            adoRegistroAuxPEP.MoveLast
            End If
            
            TDBPEP.Refresh
        
        End If
    
    Else
        MsgBox "Debe seleccinar un Registro", vbCritical, Me.Caption
    End If

    
        Call LimpiarDatosPEP
End Sub

Private Sub btnRepAgregar_Click()

    Dim adoRegistro As ADODB.Recordset
    Dim intSecuencial As Integer
    Dim dblBookmark As Double
            
    If TodoOkListas() Then
            adoRegistroAux.AddNew
            'adoRegistroAux.Fields("CodCliente") = Trim(lblCodigoCliente)
            adoRegistroAux.Fields("CodRepresentante") = Format(arrTipoDocRepresentante(cboTipoDocLegales.ListIndex) & Trim(txtDocumentoLegales.Text) & "01", "000000000000000")
            adoRegistroAux.Fields("Nombres") = Trim(txtNombresLegales.Text)
            adoRegistroAux.Fields("Apellidos") = Trim(txtApellidosLegales.Text)
            adoRegistroAux.Fields("TipoDocumento") = arrTipoDocRepresentante(cboTipoDocLegales.ListIndex)
            adoRegistroAux.Fields("TipoDocDesc") = cboTipoDocLegales.List(cboTipoDocLegales.ListIndex)
            adoRegistroAux.Fields("NroDocumento") = Trim(txtDocumentoLegales.Text)
            adoRegistroAux.Fields("VinculacionRep") = arrVinculacionLegales(cboVinculacionLegales.ListIndex)
            adoRegistroAux.Fields("VincDesc") = cboVinculacionLegales.List(cboVinculacionLegales.ListIndex)
             
            adoRegistroAux.Update
            
            dblBookmark = adoRegistroAux.Bookmark
            
            TDBRepresentantes.Refresh
              
            
            btnRepEliminar.Enabled = True
            
            Call LimpiarDatosRep
    End If
End Sub

Private Sub btnRepEditar_Click()
    If TodoOkListas() Then
            adoRegistroAux.Fields("Nombres") = Trim(txtNombresLegales.Text)
            adoRegistroAux.Fields("Apellidos") = Trim(txtApellidosLegales.Text)
            adoRegistroAux.Fields("TipoDocumento") = arrTipoDocRepresentante(cboTipoDocLegales.ListIndex)
            adoRegistroAux.Fields("TipoDocDesc") = cboTipoDocLegales.List(cboTipoDocLegales.ListIndex)
            adoRegistroAux.Fields("NroDocumento") = Trim(txtDocumentoLegales.Text)
            adoRegistroAux.Fields("VinculacionRep") = arrVinculacionLegales(cboVinculacionLegales.ListIndex)
            adoRegistroAux.Fields("VincDesc") = cboVinculacionLegales.List(cboVinculacionLegales.ListIndex)
            
            Call LimpiarDatosRep '20141128_JJCC
    End If
End Sub

Private Sub btnRepEliminar_Click()
    Dim dblBookmark As Double
    
        If adoRegistroAux.RecordCount > 0 Then
        
            dblBookmark = adoRegistroAux.Bookmark
        
            adoRegistroAux.Delete adAffectCurrent
            
            If adoRegistroAux.EOF Then
                adoRegistroAux.MovePrevious
                TDBRepresentantes.MovePrevious
            End If
                
            adoRegistroAux.Update
            
            If adoRegistroAux.RecordCount = Valor_Numero Then btnRepEliminar.Enabled = False
    
            If adoRegistroAux.RecordCount > 0 And Not adoRegistroAux.BOF And Not adoRegistroAux.EOF And dblBookmark > 1 Then adoRegistroAux.Bookmark = dblBookmark - 1
            
            TDBRepresentantes.Refresh
        
        End If
End Sub

Private Sub cboEnvioInformacion_Click()

    Dim strSql As String
    
    strCodEnvioInformacion = Valor_Caracter
    If cboEnvioInformacion.ListIndex < 0 Then Exit Sub
    
    strCodEnvioInformacion = Trim(arrEnvioInformacion(cboEnvioInformacion.ListIndex))
    
End Sub

Private Sub cboOcupacion_Click()

    Dim strSql As String
    
    strCodOcupacion = Valor_Caracter
    If cboOcupacion.ListIndex < 0 Then Exit Sub
    
    strCodOcupacion = Trim(arrOcupacion(cboOcupacion.ListIndex))

End Sub

Private Sub cboClasePersona_Click()

    Dim strSql As String
    
    strCodClasePersona = Valor_Caracter
    If cboClasePersona.ListIndex < 0 Then Exit Sub
    
    strCodClasePersona = Trim(arrClasePersona(cboClasePersona.ListIndex))
    
    '*** Tipo Documento Identidad ***
    strSql = "{ call up_ACSelDatosParametro(4,'" & strCodClasePersona & "') }"
    CargarControlLista strSql, cboTipoDocumento, arrTipoDocumento(), Sel_Defecto
    
    If cboTipoDocumento.ListCount > 0 Then cboTipoDocumento.ListIndex = Valor_Numero
    
    '*** Tipo Documento Identidad Doc Laboral ***
    strSql = "{ call up_ACSelDatosParametro(4,'" & Codigo_Persona_Natural & "') }"
    CargarControlLista strSql, cboTipoDocLaboral, arrTipoDocVinculado(), Sel_Defecto
    If cboTipoDocLaboral.ListCount > 0 Then cboTipoDocLaboral.ListIndex = Valor_Numero
    
    '*** Tipo Documento Identidad Doc Representantes ***
    strSql = "{ call up_ACSelDatosParametro(4,'" & Codigo_Persona_Natural & "') }"
    CargarControlLista strSql, cboTipoDocLegales, arrTipoDocRepresentante(), Sel_Defecto
    If cboTipoDocLegales.ListCount > 0 Then cboTipoDocLegales.ListIndex = Valor_Numero
       

        
    If strCodClasePersona = Codigo_Persona_Natural Then
    
        txtApellidoPaterno.Enabled = True
        txtApellidoMaterno.Enabled = True
        txtNombres.Enabled = True
        cboSexo.Enabled = True
        cboEstadoCivil.Enabled = True
        dtpFechaNacimiento.Enabled = True
        dtpFechaConstitucion.Enabled = False
        dtpFechaConstitucion.Value = Valor_Fecha
        txtPartidaElectronica.Enabled = False
        txtPartidaElectronica.Text = Valor_Caracter
        txtObjetivoSocial.Enabled = False
        txtObjetivoSocial.Text = Valor_Caracter
        txtWeb.Enabled = False
        txtWeb.Text = Valor_Caracter
        txtFaxDomicilio.Enabled = False
        txtFaxDomicilio.Visible = False
        txtCelularDomicilio.Enabled = True
        txtCelularDomicilio.Visible = True
        lblDescrip(12).Caption = "Celular"
        txtEMailDomicilio.Enabled = True
        cboPaisResidencia.Enabled = True
        cboPaisNacimiento.Enabled = True
        Call ColorControlHabilitado(txtApellidoPaterno)
        Call ColorControlHabilitado(txtApellidoMaterno)
        Call ColorControlHabilitado(txtNombres)
        Call ColorControlHabilitado(cboSexo)
        Call ColorControlHabilitado(cboEstadoCivil)
        Call ColorControlHabilitado(txtApellidoPaterno)
        Call ColorControlHabilitado(txtEMailDomicilio)
        Call ColorControlHabilitado(cboPaisResidencia)
        Call ColorControlHabilitado(cboPaisNacimiento)
        Call ColorControlFechaHabilitado(dtpFechaNacimiento)
        Call ColorControlFechaDeshabilitado(dtpFechaConstitucion)
        Call ColorControlHabilitado(txtCelularDomicilio)
        Call ColorControlDeshabilitado(txtPartidaElectronica)
        Call ColorControlDeshabilitado(txtObjetivoSocial)
        Call ColorControlDeshabilitado(txtWeb)
        Call ColorControlDeshabilitado(txtFaxDomicilio)
        txtRazonSocial.Enabled = False
        txtRazonSocial.Text = Valor_Caracter
        Call ColorControlDeshabilitado(txtRazonSocial)
        Call ColorControlDeshabilitado(txtOtroDocumento)
        txtOtroDocumento.Enabled = False
        txtOtroDocumento.Text = Valor_Caracter
        
        txtNombreEmpresa.Enabled = True
        optTipoTrabajador(0).Enabled = True
        optTipoTrabajador(1).Enabled = True
        txtRUCEmpresa.Enabled = True
        cboOcupacion.Enabled = True
        txtOcupCompletar.Enabled = True
        txtDireccionTrabajo1.Enabled = True
        txtDireccionTrabajo2.Enabled = True
        txtCargo.Enabled = True
        txtEMailTrabajo.Enabled = True
        txtTelefonoTrabajo.Enabled = True
        txtFaxTrabajo.Enabled = True
        txtCelularTrabajo.Enabled = True
        cboPaisTrabajo.Enabled = True
        cboDepartamentoTrabajo.Enabled = True
        cboProvinciaTrabajo.Enabled = True
        cboDistritoTrabajo.Enabled = True
        Call ColorControlHabilitado(txtNombreEmpresa)
        'Call ColorControlHabilitado(optTipoTrabajador)
        Call ColorControlHabilitado(txtRUCEmpresa)
        Call ColorControlHabilitado(cboOcupacion)
        Call ColorControlHabilitado(txtOcupCompletar)
        Call ColorControlHabilitado(txtDireccionTrabajo1)
        Call ColorControlHabilitado(txtDireccionTrabajo2)
        Call ColorControlHabilitado(txtCargo)
        Call ColorControlHabilitado(txtEMailTrabajo)
        Call ColorControlHabilitado(txtTelefonoTrabajo)
        Call ColorControlHabilitado(txtFaxTrabajo)
        Call ColorControlHabilitado(txtCelularTrabajo)
        Call ColorControlHabilitado(cboPaisTrabajo)
        Call ColorControlHabilitado(cboDepartamentoTrabajo)
        Call ColorControlHabilitado(cboProvinciaTrabajo)
        Call ColorControlHabilitado(cboDistritoTrabajo)
        
        tabCliente.TabVisible(4) = False
        tabCliente.TabVisible(5) = True
        tabCliente.TabVisible(6) = True
        
    Else
        txtApellidoPaterno.Enabled = False
        txtApellidoPaterno.Text = Valor_Caracter
        txtApellidoMaterno.Enabled = False
        txtApellidoMaterno.Text = Valor_Caracter
        txtNombres.Enabled = False
        txtNombres.Text = Valor_Caracter
        cboSexo.Enabled = False
        cboSexo.ListIndex = Valor_Numero
        cboEstadoCivil.Enabled = False
        cboEstadoCivil.ListIndex = Valor_Numero
        dtpFechaNacimiento.Enabled = False
        dtpFechaNacimiento.Value = Valor_Fecha
        dtpFechaConstitucion.Enabled = True
        txtObjetivoSocial.Enabled = True
        txtPartidaElectronica.Enabled = True
        txtWeb.Enabled = True
        txtFaxDomicilio.Enabled = True
        txtFaxDomicilio.Visible = True
        txtCelularDomicilio.Enabled = False
        txtCelularDomicilio.Visible = False
        lblDescrip(12).Caption = "Fax"
        txtEMailDomicilio.Enabled = False
        txtEMailDomicilio.Text = Valor_Caracter
        cboPaisResidencia.Enabled = False
        cboPaisResidencia.ListIndex = Valor_Numero
        cboPaisNacimiento.Enabled = False
        cboPaisNacimiento.ListIndex = Valor_Numero
        Call ColorControlDeshabilitado(txtApellidoPaterno)
        Call ColorControlDeshabilitado(txtApellidoMaterno)
        Call ColorControlDeshabilitado(txtNombres)
        Call ColorControlDeshabilitado(cboSexo)
        Call ColorControlDeshabilitado(cboEstadoCivil)
        Call ColorControlDeshabilitado(txtEMailDomicilio)
        Call ColorControlFechaDeshabilitado(dtpFechaNacimiento)
        Call ColorControlFechaHabilitado(dtpFechaConstitucion)
        Call ColorControlHabilitado(txtPartidaElectronica)
        Call ColorControlHabilitado(txtObjetivoSocial)
        Call ColorControlHabilitado(txtFaxDomicilio)
        Call ColorControlDeshabilitado(cboPaisResidencia)
        Call ColorControlDeshabilitado(cboPaisNacimiento)
        txtRazonSocial.Enabled = True
        Call ColorControlHabilitado(txtRazonSocial)
        Call ColorControlHabilitado(txtWeb)
        Call ColorControlDeshabilitado(txtOtroDocumento)
        Call ColorControlDeshabilitado(txtCelularDomicilio)
        txtOtroDocumento.Enabled = False
        txtOtroDocumento.Text = Valor_Caracter
        
        txtNombreEmpresa.Enabled = False
        txtNombreEmpresa.Text = Valor_Caracter
        optTipoTrabajador(0).Enabled = False
        optTipoTrabajador(1).Enabled = False
        txtRUCEmpresa.Enabled = False
        txtRUCEmpresa.Text = Valor_Caracter
        cboOcupacion.Enabled = False
        cboOcupacion.ListIndex = Valor_Numero
        txtOcupCompletar.Enabled = False
        txtOcupCompletar.Text = Valor_Caracter
        txtDireccionTrabajo1.Enabled = False
        txtDireccionTrabajo1.Text = Valor_Caracter
        txtDireccionTrabajo2.Enabled = False
        txtDireccionTrabajo2.Text = Valor_Caracter
        txtCargo.Enabled = False
        txtCargo.Text = Valor_Caracter
        txtEMailTrabajo.Enabled = False
        txtEMailTrabajo.Text = Valor_Caracter
        txtTelefonoTrabajo.Enabled = False
        txtTelefonoTrabajo.Text = Valor_Caracter
        txtFaxTrabajo.Enabled = False
        txtFaxTrabajo.Text = Valor_Caracter
        txtCelularTrabajo.Enabled = False
        txtCelularTrabajo.Text = Valor_Caracter
        cboPaisTrabajo.Enabled = False
        cboPaisTrabajo.ListIndex = Valor_Numero
        cboDepartamentoTrabajo.Enabled = False
        cboDepartamentoTrabajo.ListIndex = Valor_Numero
        cboProvinciaTrabajo.Enabled = False
        cboProvinciaTrabajo.ListIndex = Valor_Numero
        cboDistritoTrabajo.Enabled = False
        cboDistritoTrabajo.ListIndex = Valor_Numero
        
        Call ColorControlDeshabilitado(txtNombreEmpresa)
        'Call ColorControlDeshabilitado(optTipoTrabajador)
        Call ColorControlDeshabilitado(txtRUCEmpresa)
        Call ColorControlDeshabilitado(cboOcupacion)
        Call ColorControlDeshabilitado(txtOcupCompletar)
        Call ColorControlDeshabilitado(txtDireccionTrabajo1)
        Call ColorControlDeshabilitado(txtDireccionTrabajo2)
        Call ColorControlDeshabilitado(txtCargo)
        Call ColorControlDeshabilitado(txtEMailTrabajo)
        Call ColorControlDeshabilitado(txtTelefonoTrabajo)
        Call ColorControlDeshabilitado(txtFaxTrabajo)
        Call ColorControlDeshabilitado(txtCelularTrabajo)
        Call ColorControlDeshabilitado(cboPaisTrabajo)
        Call ColorControlDeshabilitado(cboDepartamentoTrabajo)
        Call ColorControlDeshabilitado(cboProvinciaTrabajo)
        Call ColorControlDeshabilitado(cboDistritoTrabajo)
        
        tabCliente.TabVisible(4) = True
        tabCliente.TabVisible(5) = False
        tabCliente.TabVisible(6) = False
        
    End If
    
    If strEstado = Reg_Adicion Then lblCodigoCliente.Caption = Format(strCodTipoDocumento & Trim(txtNumIdentidad.Text) & strCodClasePersona, "00000000000000000000")
    
    If strEstado = Reg_Adicion Then
            Call ConfiguraRecordsetAuxiliar
            Call ConfiguraRecordsetAuxiliarCtaCte
            TDBCtasCtes.DataSource = adoRegistroAuxCtaCte
            If strCodClasePersona = Codigo_Persona_Natural Then
                TDBDependientes.DataSource = adoRegistroAux
            End If
            If strCodClasePersona = Codigo_Persona_Juridica Then
                TDBRepresentantes.DataSource = adoRegistroAux
            End If
    Else
            Call CargarDetalleGrilla
            Call CargarDetalleGrillaCtaCte
    End If
End Sub

Private Sub cboDepartamento_Click()

    Dim strSql As String
    
    strCodDepartamento = Valor_Caracter
    If cboDepartamento.ListIndex < 0 Then Exit Sub
    
    strCodDepartamento = Trim(arrDepartamento(cboDepartamento.ListIndex))
    
    strSql = "{ call up_ACSelDatosParametro(2,'" & strCodPais & "','" & strCodDepartamento & "') }"
    CargarControlLista strSql, cboProvincia, arrProvincia(), Sel_Defecto
    
    If cboProvincia.ListCount > -1 Then cboProvincia.ListIndex = Valor_Numero
    
End Sub

Private Sub cboDepartamentoTrabajo_Click()

    Dim strSql As String
    
    strCodDepartamentoTrabajo = Valor_Caracter
    If cboDepartamentoTrabajo.ListIndex < 0 Then Exit Sub
    
    strCodDepartamentoTrabajo = Trim(arrDepartamentoTrabajo(cboDepartamentoTrabajo.ListIndex))
    
    strSql = "{ call up_ACSelDatosParametro(2,'" & strCodPaisTrabajo & "','" & strCodDepartamentoTrabajo & "') }"
    CargarControlLista strSql, cboProvinciaTrabajo, arrProvinciaTrabajo(), Sel_Defecto
    
    If cboProvinciaTrabajo.ListCount > -1 Then cboProvinciaTrabajo.ListIndex = Valor_Numero
    
End Sub

Private Sub cboDistrito_Click()
    
    strCodDistrito = Valor_Caracter
    If cboDistrito.ListIndex < 0 Then Exit Sub
    
    strCodDistrito = Trim(arrDistrito(cboDistrito.ListIndex))
        
End Sub

Private Sub cboDistritoTrabajo_Click()

    strCodDistritoTrabajo = Valor_Caracter
    If cboDistritoTrabajo.ListIndex < 0 Then Exit Sub
    
    strCodDistritoTrabajo = Trim(arrDistritoTrabajo(cboDistritoTrabajo.ListIndex))
    
End Sub

Private Sub cboEstadoCivil_Click()

    strCodEstadoCivil = Valor_Caracter
    If cboEstadoCivil.ListIndex < 0 Then Exit Sub
    
    strCodEstadoCivil = Trim(arrEstadoCivil(cboEstadoCivil.ListIndex))

End Sub

 

Private Sub cboMoneda_Click()

    strCodMoneda = Valor_Caracter
    If cboMoneda.ListIndex < 0 Then Exit Sub
    
    strCodMoneda = Trim(arrMoneda(cboMoneda.ListIndex))
    
End Sub

Private Sub cboNacionalidad_Click()

    strCodNacionalidad = Valor_Caracter
    If cboNacionalidad.ListIndex < 0 Then Exit Sub
    
    strCodNacionalidad = Trim(arrNacionalidad(cboNacionalidad.ListIndex))
        
End Sub

Private Sub cboPais_Click()

    Dim strSql As String
    
    strCodPais = Valor_Caracter
    If cboPais.ListIndex < 0 Then Exit Sub
    
    strCodPais = Trim(arrPais(cboPais.ListIndex))
    
    strSql = "{ call up_ACSelDatosParametro(1,'" & strCodPais & "') }"
    CargarControlLista strSql, cboDepartamento, arrDepartamento(), Sel_Defecto
    
    If cboDepartamento.ListCount > -1 Then cboDepartamento.ListIndex = Valor_Numero
    
End Sub

Private Sub cboPaisNacimiento_Click()
    strCodPaisNacimiento = Valor_Caracter
    If cboPaisNacimiento.ListIndex < 0 Then Exit Sub
    
    strCodPaisNacimiento = Trim(arrPaisNacimiento(cboPaisNacimiento.ListIndex))
End Sub

Private Sub cboPaisResidencia_Click()
    strCodPaisResidencia = Valor_Caracter
    If cboPaisResidencia.ListIndex < 0 Then Exit Sub
    
    strCodPaisResidencia = Trim(arrPaisResidencia(cboPaisResidencia.ListIndex))
End Sub

Private Sub cboPaisTrabajo_Click()

    Dim strSql As String
    
    strCodPaisTrabajo = Valor_Caracter
    If cboPaisTrabajo.ListIndex < 0 Then Exit Sub
    
    strCodPaisTrabajo = Trim(arrPaisTrabajo(cboPaisTrabajo.ListIndex))
    
    strSql = "{ call up_ACSelDatosParametro(1,'" & strCodPaisTrabajo & "') }"
    CargarControlLista strSql, cboDepartamentoTrabajo, arrDepartamentoTrabajo(), Sel_Defecto
    
    If cboDepartamentoTrabajo.ListCount > -1 Then cboDepartamentoTrabajo.ListIndex = Valor_Numero
    
End Sub

Private Sub cboProvincia_Click()
    
    Dim strSql As String
    
    strCodProvincia = Valor_Caracter
    If cboProvincia.ListIndex < 0 Then Exit Sub
    
    strCodProvincia = Trim(arrProvincia(cboProvincia.ListIndex))
    
    strSql = "{ call up_ACSelDatosParametro(3,'" & strCodPais & "','" & strCodDepartamento & "','" & strCodProvincia & "') }"
    CargarControlLista strSql, cboDistrito, arrDistrito(), Sel_Defecto
    
    If cboDistrito.ListCount > -1 Then cboDistrito.ListIndex = Valor_Numero
    
End Sub

Private Sub cboProvinciaTrabajo_Click()

    Dim strSql As String
    
    strCodProvinciaTrabajo = Valor_Caracter
    If cboProvinciaTrabajo.ListIndex < 0 Then Exit Sub
    
    strCodProvinciaTrabajo = Trim(arrProvinciaTrabajo(cboProvinciaTrabajo.ListIndex))
    
    strSql = "{ call up_ACSelDatosParametro(3,'" & strCodPaisTrabajo & "','" & strCodDepartamentoTrabajo & "','" & strCodProvinciaTrabajo & "') }"
    CargarControlLista strSql, cboDistritoTrabajo, arrDistritoTrabajo(), Sel_Defecto
    
    If cboDistritoTrabajo.ListCount > -1 Then cboDistritoTrabajo.ListIndex = Valor_Numero
    
End Sub

Private Sub cboSexo_Click()

    strCodSexo = Valor_Caracter
    If cboSexo.ListIndex < 0 Then Exit Sub
    
    strCodSexo = Trim(arrSexo(cboSexo.ListIndex))
    
End Sub

Private Sub cboTipoDocumento_Click()

    strCodTipoDocumento = Valor_Caracter
    If cboTipoDocumento.ListIndex < 0 Then Exit Sub
    
    strCodTipoDocumento = Trim(arrTipoDocumento(cboTipoDocumento.ListIndex))
    txtNumIdentidad.Text = Valor_Caracter
    txtNumIdentidad.MaxLength = ObtenerNumMaximoDocumentoIdentidad(strCodTipoDocumento)
    
    If strEstado = Reg_Adicion Then lblCodigoCliente.Caption = Format(strCodTipoDocumento & Trim(txtNumIdentidad.Text) & strCodClasePersona, "00000000000000000000")
    
    'If strCodTipoDocumento = "21" Then
    If (strCodTipoDocumento = Codigo_Tipo_Otro_Documento_Juridico Or strCodTipoDocumento = Codigo_Tipo_Otro_Documento_Natural) Then '20141128_JJCC
        Call ColorControlHabilitado(txtOtroDocumento)
        txtOtroDocumento.Enabled = True
        txtOtroDocumento.Text = Valor_Caracter
    Else
        Call ColorControlDeshabilitado(txtOtroDocumento)
        txtOtroDocumento.Enabled = False
        txtOtroDocumento.Text = Valor_Caracter
    End If
End Sub

Private Sub chkAdminRecPubNo_Click()
    
   If chkAdminRecPubNo.Value = 1 Then
    chkAdminRecPubSi.Value = Valor_Numero
    strRecursosPublicos = Valor_Caracter
   End If
   
End Sub

Private Sub chkAdminRecPubSi_Click()
    
   If chkAdminRecPubSi.Value = 1 Then
    chkAdminRecPubNo.Value = Valor_Numero
    strRecursosPublicos = Valor_Indicador
   End If
    
End Sub

Private Sub chkPepNo_Click()
    
    
   If chkPepNo.Value = 1 Then
   
    chkAdminRecPubSi.Value = Valor_Numero
    chkAdminRecPubNo.Value = Valor_Numero
    chkPepSi.Value = Valor_Numero
    strPEPS = Valor_Caracter
    txtInstitucionPEPS.Text = Valor_Caracter
    txtCargoDesemPEPS.Text = Valor_Caracter

    Call DesactivarPEPS
    
    End If
    
End Sub

Private Sub chkPepSi_Click()
    
    
    If chkPepSi.Value = 1 Then
    
    chkAdminRecPubNo.Value = 1
    chkPepNo.Value = Valor_Numero
    strPEPS = Valor_Indicador
    strRecursosPublicos = Valor_Caracter
    
    Call ActivarPEPS
    
    End If
End Sub


Private Sub cmdDefault_Click()

    Dim adoRegistro As ADODB.Recordset
    Dim intRegistro As Integer
    
    Set adoRegistro = New ADODB.Recordset
    
    '*** Cargar dirección de la administradora ***
    With adoComm
        .CommandText = "SELECT DescripDireccion1,DescripDireccion2,CodPais,CodDepartamento,CodProvincia,CodDistrito " & _
            "FROM Administradora WHERE CodAdministradora='" & gstrCodAdministradora & "'"
            
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            frmCliente.txtDireccionDomicilio1 = Trim(adoRegistro("DescripDireccion1"))
            frmCliente.txtDireccionDomicilio2 = Trim(adoRegistro("DescripDireccion2"))
            
            intRegistro = ObtenerItemLista(arrPais(), Trim(adoRegistro("CodPais")))
            If intRegistro >= Valor_Numero Then cboPais.ListIndex = intRegistro
            
            intRegistro = ObtenerItemLista(arrDepartamento(), Trim(adoRegistro("CodDepartamento")))
            If intRegistro >= Valor_Numero Then cboDepartamento.ListIndex = intRegistro
            
            intRegistro = ObtenerItemLista(arrProvincia(), Trim(adoRegistro("CodProvincia")))
            If intRegistro >= Valor_Numero Then cboProvincia.ListIndex = intRegistro
            
            intRegistro = ObtenerItemLista(arrDistrito(), Trim(adoRegistro("CodDistrito")))
            If intRegistro >= Valor_Numero Then cboDistrito.ListIndex = intRegistro
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With

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
    'Call DarFormato
    
    Call ValidarPermisoUsoControl(Trim(gstrLoginUS), Me, Trim(App.Title) + Separador_Codigo_Objeto + _
    gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)
    
    CentrarForm Me
    
End Sub

Private Function TodoOK() As Boolean

    TodoOK = False
    
    If cboTipoDocumento.ListIndex = Valor_Numero Then
        MsgBox "Seleccione el Tipo de Documento.", vbCritical
        tabCliente.Tab = 1
        cboTipoDocumento.SetFocus
        Exit Function
    End If
    
    '20141128_JJCC - Inicio
    If (strCodTipoDocumento = Codigo_Tipo_Otro_Documento_Juridico Or strCodTipoDocumento = Codigo_Tipo_Otro_Documento_Natural) _
    And txtOtroDocumento = Valor_Caracter Then
        MsgBox "Especifique de qué tipo es el Documento.", vbCritical
        tabCliente.Tab = 1
        txtOtroDocumento.SetFocus
    End If
    
    '20141128_JJCC - Fin
    
    If Trim(txtNumIdentidad.Text) = Valor_Caracter Then
        MsgBox "El Campo Número de Documento no es Válido!.", vbCritical
        tabCliente.Tab = 1
        txtNumIdentidad.SetFocus
        Exit Function
    End If
    
    If strCodClasePersona = Codigo_Persona_Natural Then
        If Trim(txtApellidoPaterno.Text) = Valor_Caracter Then
            MsgBox "El Campo Apellido Paterno no es Válido!.", vbCritical
            tabCliente.Tab = 1
            txtApellidoPaterno.SetFocus
            Exit Function
        End If
        
        If Trim(txtApellidoMaterno.Text) = Valor_Caracter Then
            If MsgBox("El Campo Apellido Materno no es Válido!." & vbNewLine & vbNewLine & _
                "Seguro de Continuar ?", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption) = vbNo Then
                tabCliente.Tab = 1
                txtApellidoMaterno.SetFocus
                Exit Function
            End If
        End If
        
        If Trim(txtNombres.Text) = Valor_Caracter Then
            MsgBox "El Campo Nombres no es Válido!.", vbCritical
            tabCliente.Tab = 1
            txtNombres.SetFocus
            Exit Function
        End If

        If cboEstadoCivil.ListIndex = Valor_Numero Then
            MsgBox "Seleccione el Estado Civil.", vbCritical
            tabCliente.Tab = 1
            cboEstadoCivil.SetFocus
            Exit Function
        End If
        
        If dtpFechaNacimiento.Value >= Now Then                       'HMC
            MsgBox "Seleccione una Fecha Valida.", vbCritical   '
            tabCliente.Tab = 1                                  '
            dtpFechaNacimiento.SetFocus                         '
            Exit Function                                       '
        End If                                                  '
        
        If cboPaisNacimiento.ListIndex = Valor_Numero Then
            MsgBox "Seleccione el País de Nacimiento", vbCritical
            cboPaisNacimiento.SetFocus
            Exit Function
        End If
        
        If cboSexo.ListIndex = Valor_Numero Then
            MsgBox "Seleccione el Sexo.", vbCritical
            tabCliente.Tab = 1
            cboSexo.SetFocus
            Exit Function
        End If

    Else
        If Trim(txtRazonSocial.Text) = Valor_Caracter Then
            MsgBox "El Campo Razón Social no es Válido!.", vbCritical
            tabCliente.Tab = 1
            txtRazonSocial.SetFocus
            Exit Function
        End If
    End If
    
    If cboNacionalidad.ListIndex = Valor_Numero Then
        MsgBox "Seleccione la Nacionalidad.", vbCritical
        tabCliente.Tab = 1
        cboNacionalidad.SetFocus
        Exit Function
    End If
    
    If Trim(txtDireccionDomicilio1.Text) = Valor_Caracter And Trim(txtDireccionDomicilio2.Text) = Valor_Caracter Then
        MsgBox "El Campo Dirección no es Válido!.", vbCritical
        txtDireccionDomicilio1.SetFocus
        Exit Function
    End If
    
    If cboPais.ListIndex = Valor_Numero Then
        MsgBox "Seleccione el País.", vbCritical
        cboPais.SetFocus
        Exit Function
    End If
    
    If cboDepartamento.ListIndex = Valor_Numero And cboDepartamento.ListCount > 1 Then
        MsgBox "Seleccione el Departamento.", vbCritical
        cboDepartamento.SetFocus
        Exit Function
    End If
    
    If cboProvincia.ListIndex = Valor_Numero And cboProvincia.ListCount > 1 Then
        MsgBox "Seleccione la Provincia.", vbCritical
        cboProvincia.SetFocus
        Exit Function
    End If
    
    If cboDistrito.ListIndex = Valor_Numero And cboDistrito.ListCount > 1 Then
        MsgBox "Seleccione el Distrito.", vbCritical
        cboDistrito.SetFocus
        Exit Function
    End If
    
    'If Trim(txtEMailDomicilio.Text) = Valor_Caracter And strCodClasePersona = Codigo_Persona_Natural Then
    '    MsgBox "El Campo E-Mail no es Válido!.", vbCritical
    '    txtEMailDomicilio.SetFocus
    '    Exit Function
    'End If
    
    If Trim(txtTelefonoDomicilio.Text) = Valor_Caracter Then
        MsgBox "El Campo Teléfono no es Válido!.", vbCritical
        txtTelefonoDomicilio.SetFocus
        Exit Function
    End If
                                                                    
    '*** Si todo pasó OK ***
    TodoOK = True

End Function

Private Sub DarFormato()

    Dim intCont As Integer
    Dim elemento As Object
    
    For intCont = Valor_Numero To (lblDescrip.Count - 1)
        Call FormatoEtiqueta(lblDescrip(intCont), vbLeftJustify)
    Next
    
    For intCont = Valor_Numero To (fraCliente.Count - 1)
        Call FormatoMarco(fraCliente(intCont))
    Next
    
    For Each elemento In Me.Controls
    
        If TypeOf elemento Is TDBGrid Then
            Call FormatoGrilla(elemento)
        End If
    
    Next
            
End Sub
Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Listado de Inversionistas por Fecha de Ingreso"

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Text = "Listado de Inversionistas Naturales"

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo3").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo3").Text = "Listado de Inversionistas Jurídicos"
    
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo4").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo4").Text = "Listado de Inversionistas PEPS"
    
    
End Sub

Private Sub CargarListas()

    Dim strSql  As String
    
    '*** Clase Persona ***
    strSql = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='CLSPER' and CodParametro<>'03' ORDER BY DescripParametro"
    CargarControlLista strSql, cboClasePersona, arrClasePersona(), ""
        
    '*** Sexo ***
    strSql = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='SEXPER' ORDER BY DescripParametro"
    CargarControlLista strSql, cboSexo, arrSexo(), Sel_Defecto
    
    '*** Estado Civil ***
    strSql = "{ call up_ACSelDatos(14) }"
    CargarControlLista strSql, cboEstadoCivil, arrEstadoCivil(), Sel_Defecto
    
    '*** Nacionalidad ***
    strSql = "{ call up_ACSelDatos(12) }"
    CargarControlLista strSql, cboNacionalidad, arrNacionalidad(), Sel_Defecto

    '*** Pais  ***
    strSql = "{ call up_ACSelDatos(13) }"
    CargarControlLista strSql, cboPais, arrPais(), Sel_Defecto
    CargarControlLista strSql, cboPaisTrabajo, arrPaisTrabajo(), Sel_Defecto
    CargarControlLista strSql, cboPaisNacimiento, arrPaisNacimiento(), Sel_Defecto
    CargarControlLista strSql, cboPaisResidencia, arrPaisResidencia(), Sel_Defecto

    '*** Ocupacion  ***
    strSql = "{ call up_ACSelDatos(40) }"
    CargarControlLista strSql, cboOcupacion, arrOcupacion(), Sel_Defecto
    If cboOcupacion.ListCount > 0 Then cboOcupacion.ListIndex = Valor_Numero
    
    '*** Vinculacion ***
    strSql = "SELECT CodVinculacion CODIGO,Descripcion DESCRIP from Vinculacion"
    CargarControlLista strSql, cboVinculacion, arrVinculacion(), Sel_Defecto
    If cboVinculacion.ListCount > 0 Then cboVinculacion.ListIndex = Valor_Numero
    CargarControlLista strSql, cboVinculacionLegales, arrVinculacionLegales(), Sel_Defecto
    If cboVinculacionLegales.ListCount > 0 Then cboVinculacionLegales.ListIndex = Valor_Numero
    
    '*** BANCOS ***
    strSql = "Select CodPersona CODIGO,DescripPersona DESCRIP from InstitucionPersona where IndBanco = 'X'"
    CargarControlLista strSql, cboBancos, arrBancos(), Sel_Defecto
    If cboBancos.ListCount > 0 Then cboBancos.ListIndex = Valor_Numero
    
    '*** Tipos Cuentas Bancarias ***
    strSql = "Select CodParametro CODIGO,DescripParametro DESCRIP from AuxiliarParametro where CodTipoParametro = 'CTAFON' and CodParametro <> '03'"
    CargarControlLista strSql, cboTipoCta, arrTipoCta(), Sel_Defecto
    If cboTipoCta.ListCount > 0 Then cboTipoCta.ListIndex = Valor_Numero
    
    '*** Monedas ***
    strSql = "SELECT CodMoneda CODIGO,DescripMoneda DESCRIP FROM Moneda " & _
        "WHERE Estado='" & Estado_Activo & "' " & _
        "ORDER BY DescripMoneda"
    CargarControlLista strSql, cboMoneda, arrMoneda(), Sel_Defecto
    
    If cboMoneda.ListCount > 0 Then cboMoneda.ListIndex = Valor_Numero
    
    '*** Envío de Información ***
    strSql = "Select CodParametro CODIGO,DescripParametro DESCRIP from AuxiliarParametro where CodTipoParametro = 'ENVINF' "
    CargarControlLista strSql, cboEnvioInformacion, arrEnvioInformacion(), Sel_Defecto
    If cboEnvioInformacion.ListCount > 0 Then cboEnvioInformacion.ListIndex = Valor_Numero
    
End Sub
Private Sub InicializarValores()

    '*** Valores Iniciales ***
    strEstado = Reg_Defecto
    tabCliente.Tab = Valor_Numero
    tabCliente.TabEnabled(1) = False
    tabCliente.TabEnabled(2) = False
    tabCliente.TabEnabled(3) = False
    tabCliente.TabEnabled(4) = False
    tabCliente.TabEnabled(5) = False
    tabCliente.TabEnabled(6) = False
    
    
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgConsulta.Columns(0).Width = tdgConsulta.Width * 0.01 * 18
    tdgConsulta.Columns(1).Width = tdgConsulta.Width * 0.01 * 16
    tdgConsulta.Columns(2).Width = tdgConsulta.Width * 0.01 * 11
    tdgConsulta.Columns(3).Width = tdgConsulta.Width * 0.01 * 40
    
    '*** Verificando Nivel de Acceso de Usuario ***
'    strNivAcceso = AccesoForm(gstrNomOpc, gstrNumInd)

    
    strPEPS = Valor_Caracter
    
    dtpPepFechaDesde = gdatFechaActual
    dtpPepFechaHasta = gdatFechaActual
        
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmCliente = Nothing
    Call OcultarReportes
    frmMainMdi.stbMdi.Panels(3).Text = "Acción..."
    
End Sub


Private Sub optCriterios_Click(index As Integer)

    If index = Valor_Numero Then
        txtNumDocumento.Enabled = True
        txtDescripCliente.Enabled = False
        txtNumDocumento.Text = Valor_Caracter
        txtNumDocumento.SetFocus
    Else
        txtNumDocumento.Enabled = False
        txtDescripCliente.Enabled = True
        txtDescripCliente.Text = Valor_Caracter
        txtDescripCliente.SetFocus
    End If
    
End Sub

Private Sub optTipoTrabajador_Click(index As Integer)
    Select Case index
    Case 0
        strCodTipoTrab = "D"
        'txtNombreEmpresa.Enabled = True
    Case 1
        strCodTipoTrab = "I"
        'txtNombreEmpresa.Enabled = False
    End Select
End Sub

Private Sub tabCliente_Click(PreviousTab As Integer)

    Select Case tabCliente.Tab
        Case 1, 2, 3, 4, 5, 6
            If PreviousTab = Valor_Numero And strEstado = Reg_Consulta Then Call Accion(vModify)
            If strEstado = Reg_Defecto Then tabCliente.Tab = Valor_Numero
        
    End Select
    
End Sub


Private Sub TDBDependientes_DblClick()
    Dim intRegistro As Integer
    If strEstado = Reg_Edicion Then
        If adoRegistroAux.RecordCount > 0 Then
            txtNombresLaboral.Text = adoRegistroAux.Fields("NombresApellidos")
            'cboTipoDocLaboral.ListIndex = Trim(adoRegistroAux.Fields("TipoDocumento"))
            intRegistro = ObtenerItemLista(arrTipoDocVinculado(), adoRegistroAux.Fields("TipoDocumento"))
            If intRegistro >= Valor_Numero Then cboTipoDocLaboral.ListIndex = intRegistro
            txtDocumentoLaboral.Text = adoRegistroAux.Fields("NroDocumento")
            'cboVinculacion.ListIndex = adoRegistroAux.Fields("Vinculacion")
            intRegistro = ObtenerItemLista(arrVinculacion(), adoRegistroAux.Fields("Vinculacion"))
            If intRegistro >= Valor_Numero Then cboVinculacion.ListIndex = intRegistro
        End If
    End If
    
End Sub

'Carga la data de la linea selec en los campos correspondientes
Private Sub TDBPEP_Click()
    Dim intRegistro As Integer
    
    If strEstado = Reg_Edicion Then


        'If TDBPEP.SelBookmarks.Count >= 1 Then
        
        If adoRegistroAuxPEP.RecordCount > 0 Then
        
            txtInstitucionPEPS.Text = TDBPEP.Columns("InstitucionPEPS").Value
            txtCargoDesemPEPS.Text = TDBPEP.Columns("CargoInstitucionPEPS").Value
            
            If TDBPEP.Columns("RecursosPublicosPEPS").Value = "Si" Then
                chkAdminRecPubSi.Value = 1
                chkAdminRecPubNo.Value = Valor_Numero
            End If
            
            If TDBPEP.Columns("RecursosPublicosPEPS").Value = "No" Then
                chkAdminRecPubSi.Value = Valor_Numero
                chkAdminRecPubNo.Value = 1
            End If
            
            dtpPepFechaDesde.Value = TDBPEP.Columns("FecDesdePEPS").Value
            dtpPepFechaHasta.Value = TDBPEP.Columns("FecHastaPEPS").Value
        
        End If
       ' End If
       
    End If
End Sub

Private Sub TDBRepresentantes_DblClick()
    Dim intRegistro As Integer
    If strEstado = Reg_Edicion Then
        If adoRegistroAux.RecordCount > 0 Then
            txtNombresLegales.Text = adoRegistroAux.Fields("Nombres")
            txtApellidosLegales.Text = adoRegistroAux.Fields("Apellidos")
            'cboTipoDocLegales.ListIndex = Trim(adoRegistroAux.Fields("TipoDocumento"))
            intRegistro = ObtenerItemLista(arrTipoDocumento(), Trim(adoRegistroAux.Fields("TipoDocumento")))
            If intRegistro >= Valor_Numero Then cboTipoDocLegales.ListIndex = intRegistro
            txtDocumentoLegales.Text = adoRegistroAux.Fields("NroDocumento")
            intRegistro = ObtenerItemLista(arrVinculacionLegales(), Trim(adoRegistroAux.Fields("VinculacionRep")))
            If intRegistro >= Valor_Numero Then cboVinculacionLegales.ListIndex = intRegistro
        End If
    End If
    
End Sub

Private Sub TDBCtasCtes_DblClick()

    Dim intRegistro As Integer
    If strEstado = Reg_Edicion Then
        If adoRegistroAuxCtaCte.RecordCount > 0 Then
            intRegistro = ObtenerItemLista(arrBancos(), Trim(adoRegistroAuxCtaCte.Fields("Banco")))
            If intRegistro >= Valor_Numero Then cboBancos.ListIndex = intRegistro
            intRegistro = ObtenerItemLista(arrTipoCta(), Trim(adoRegistroAuxCtaCte.Fields("TipoCtaCte")))
            If intRegistro >= Valor_Numero Then cboTipoCta.ListIndex = intRegistro
            'cboMoneda.ListIndex = Trim(adoRegistroAuxCtaCte.Fields("TipoMoneda"))
            intRegistro = ObtenerItemLista(arrMoneda(), Trim(adoRegistroAuxCtaCte.Fields("TipoMoneda")))
            If intRegistro >= Valor_Numero Then cboMoneda.ListIndex = intRegistro
            txtCCI.Text = adoRegistroAuxCtaCte.Fields("CCI")
            txtCtaCte.Text = adoRegistroAuxCtaCte.Fields("NroCtaCte")
        End If
    End If
    
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

Private Sub txtApellidoMaterno_KeyPress(KeyAscii As Integer)

    'KeyAscii = Asc(UCase(Chr(KeyAscii)))
    'Call ValidaCajaTexto(KeyAscii, "L", txtApellidoMaterno, 0)
 
    KeyAscii = ValiText(KeyAscii, "L", True)
 
End Sub

Private Sub txtApellidoPaterno_KeyPress(KeyAscii As Integer)

    'KeyAscii = Asc(UCase(Chr(KeyAscii)))
    'Call ValidaCajaTexto(KeyAscii, "L", txtApellidoPaterno, 0)
    
    KeyAscii = ValiText(KeyAscii, "L", True)
    
End Sub


Private Sub txtCargo_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
End Sub

Private Sub txtDescripCliente_KeyPress(KeyAscii As Integer)

    'KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
End Sub

Private Sub txtDireccionDomicilio1_KeyPress(KeyAscii As Integer)

    'KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
End Sub

Private Sub txtDireccionDomicilio2_KeyPress(KeyAscii As Integer)

    'KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
End Sub

Private Sub txtDireccionTrabajo1_KeyPress(KeyAscii As Integer)

    'KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
End Sub

Private Sub txtDireccionTrabajo2_KeyPress(KeyAscii As Integer)

    'KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
End Sub


Private Sub txtDocumentoLaboral_KeyPress(KeyAscii As Integer)

    KeyAscii = ValiText(KeyAscii, "N", True)

End Sub


Private Sub txtDocumentoLegales_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "N", txtDocumentoLegales, 0)
    
End Sub

Private Sub txtEMailDomicilio_KeyPress(KeyAscii As Integer)

    'KeyAscii = Asc(LCase(Chr(KeyAscii)))
    
End Sub

Private Sub txtEMailTrabajo_KeyPress(KeyAscii As Integer)

    'KeyAscii = Asc(LCase(Chr(KeyAscii)))
    
End Sub

Private Sub txtRUCEmpresa_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "N", txtRUCEmpresa, 0)
    
End Sub

Private Sub txtTelefonoTrabajo_KeyPress(KeyAscii As Integer)

    KeyAscii = ValiText(KeyAscii, "N", True)
    
End Sub

Private Sub txtFaxTrabajo_KeyPress(KeyAscii As Integer)

    KeyAscii = ValiText(KeyAscii, "N", True)
    
End Sub

Private Sub txtCelularTrabajo_KeyPress(KeyAscii As Integer)

    KeyAscii = ValiText(KeyAscii, "N", True)
    
End Sub

Private Sub txtCelularDomicilio_KeyPress(KeyAscii As Integer)

    KeyAscii = ValiText(KeyAscii, "N", True)
    
End Sub

Private Sub txtFaxDomicilio_KeyPress(KeyAscii As Integer)

    KeyAscii = ValiText(KeyAscii, "N", True)
    
End Sub

Private Sub txtNombreEmpresa_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
End Sub

Private Sub txtNombres_KeyPress(KeyAscii As Integer)
    
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'    Call ValidaCajaTexto(KeyAscii, "L", txtNombres, 0)

    KeyAscii = ValiText(KeyAscii, "L", True)

End Sub

Private Sub txtNumDocumento_KeyPress(KeyAscii As Integer)
        
    If KeyAscii >= 48 And KeyAscii <= 57 Then
      KeyAscii = KeyAscii
    ElseIf KeyAscii <> 8 Then
      KeyAscii = Valor_Numero
      Beep
   End If
   
End Sub

Private Sub txtNumIdentidad_KeyPress(KeyAscii As Integer)
    
    If strCodTipoDocumento = "07" Then
        Call ValidaCajaTexto(KeyAscii, "A", txtNumIdentidad, 0)
    Else
        Call ValidaCajaTexto(KeyAscii, "N", txtNumIdentidad, 0)
    End If
    
End Sub


Private Sub txtNumIdentidad_LostFocus()

On Error GoTo Error1            '/**/ HMC Habilitamos la rutina de Errores.

'    If strEstado = Reg_Adicion Then
    
        Dim adoRegistro As ADODB.Recordset
        
        lblCodigoCliente.Caption = Format(strCodTipoDocumento & Trim(txtNumIdentidad.Text) & strCodClasePersona, "00000000000000000000")
        
        Set adoRegistro = New ADODB.Recordset
        
        adoComm.CommandText = "{ call up_ACSelDatosParametro(7,'" & Trim(lblCodigoCliente.Caption) & "') }"
        Set adoRegistro = adoComm.Execute
        
        If Not adoRegistro.EOF Then
            MsgBox "Código de Inversionista ya existe", vbCritical, gstrNombreEmpresa
            txtNumIdentidad.Text = Valor_Caracter
            txtNumIdentidad.SetFocus
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
        
'    End If

On Error GoTo 0                  '/**/
Exit Sub                         '/**/
Error1:     MsgBox DescripcionError & vbNewLine & DescripcionTecnica & err.Description, vbExclamation, TituloError ' Mostrar Error
    
End Sub

'Private Sub txtOcupacion_KeyPress(KeyAscii As Integer)
'
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))

'End Sub


Private Sub txtRazonSocial_KeyPress(KeyAscii As Integer)
    
    'KeyAscii = Asc(UCase(Chr(KeyAscii)))
    'Call ValidaCajaTexto(KeyAscii, "L", txtRazonSocial, 0)
    KeyAscii = ValiText(KeyAscii, "AN", True)

End Sub

Private Sub txtTelefonoDomicilio_KeyPress(KeyAscii As Integer)

    'Call ValidaCajaTexto(KeyAscii, "N", txtNumIdentidad, 0)
    KeyAscii = ValiText(KeyAscii, "N", True)
    
End Sub

Private Function TodoOkCtaCte() As Boolean

    TodoOkCtaCte = False
        
        If cboBancos.ListIndex = Valor_Numero Then
            MsgBox "Banco no seleccionado", vbCritical, gstrNombreEmpresa
            cboBancos.SetFocus
            Exit Function
        End If
        
        If cboTipoCta.ListIndex = Valor_Numero Then
            MsgBox "Tipo Cuenta Bancaria no seleccionada", vbCritical, gstrNombreEmpresa
            cboTipoCta.SetFocus
            Exit Function
        End If
        
        If cboMoneda.ListIndex = Valor_Numero Then
            MsgBox "Tipo Moneda no seleccionada", vbCritical, gstrNombreEmpresa
            cboMoneda.SetFocus
            Exit Function
        End If
        
        If Trim(txtCCI.Text) = Valor_Caracter Then
            MsgBox "Nro CCI no ingresado", vbCritical, gstrNombreEmpresa
            txtCCI.SetFocus
            Exit Function
        End If
        
        If Trim(txtCtaCte.Text) = Valor_Caracter Then
            MsgBox "Nro Cuenta no ingresado", vbCritical, gstrNombreEmpresa
            txtCtaCte.SetFocus
            Exit Function
        End If
        
    TodoOkCtaCte = True
    
End Function

Private Function TodoOkListas() As Boolean

    TodoOkListas = False
    
    If strCodClasePersona = Codigo_Persona_Natural Then
        If Trim(txtNombresLaboral.Text) = Valor_Caracter Then
            MsgBox "Nombres y Apellidos no ingresados", vbCritical, gstrNombreEmpresa
            txtNombresLaboral.SetFocus
            Exit Function
        End If
        
        If cboTipoDocLaboral.ListIndex = Valor_Numero Then
            MsgBox "Tipo Documento de Identidad no seleccionada", vbCritical, gstrNombreEmpresa
            cboTipoDocLaboral.SetFocus
            Exit Function
        End If
        
        If Trim(txtDocumentoLaboral.Text) = Valor_Caracter Then
            MsgBox "Documento de Identidad no ingresado", vbCritical, gstrNombreEmpresa
            txtDocumentoLaboral.SetFocus
            Exit Function
        End If
        
        If cboVinculacion.ListIndex = Valor_Numero Then
            MsgBox "Vinculación no seleccionada", vbCritical, gstrNombreEmpresa
            cboVinculacion.SetFocus
            Exit Function
        End If
    End If
    
    If strCodClasePersona = Codigo_Persona_Juridica Then
        If Trim(txtNombresLegales.Text) = Valor_Caracter Then
            MsgBox "Nombres no ingresados", vbCritical, gstrNombreEmpresa
            txtNombresLaboral.SetFocus
            Exit Function
        End If
        If Trim(txtApellidosLegales.Text) = Valor_Caracter Then
            MsgBox "Apellidos no ingresados", vbCritical, gstrNombreEmpresa
            txtApellidosLegales.SetFocus
            Exit Function
        End If
        
        If cboTipoDocLegales.ListIndex = Valor_Numero Then
            MsgBox "Tipo Documento de Identidad no seleccionada", vbCritical, gstrNombreEmpresa
            cboTipoDocLaboral.SetFocus
            Exit Function
        End If
        
        If Trim(txtDocumentoLegales.Text) = Valor_Caracter Then
            MsgBox "Documento de Identidad no ingresado", vbCritical, gstrNombreEmpresa
            txtDocumentoLaboral.SetFocus
            Exit Function
        End If
        
        If cboVinculacionLegales.ListIndex = Valor_Numero Then
            MsgBox "Tipo Vinculación no seleccionada", vbCritical, gstrNombreEmpresa
            cboVinculacionLegales.SetFocus
            Exit Function
        End If
    End If
    
    TodoOkListas = True
  
End Function

'Validaciones para el formulario de PEPS
Private Function TodoOKPEP() As Boolean
    
    TodoOKPEP = False
    
    If Trim(txtInstitucionPEPS.Text) = Valor_Caracter Then
        MsgBox "Institucion no ingresada", vbCritical
        txtInstitucionPEPS.SetFocus
        Exit Function
    End If
    
    If Trim(txtCargoDesemPEPS.Text) = Valor_Caracter Then
        MsgBox "Cargo no ingresado", vbCritical
        txtCargoDesemPEPS.SetFocus
        Exit Function
    End If
    
    If (chkAdminRecPubNo.Value = Valor_Numero) And (chkAdminRecPubSi.Value = Valor_Numero) Then
        MsgBox "Especifique si Admnistra Recursos Publicos", vbCritical, gstrNombreEmpresa
        cboTipoDocLaboral.SetFocus
        Exit Function
    End If
    
    TodoOKPEP = True
    
End Function

Private Sub ConfiguraRecordsetAuxiliarCtaCte()

    Set adoRegistroAuxCtaCte = New ADODB.Recordset

        With adoRegistroAuxCtaCte
           .CursorLocation = adUseClient
           '.Fields.Append "CodCliente", adVarChar, 20
           .Fields.Append "Banco", adVarChar, 8
           .Fields.Append "BancoDesc", adVarChar, 100
           .Fields.Append "TipoCtaCte", adVarChar, 2
           .Fields.Append "TipoCtaCteDesc", adVarChar, 50
           .Fields.Append "TipoMoneda", adVarChar, 2
           .Fields.Append "TipoMonedaDesc", adVarChar, 30
           .Fields.Append "CCI", adVarChar, 50
           .Fields.Append "NroCtaCte", adVarChar, 30
           .LockType = adLockBatchOptimistic
        End With
 
    adoRegistroAuxCtaCte.Open

End Sub

Private Sub ConfiguraRecordsetAuxiliar()

    Set adoRegistroAux = New ADODB.Recordset
    
    If strCodClasePersona = Codigo_Persona_Natural Then
        With adoRegistroAux
           .CursorLocation = adUseClient
           '.Fields.Append "CodCliente", adVarChar, 20
           .Fields.Append "NombresApellidos", adVarChar, 150
           .Fields.Append "TipoDocumento", adVarChar, 2
           .Fields.Append "TipoDocDesc", adVarChar, 100
           .Fields.Append "NroDocumento", adVarChar, 15
           .Fields.Append "Vinculacion", adVarChar, 4
           .Fields.Append "VincDesc", adVarChar, 100
           .LockType = adLockBatchOptimistic
        End With
    adoRegistroAux.Open
    End If
    
    If strCodClasePersona = Codigo_Persona_Juridica Then
        With adoRegistroAux
           .CursorLocation = adUseClient
           .Fields.Append "CodRepresentante", adVarChar, 15
           '.Fields.Append "CodCliente", adVarChar, 20
           .Fields.Append "Nombres", adVarChar, 150
           .Fields.Append "Apellidos", adVarChar, 150
           .Fields.Append "TipoDocumento", adVarChar, 2
           .Fields.Append "TipoDocDesc", adVarChar, 100
           .Fields.Append "NroDocumento", adVarChar, 15
           .Fields.Append "VinculacionRep", adVarChar, 4
           .Fields.Append "VincDesc", adVarChar, 100
           .LockType = adLockBatchOptimistic
        End With
    adoRegistroAux.Open
    End If
 

End Sub

Private Sub ConfiguraRecordsetPep()
    
    Set adoRegistroAuxPEP = New ADODB.Recordset
    With adoRegistroAuxPEP
           .CursorLocation = adUseClient
           .Fields.Append "CodUnico", adVarChar, 20
           .Fields.Append "CodInstitucionPEPS", adInteger
           .Fields.Append "InstitucionPEPS", adVarChar, 250
           .Fields.Append "CargoInstitucionPEPS", adVarChar, 250
           .Fields.Append "RecursosPublicosPEPS", adVarChar, 2
           .Fields.Append "FecDesdePEPS", adDate
           .Fields.Append "FecHastaPEPS", adDate
           .LockType = adLockBatchOptimistic
        End With
    adoRegistroAuxPEP.Open
    
End Sub


Private Sub LimpiarDatosDep()

    txtNombresLaboral.Text = Valor_Caracter
    txtDocumentoLaboral.Text = Valor_Caracter
    cboTipoDocLaboral.ListIndex = Valor_Numero
    cboVinculacion.ListIndex = Valor_Numero
    
End Sub

Private Sub LimpiarDatosRep()

    txtNombresLegales.Text = Valor_Caracter
    txtApellidosLegales.Text = Valor_Caracter
    txtDocumentoLegales.Text = Valor_Caracter
    cboTipoDocLegales.ListIndex = Valor_Numero
    cboVinculacionLegales.ListIndex = Valor_Numero
    
End Sub

Private Sub LimpiarDatosCtaCte()

    txtCCI.Text = Valor_Caracter
    txtCtaCte.Text = Valor_Caracter
    cboMoneda.ListIndex = Valor_Numero
    cboBancos.ListIndex = Valor_Numero
    cboTipoCta.ListIndex = Valor_Numero
    
End Sub
 
'Limpia los campos de el Formulario PEP
Private Sub LimpiarDatosPEP()

    txtInstitucionPEPS.Text = Valor_Caracter
    txtCargoDesemPEPS.Text = Valor_Caracter
    chkAdminRecPubSi.Value = Valor_Numero
    chkAdminRecPubNo.Value = Valor_Numero
    dtpPepFechaDesde.Value = gdatFechaActual
    dtpPepFechaHasta.Value = gdatFechaActual

End Sub
Private Sub CargarDetallePEPS()

    Dim adoRegistroPEP As ADODB.Recordset
    Dim adoField As ADODB.Field
    
    Dim strSql As String
    
    Set adoRegistroPEP = New ADODB.Recordset
    
    Call ConfiguraRecordsetPep
    
     strSql = "SELECT CodUnico,CodInstitucionPEPS,InstitucionPEPS,CargoInstitucionPEPS, RecursosPublicosPEPS =  CASE RecursosPublicosPEPS WHEN 'X' THEN 'Si' ELSE 'No' END,FecDesdePEPS,FecHastaPEPS FROM ClientePEP WHERE CodUnico= '" & strCodCliente & "'"
        
        With adoRegistroPEP
            .ActiveConnection = gstrConnectConsulta
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockBatchOptimistic
            .Open strSql
        
            If .RecordCount > 0 Then
                .MoveFirst
                chkPepSi.Value = 1
                chkPepNo.Value = Valor_Numero
                Do While Not .EOF
                    adoRegistroAuxPEP.AddNew
                    For Each adoField In adoRegistroAuxPEP.Fields
                        adoRegistroAuxPEP.Fields(adoField.Name) = adoRegistroPEP.Fields(adoField.Name)
                    Next
                    adoRegistroAuxPEP.Update
                    adoRegistroPEP.MoveNext
                Loop
                adoRegistroAuxPEP.MoveFirst
            Else
                chkPepNo.Value = 1
            End If
            
        End With
         
        TDBPEP.DataSource = adoRegistroAuxPEP
        
        If adoRegistroAuxPEP.RecordCount > 0 Then
            adoRegistroAuxPEP.MoveLast
            intCodIndti = adoRegistroAuxPEP.Fields(1) 'busca el maximo CodInstitucionPEPS
            adoRegistroAuxPEP.MoveFirst
        End If

        
        
        'intCodIndti = adoRegistroAuxPEP.RecordCount
        btnPEPEliminar.Enabled = True
 
        
End Sub
 
Private Sub CargarDetalleGrillaCtaCte()
    
    Dim adoRegistroCtaCte As ADODB.Recordset
    Dim adoField As ADODB.Field
    
    Dim strSql As String
    
    Set adoRegistroCtaCte = New ADODB.Recordset
    
    Call ConfiguraRecordsetAuxiliarCtaCte
    
    'strSQL = "{ call up_ACSelDatosParametro(55,'" & strCodCliente & "') }"
    strSql = "SELECT CB.*,IP.DescripPersona as BancoDesc,AP.DescripParametro as TipoCtaCteDesc,M.DescripMoneda as TipoMonedaDesc FROM ClienteBancarios CB LEFT JOIN Moneda M ON M.CodMoneda = CB.TipoMoneda LEFT JOIN InstitucionPersona IP ON (IP.IndBanco = 'X' and CB.Banco = IP.CodPersona) LEFT JOIN AuxiliarParametro AP ON (AP.CodParametro = CB.TipoCtaCte and CodTipoParametro = 'CTAFON') WHERE CB.CodCliente='" & strCodCliente & "'"
        
        With adoRegistroCtaCte
            .ActiveConnection = gstrConnectConsulta
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockBatchOptimistic
            .Open strSql
        
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    adoRegistroAuxCtaCte.AddNew
                    For Each adoField In adoRegistroAuxCtaCte.Fields
                        adoRegistroAuxCtaCte.Fields(adoField.Name) = adoRegistroCtaCte.Fields(adoField.Name)
                    Next
                    adoRegistroAuxCtaCte.Update
                    adoRegistroCtaCte.MoveNext
                Loop
                adoRegistroAuxCtaCte.MoveFirst
            End If
            
        End With
     
    
        TDBCtasCtes.DataSource = adoRegistroAuxCtaCte
    
    'If adoRegistroAux.RecordCount > 0 Then strEstado = Reg_Consulta
            
End Sub

Private Sub CargarDetalleGrilla()
    
    Dim adoRegistro As ADODB.Recordset
    Dim adoField As ADODB.Field
    
    Dim strSql As String
    
    Set adoRegistro = New ADODB.Recordset
    
    Call ConfiguraRecordsetAuxiliar
    
    If strCodClasePersona = Codigo_Persona_Natural Then
        strSql = "Select d.*,t.DescripParametro as TipoDocDesc,v.Descripcion as VincDesc from ClienteDependientes d left join Vinculacion v on v.CodVinculacion = d.Vinculacion left join (Select * from AuxiliarParametro where CodTipoParametro = 'TIPIDE') t on t.CodParametro = d.TipoDocumento where d.CodCliente = '" & strCodCliente & "'"
    End If
    
    If strCodClasePersona = Codigo_Persona_Juridica Then
        strSql = "Select d.CodRepresentante,d.CodCliente,d.Nombres,d.Apellidos,d.TipoDocumento,d.NroDocumento,d.Vinculacion as VinculacionRep,v.Descripcion as VincDesc,t.DescripParametro as TipoDocDesc  from ClienteRepresentantes d  left join Vinculacion v on v.CodVinculacion = d.Vinculacion left join (Select * from AuxiliarParametro where CodTipoParametro = 'TIPIDE') t on t.CodParametro = d.TipoDocumento where d.CodCliente = '" & strCodCliente & "'"
    End If

        With adoRegistro
            .ActiveConnection = gstrConnectConsulta
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockBatchOptimistic
            .Open strSql
        
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    adoRegistroAux.AddNew
                    For Each adoField In adoRegistroAux.Fields
                        adoRegistroAux.Fields(adoField.Name) = adoRegistro.Fields(adoField.Name)
                    Next
                    adoRegistroAux.Update
                    adoRegistro.MoveNext
                Loop
                adoRegistroAux.MoveFirst
            End If
            
        End With
     
    
    
    If strCodClasePersona = Codigo_Persona_Natural Then
        TDBDependientes.DataSource = adoRegistroAux
    End If
    
    If strCodClasePersona = Codigo_Persona_Juridica Then
        TDBRepresentantes.DataSource = adoRegistroAux
    End If
    
    'If adoRegistroAux.RecordCount > 0 Then strEstado = Reg_Consulta
            
End Sub


Public Sub DesactivarPEPS()
    
    chkAdminRecPubSi.Enabled = False
    chkAdminRecPubNo.Enabled = False
    dtpPepFechaDesde.Enabled = False
    dtpPepFechaHasta.Enabled = False
    txtInstitucionPEPS.Enabled = False
    txtCargoDesemPEPS.Enabled = False
    btnPEPAgregar.Enabled = False
    btnPEPEditar.Enabled = False
    btnPEPEliminar.Enabled = False
            
End Sub

Public Sub ActivarPEPS()
    
    chkAdminRecPubSi.Enabled = True
    chkAdminRecPubNo.Enabled = True
    dtpPepFechaDesde.Enabled = True
    dtpPepFechaHasta.Enabled = True
    txtInstitucionPEPS.Enabled = True
    txtCargoDesemPEPS.Enabled = True
    btnPEPAgregar.Enabled = True
    btnPEPEditar.Enabled = True
    btnPEPEliminar.Enabled = True
    
End Sub
