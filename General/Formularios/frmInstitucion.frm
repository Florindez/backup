VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmInstitucion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Personas"
   ClientHeight    =   7650
   ClientLeft      =   1650
   ClientTop       =   1065
   ClientWidth     =   13035
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7650
   ScaleWidth      =   13035
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   11160
      TabIndex        =   2
      Top             =   6840
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
      TabIndex        =   1
      Top             =   6840
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
      Caption2        =   "&Eliminar"
      Tag2            =   "4"
      ToolTipText2    =   "Eliminar"
      UserControlWidth=   4200
   End
   Begin TabDlg.SSTab tabInstitucion 
      Height          =   6765
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Width           =   13005
      _ExtentX        =   22939
      _ExtentY        =   11933
      _Version        =   393216
      Style           =   1
      Tabs            =   4
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
      TabPicture(0)   =   "frmInstitucion.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblInstitucion(0)"
      Tab(0).Control(1)=   "cboTipoInstitucion"
      Tab(0).Control(2)=   "tdgConsulta"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Datos Básicos"
      TabPicture(1)   =   "frmInstitucion.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraDatos"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Datos Complementarios"
      TabPicture(2)   =   "frmInstitucion.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraDireccion"
      Tab(2).Control(1)=   "fraClasificacion"
      Tab(2).Control(2)=   "cmdAccion"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Calificación Instrumento"
      TabPicture(3)   =   "frmInstitucion.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraEmision"
      Tab(3).ControlCount=   1
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   -68040
         TabIndex        =   31
         Top             =   4920
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
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmInstitucion.frx":0070
         Height          =   4425
         Left            =   -74640
         OleObjectBlob   =   "frmInstitucion.frx":008A
         TabIndex        =   85
         Top             =   1200
         Width           =   12135
      End
      Begin VB.Frame fraEmision 
         Caption         =   "Calificación"
         Height          =   4545
         Left            =   -74760
         TabIndex        =   73
         Top             =   540
         Width           =   9495
         Begin TrueOleDBGrid60.TDBGrid tdgEmisor 
            Bindings        =   "frmInstitucion.frx":2EE6
            Height          =   2325
            Left            =   720
            OleObjectBlob   =   "frmInstitucion.frx":2EFE
            TabIndex        =   86
            Top             =   1920
            Width           =   8535
         End
         Begin VB.CommandButton cmdQuitar 
            Caption         =   "<"
            Height          =   375
            Left            =   240
            TabIndex        =   39
            Top             =   3120
            Width           =   375
         End
         Begin VB.ComboBox cboPlazo 
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
            Left            =   6360
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   360
            Width           =   2895
         End
         Begin VB.ComboBox cboClasificadora2 
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
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   1065
            Width           =   2895
         End
         Begin VB.ComboBox cboSubRiesgo2 
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
            Left            =   6360
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   1065
            Width           =   1695
         End
         Begin VB.ComboBox cboInstrumentoEmitido 
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
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   360
            Width           =   2895
         End
         Begin VB.ComboBox cboClasificadora1 
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
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   720
            Width           =   2895
         End
         Begin VB.ComboBox cboSubRiesgo1 
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
            Left            =   6360
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   720
            Width           =   1695
         End
         Begin VB.CommandButton cmdAgregar 
            Caption         =   ">"
            Height          =   375
            Left            =   240
            TabIndex        =   38
            Top             =   2460
            Width           =   375
         End
         Begin MSAdodcLib.Adodc adoEmisor 
            Height          =   330
            Left            =   2520
            Top             =   1560
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin VB.Label lblRiesgo 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7440
            TabIndex        =   84
            Top             =   1515
            Width           =   495
         End
         Begin VB.Label lblSubRiesgo 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8145
            TabIndex        =   83
            Top             =   1515
            Width           =   1095
         End
         Begin VB.Label lblInstitucion 
            Caption         =   "Clasificación Final"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   26
            Left            =   4920
            TabIndex        =   82
            Top             =   1545
            Width           =   1575
         End
         Begin VB.Label lblSubRiesgo2 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   8145
            TabIndex        =   81
            Top             =   1065
            Width           =   1095
         End
         Begin VB.Label lblInstitucion 
            Caption         =   "Clasificación II"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   25
            Left            =   4920
            TabIndex        =   80
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label lblInstitucion 
            Caption         =   "Clasificadora II"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   24
            Left            =   240
            TabIndex        =   79
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label lblInstitucion 
            Caption         =   "Categoría"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   23
            Left            =   4920
            TabIndex        =   78
            Top             =   375
            Width           =   1095
         End
         Begin VB.Label lblSubRiesgo1 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   8145
            TabIndex        =   77
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label lblInstitucion 
            Caption         =   "Clasificación I"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   22
            Left            =   4920
            TabIndex        =   76
            Top             =   735
            Width           =   1215
         End
         Begin VB.Label lblInstitucion 
            Caption         =   "Clasificadora I"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   21
            Left            =   240
            TabIndex        =   75
            Top             =   735
            Width           =   1335
         End
         Begin VB.Label lblInstitucion 
            Caption         =   "Instrumento"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   20
            Left            =   240
            TabIndex        =   74
            Top             =   375
            Width           =   1095
         End
      End
      Begin VB.Frame fraClasificacion 
         Caption         =   "Clasificación"
         Height          =   1815
         Left            =   -74760
         TabIndex        =   62
         Top             =   2790
         Width           =   9495
         Begin VB.ComboBox cboCategoria 
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
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   360
            Width           =   3015
         End
         Begin VB.ComboBox cboClasificadoraII 
            Appearance      =   0  'Flat
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
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   1185
            Width           =   3015
         End
         Begin VB.ComboBox cboClasificadoraI 
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
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   840
            Width           =   3015
         End
         Begin VB.ComboBox cboSubRiesgoI 
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
            Left            =   6360
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   840
            Width           =   1575
         End
         Begin VB.ComboBox cboSubRiesgoII 
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
            Left            =   6360
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   1185
            Width           =   1575
         End
         Begin VB.Label lblInstitucion 
            Caption         =   "Tipo"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   30
            Left            =   240
            TabIndex        =   72
            Top             =   380
            Width           =   615
         End
         Begin VB.Label lblSubRiesgoEntidad 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8085
            TabIndex        =   71
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lblRiesgoEntidad 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7420
            TabIndex        =   70
            Top             =   360
            Width           =   495
         End
         Begin VB.Label lblSubRiesgoII 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   8085
            TabIndex        =   69
            Top             =   1185
            Width           =   1095
         End
         Begin VB.Label lblSubRiesgoI 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   8085
            TabIndex        =   68
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label lblInstitucion 
            Caption         =   "Clasificadora I"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   15
            Left            =   240
            TabIndex        =   67
            Top             =   855
            Width           =   1335
         End
         Begin VB.Label lblInstitucion 
            Caption         =   "Clasificadora II"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   16
            Left            =   240
            TabIndex        =   66
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label lblInstitucion 
            Caption         =   "Clasificación I"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   17
            Left            =   4905
            TabIndex        =   65
            Top             =   855
            Width           =   1215
         End
         Begin VB.Label lblInstitucion 
            Caption         =   "Clasificación II"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   18
            Left            =   4905
            TabIndex        =   64
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label lblInstitucion 
            Caption         =   "Clasificación Final"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   19
            Left            =   4920
            TabIndex        =   63
            Top             =   380
            Width           =   1575
         End
      End
      Begin VB.Frame fraDireccion 
         Caption         =   "Dirección"
         Height          =   1935
         Left            =   -74760
         TabIndex        =   55
         Top             =   480
         Width           =   9495
         Begin VB.ComboBox cboDistrito 
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
            Left            =   6450
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   1340
            Width           =   2775
         End
         Begin VB.TextBox txtDireccion2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            MaxLength       =   50
            TabIndex        =   19
            Text            =   " "
            Top             =   645
            Width           =   4740
         End
         Begin VB.TextBox txtDireccion1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            MaxLength       =   50
            TabIndex        =   18
            Text            =   " "
            Top             =   360
            Width           =   4740
         End
         Begin VB.TextBox txtFax 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1200
            MaxLength       =   15
            TabIndex        =   21
            Top             =   1340
            Width           =   1815
         End
         Begin VB.TextBox txtTelefono 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1200
            MaxLength       =   15
            TabIndex        =   20
            Text            =   " "
            Top             =   1000
            Width           =   1815
         End
         Begin VB.ComboBox cboPais 
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
            Left            =   6450
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   360
            Width           =   2775
         End
         Begin VB.ComboBox cboDepartamento 
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
            Left            =   6450
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   680
            Width           =   2775
         End
         Begin VB.ComboBox cboProvincia 
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
            Left            =   6450
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   1000
            Width           =   2775
         End
         Begin VB.Label lblInstitucion 
            Caption         =   "Distrito"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   29
            Left            =   5160
            TabIndex        =   61
            Top             =   1365
            Width           =   855
         End
         Begin VB.Label lblInstitucion 
            Caption         =   "Teléfono"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   10
            Left            =   240
            TabIndex        =   60
            Top             =   1020
            Width           =   855
         End
         Begin VB.Label lblInstitucion 
            Caption         =   "Fax"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   11
            Left            =   240
            TabIndex        =   59
            Top             =   1360
            Width           =   735
         End
         Begin VB.Label lblInstitucion 
            Caption         =   "País"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   12
            Left            =   5160
            TabIndex        =   58
            Top             =   375
            Width           =   615
         End
         Begin VB.Label lblInstitucion 
            Caption         =   "Departamento"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   13
            Left            =   5160
            TabIndex        =   57
            Top             =   705
            Width           =   1215
         End
         Begin VB.Label lblInstitucion 
            Caption         =   "Provincia"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   14
            Left            =   5160
            TabIndex        =   56
            Top             =   1020
            Width           =   855
         End
      End
      Begin VB.Frame fraDatos 
         Height          =   6075
         Left            =   360
         TabIndex        =   42
         Top             =   480
         Width           =   12285
         Begin TAMControls2.TAMTextBox2 txtNemonico 
            Height          =   315
            Left            =   1740
            TabIndex        =   96
            Top             =   1750
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            MaxLength       =   4
            Container       =   "frmInstitucion.frx":93AC
            Apariencia      =   1
            Borde           =   1
            MaximoValor     =   0
         End
         Begin VB.ComboBox cboClasePersona 
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
            Left            =   1740
            Style           =   2  'Dropdown List
            TabIndex        =   93
            Top             =   850
            Width           =   2550
         End
         Begin VB.CommandButton cmdParametros 
            Caption         =   "Parámetros"
            Height          =   435
            Left            =   1740
            TabIndex        =   91
            Top             =   4900
            Width           =   1545
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
            Left            =   1740
            TabIndex        =   89
            Top             =   2650
            Width           =   2220
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
            Left            =   1740
            Style           =   2  'Dropdown List
            TabIndex        =   88
            Top             =   2200
            Width           =   6270
         End
         Begin VB.TextBox txtInstrumentosEmitidos 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   9915
            TabIndex        =   17
            Top             =   3690
            Width           =   1860
         End
         Begin VB.Frame fraRetencion 
            Caption         =   "Retención de cheques (días)"
            Height          =   1215
            Left            =   8340
            TabIndex        =   52
            Top             =   960
            Visible         =   0   'False
            Width           =   3495
            Begin VB.TextBox txtDiasBanco 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   2400
               TabIndex        =   12
               Top             =   360
               Width           =   615
            End
            Begin VB.TextBox txtDiasOtro 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   2400
               TabIndex        =   13
               Top             =   675
               Width           =   615
            End
            Begin VB.Label lblInstitucion 
               Caption         =   "Del mismo banco"
               Height          =   195
               Index           =   27
               Left            =   240
               TabIndex        =   54
               Top             =   360
               Width           =   1515
            End
            Begin VB.Label lblInstitucion 
               Caption         =   "De otro banco"
               Height          =   195
               Index           =   28
               Left            =   240
               TabIndex        =   53
               Top             =   720
               Width           =   1515
            End
         End
         Begin VB.TextBox txtRazonSocial 
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
            Left            =   1740
            MaxLength       =   200
            TabIndex        =   4
            Top             =   1300
            Width           =   6270
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
            TabIndex        =   5
            Top             =   5400
            Visible         =   0   'False
            Width           =   390
         End
         Begin VB.ComboBox cboExtranjero 
            Enabled         =   0   'False
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
            Left            =   1740
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   3100
            Width           =   6210
         End
         Begin VB.ComboBox cboCiiu 
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
            Left            =   1740
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   4000
            Width           =   6210
         End
         Begin VB.ComboBox cboGrupo 
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
            Left            =   1740
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   4450
            Width           =   6210
         End
         Begin VB.ComboBox cboSector 
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
            Left            =   1740
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   3550
            Width           =   6210
         End
         Begin VB.ComboBox cboTipoEntidad 
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
            Left            =   1740
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   400
            Width           =   4815
         End
         Begin VB.CheckBox chkExtranjero 
            Caption         =   "Extranjero"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   270
            TabIndex        =   6
            ToolTipText     =   "Marcar si es una entidad extranjera"
            Top             =   3150
            Width           =   1335
         End
         Begin VB.CheckBox chkBanco 
            Caption         =   "Entidad Bancaria"
            Height          =   255
            Left            =   8400
            TabIndex        =   11
            ToolTipText     =   "Marcar si el emisor es una entidad bancaria"
            Top             =   420
            Width           =   2055
         End
         Begin VB.TextBox txtCapital 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   9915
            TabIndex        =   14
            Top             =   2610
            Width           =   1860
         End
         Begin VB.TextBox txtObligaciones 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   9915
            TabIndex        =   15
            Top             =   2970
            Width           =   1860
         End
         Begin VB.TextBox txtLimiteComite 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   9915
            TabIndex        =   16
            Top             =   3330
            Width           =   1860
         End
         Begin VB.Label lblNemonico 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nemónico"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   33
            Left            =   270
            TabIndex        =   95
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label lblInstitucion 
            Caption         =   "Grupo"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   6
            Left            =   270
            TabIndex        =   94
            Top             =   4500
            Width           =   1095
         End
         Begin VB.Label lblInstitucion 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Clase Persona"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   36
            Left            =   270
            TabIndex        =   92
            Top             =   900
            Width           =   1455
         End
         Begin VB.Label lblInstitucion 
            Caption         =   "Nro.ID."
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   32
            Left            =   270
            TabIndex        =   90
            Top             =   2700
            Width           =   1215
         End
         Begin VB.Label lblInstitucion 
            Caption         =   "Instrumentos Financieros"
            ForeColor       =   &H00800000&
            Height          =   555
            Index           =   31
            Left            =   8310
            TabIndex        =   87
            Top             =   3630
            Width           =   1455
         End
         Begin VB.Label lblInstitucion 
            AutoSize        =   -1  'True
            Caption         =   "Patrimonio Neto"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   9
            Left            =   8310
            TabIndex        =   51
            Top             =   3345
            Width           =   1365
         End
         Begin VB.Label lblInstitucion 
            AutoSize        =   -1  'True
            Caption         =   "Pasivo Total"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   8
            Left            =   8310
            TabIndex        =   50
            Top             =   2985
            Width           =   1080
         End
         Begin VB.Label lblInstitucion 
            Caption         =   "Capital Social"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   7
            Left            =   8310
            TabIndex        =   49
            Top             =   2625
            Width           =   1215
         End
         Begin VB.Label lblInstitucion 
            Caption         =   "CIIU"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   5
            Left            =   270
            TabIndex        =   48
            Top             =   4050
            Width           =   975
         End
         Begin VB.Label lblInstitucion 
            Caption         =   "Sector"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   4
            Left            =   270
            TabIndex        =   47
            Top             =   3600
            Width           =   1215
         End
         Begin VB.Label lblInstitucion 
            Caption         =   "Tipo ID."
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   3
            Left            =   270
            TabIndex        =   46
            Top             =   2250
            Width           =   1215
         End
         Begin VB.Label lblInstitucion 
            Caption         =   "Razón Social"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   2
            Left            =   270
            TabIndex        =   45
            Top             =   1350
            Width           =   1455
         End
         Begin VB.Label lblInstitucion 
            Caption         =   "Tipo Institución"
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   1
            Left            =   270
            TabIndex        =   44
            Top             =   450
            Width           =   1455
         End
         Begin VB.Label lblcam 
            Alignment       =   2  'Center
            Caption         =   "Nuevos Soles"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   9630
            TabIndex        =   43
            Top             =   2310
            Width           =   2175
         End
      End
      Begin VB.ComboBox cboTipoInstitucion 
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
         Left            =   -72660
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   660
         Width           =   2700
      End
      Begin VB.Label lblInstitucion 
         Caption         =   "Tipo de Institución"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   -74520
         TabIndex        =   41
         Top             =   720
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmInstitucion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrTipoInstitucion()        As String, arrTipoEntidad()         As String
Dim arrRiesgoI()                As String, arrSubRiesgoI()          As String
Dim arrRiesgoII()               As String, arrSubRiesgoII()         As String
Dim arrRiesgo1()                As String, arrSubRiesgo1()          As String
Dim arrRiesgo2()                As String, arrSubRiesgo2()          As String
Dim arrSector()                 As String
Dim arrPais()                   As String, arrNacionalidad()        As String
Dim arrDepartamento()           As String, arrProvincia()           As String
Dim arrDistrito()               As String
Dim arrCiiu()                   As String, arrGrupo()               As String
Dim arrClasificadoraI()         As String, arrClasificadoraII()     As String
Dim arrClasificadora1()         As String, arrClasificadora2()      As String
Dim arrCategoria()              As String
Dim arrInstrumentoEmitido()     As String, arrPlazo()               As String
Dim arrTipoIdentidad()          As String, arrClasePersona()        As String

Dim intOrdenClasificadoraI      As Integer, intOrdenClasificadoraII As Integer
Dim strTipoInstitucion          As String, strTipoEntidad           As String
Dim strCodRiesgoI               As String, strCodSubRiesgoI         As String
Dim strCodRiesgoII              As String, strCodSubRiesgoII        As String
Dim strCodRiesgo1               As String, strCodSubRiesgo1         As String
Dim strCodRiesgo2               As String, strCodSubRiesgo2         As String
Dim strCodRiesgoFinal           As String, strCodSector             As String
Dim strCodPais                  As String, strCodNacionalidad       As String
Dim strCodDepartamento          As String, strCodProvincia          As String
Dim strCodDistrito              As String, strsql                   As String
Dim strCodCiiu                  As String, strCodGrupo              As String
Dim strCodClasificadoraI        As String, strCodClasificadoraII    As String
Dim strCodClasificadora1        As String, strCodClasificadora2     As String
Dim strCodCategoria             As String, strCodFile               As String
Dim strCodDetalleFile           As String, strCodPlazo              As String
Dim strEstado                   As String, strCodInstitucion        As String
Dim strTipoIdentidad            As String, strClasePersona          As String
Dim adoConsulta                 As ADODB.Recordset
Dim indSortAsc                  As Boolean, indSortDesc             As Boolean
Dim blnNemotecnicoUnico         As Boolean

Private Sub CargarTitulosEmitidos()

    Dim strsql As String
    Dim adoresultAux1 As ADODB.Recordset
                                                                                    
    Me.MousePointer = vbHourglass
            
    strsql = "SELECT EmisionInstitucionPersona.CodFile,EmisionInstitucionPersona.CodDetalleFile,CodCategoriaRiesgo,CodClasificadoraI,CodSubRiesgoI,CodClasificadoraII,CodSubRiesgoII,CodRiesgoFinal,CodSubRiesgoFinal,(DescripFile + ' ' + DescripDetalleFile) DescripTitulo,ValorParametro,(RTRIM(ValorParametro) + ' ' + RTRIM(CodSubRiesgoFinal)) DescripRiesgo " & _
        "FROM EmisionInstitucionPersona JOIN InversionDetalleFile " & _
        "ON (InversionDetalleFile.CodDetalleFile=EmisionInstitucionPersona.CodDetalleFile AND InversionDetalleFile.CodFile=EmisionInstitucionPersona.CodFile) " & _
        "JOIN InversionFile ON (InversionFile.CodFile=InversionDetalleFile.CodFile) " & _
        "JOIN AuxiliarParametro ON (AuxiliarParametro.CodParametro=EmisionInstitucionPersona.CodRiesgoFinal AND AuxiliarParametro.CodTipoParametro='TIPRIE') " & _
        "WHERE CodEmisor='" & strCodInstitucion & "'"
        
    With adoEmisor
        .ConnectionString = gstrConnectConsulta
        .RecordSource = strsql
        .Refresh
    End With
        
    tdgEmisor.Refresh
    
    Me.MousePointer = vbDefault
    
End Sub

Private Sub InicializarCalificacion()

    cboInstrumentoEmitido.ListIndex = -1
    If cboInstrumentoEmitido.ListCount > 0 Then cboInstrumentoEmitido.ListIndex = 0
    
    cboPlazo.ListIndex = -1
    If cboPlazo.ListCount > 0 Then cboPlazo.ListIndex = 0
    
    cboClasificadora1.ListIndex = -1
    If cboClasificadora1.ListCount > 0 Then cboClasificadora1.ListIndex = 0
    
    cboSubRiesgo1.ListIndex = -1
    If cboSubRiesgo1.ListCount > 0 Then cboSubRiesgo1.ListIndex = 0
    
    cboClasificadora2.ListIndex = -1
    If cboClasificadora2.ListCount > 0 Then cboClasificadora2.ListIndex = 0
    
    cboSubRiesgo2.ListIndex = -1
    If cboSubRiesgo2.ListCount > 0 Then cboSubRiesgo2.ListIndex = 0
    
    lblRiesgo.Caption = Valor_Caracter
    lblSubRiesgo.Caption = Valor_Caracter
    lblSubRiesgo1.Caption = Valor_Caracter
    lblSubRiesgo2.Caption = Valor_Caracter
    
End Sub

Private Sub LlenarFormulario(strModo As String)

    Dim adoRegistro   As ADODB.Recordset
    Dim strsql      As String
    Dim intRegistro As Integer
    
    Select Case strModo
        Case Reg_Adicion
            intRegistro = ObtenerItemLista(arrTipoEntidad(), strTipoInstitucion)
            If intRegistro >= 0 Then cboTipoEntidad.ListIndex = intRegistro
            cboTipoEntidad.Enabled = True
            
'            Set adoRegistro = New ADODB.Recordset
'
'            adoComm.CommandText = "SELECT COUNT(*) SecuencialInstitucion FROM InstitucionPersona WHERE TipoPersona='" & strTipoEntidad & "'"
'            Set adoRegistro = adoComm.Execute
'
'            If Not adoRegistro.EOF Then
'                strCodInstitucion = Format(adoRegistro("SecuencialInstitucion") + 1, "00000000")
'            Else
'                strCodInstitucion = "00000001"
'            End If
'            adoRegistro.Close: Set adoRegistro = Nothing
            
            txtRazonSocial.Text = Valor_Caracter
            txtRuc.Text = Valor_Caracter
                        
            chkBanco.Value = vbChecked
            chkBanco.Value = vbUnchecked
            chkExtranjero.Value = vbUnchecked
            If cboExtranjero.ListCount > 0 Then cboExtranjero.ListIndex = 0
            If cboSector.ListCount > 0 Then cboSector.ListIndex = 0
            If cboCiiu.ListCount > 0 Then cboCiiu.ListIndex = 0
            If cboGrupo.ListCount > 0 Then cboGrupo.ListIndex = 0
            If cboPais.ListCount > 0 Then cboPais.ListIndex = 0
            
            intRegistro = ObtenerItemLista(arrCategoria(), "03")
            If intRegistro >= 0 Then cboCategoria.ListIndex = intRegistro
            cboCategoria.Enabled = False
            
            If cboClasificadoraI.ListCount > 0 Then cboClasificadoraI.ListIndex = 0
            If cboClasificadoraII.ListCount > 0 Then cboClasificadoraII.ListIndex = 0
                        
            txtCapital.Text = "0"
            txtObligaciones.Text = "0"
            txtLimiteComite.Text = "0"
            txtInstrumentosEmitidos.Text = "0"
            txtDiasBanco.Text = "0"
            txtDiasOtro.Text = "0"
            fraRetencion.Visible = False
            txtDireccion1.Text = Valor_Caracter
            txtDireccion2.Text = Valor_Caracter
            txtTelefono.Text = Valor_Caracter
            txtFax.Text = Valor_Caracter
                                    
            tabInstitucion.TabEnabled(3) = False
            

                                    
        Case Reg_Edicion
            
            Set adoRegistro = New ADODB.Recordset
            
            intRegistro = ObtenerItemLista(arrTipoEntidad(), strTipoInstitucion)
            If intRegistro > -1 Then cboTipoEntidad.ListIndex = intRegistro
            cboTipoEntidad.Enabled = False
            
            strCodInstitucion = Trim(tdgConsulta.Columns(0).Value)
            
            adoComm.CommandText = "SELECT * FROM InstitucionPersona WHERE CodPersona='" & strCodInstitucion & "' AND TipoPersona='" & strTipoEntidad & "'"
            Set adoRegistro = adoComm.Execute
                                    
            If Not adoRegistro.EOF Then
                txtRazonSocial.Text = adoRegistro("RazonSocial")
                txtNemonico.Text = adoRegistro("DescripNemonico")
                txtNumIdentidad.Text = adoRegistro("NumIdentidad") 'adoRegistro("NumRuc")
                
                If Trim(adoRegistro("CodNacionalidad")) = Valor_Caracter Then
                    chkExtranjero.Value = vbUnchecked
                Else
                    chkExtranjero.Value = vbChecked
                    intRegistro = ObtenerItemLista(arrNacionalidad(), adoRegistro("CodNacionalidad"))
                    If intRegistro > -1 Then cboExtranjero.ListIndex = intRegistro
                End If
                
                intRegistro = ObtenerItemLista(arrSector(), adoRegistro("CodSector"))
                If intRegistro > -1 Then cboSector.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrCiiu(), adoRegistro("CodCiiu"))
                If intRegistro > -1 Then cboCiiu.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrGrupo(), adoRegistro("CodGrupo"))
                If intRegistro > -1 Then cboGrupo.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrTipoIdentidad(), adoRegistro("TipoIdentidad"))
                If intRegistro > -1 Then cboTipoIdentidad.ListIndex = intRegistro
                
                chkBanco.Value = vbChecked
                chkBanco.Value = vbUnchecked
                txtDiasBanco.Text = "0"
                txtDiasOtro.Text = "0"
                
                If Trim(adoRegistro("IndBanco")) = Valor_Caracter Then
                    chkBanco.Value = vbUnchecked
                Else
                    chkBanco.Value = vbChecked
                    txtDiasBanco.Text = CStr(adoRegistro("CantDiasRetencionBanco"))
                    txtDiasOtro.Text = CStr(adoRegistro("CantDiasRetencionOtroBanco"))
                End If
                
                txtCapital.Text = adoRegistro("MontoCapital")
                txtObligaciones.Text = adoRegistro("MontoObligacion")
                txtLimiteComite.Text = adoRegistro("MontoPatrimonioNeto")
                txtInstrumentosEmitidos.Text = adoRegistro("MontoPatrimonioNeto")
                
                txtDireccion1.Text = Trim(adoRegistro("Direccion1"))
                txtDireccion2.Text = Trim(adoRegistro("Direccion2"))
                txtTelefono.Text = Trim(adoRegistro("NumTelefono"))
                txtFax.Text = Trim(adoRegistro("NumFax"))
                
                intRegistro = ObtenerItemLista(arrPais(), adoRegistro("CodPais"))
                If intRegistro > -1 Then cboPais.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrDepartamento(), adoRegistro("CodDepartamento"))
                If intRegistro > -1 Then cboDepartamento.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrProvincia(), adoRegistro("CodProvincia"))
                If intRegistro > -1 Then cboProvincia.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrDistrito(), adoRegistro("CodDistrito"))
                If intRegistro > -1 Then cboDistrito.ListIndex = intRegistro
                                
                '*** Clasificación - Fortaleza Financiera ***
                intRegistro = ObtenerItemLista(arrCategoria(), "03")
                If intRegistro >= 0 Then cboCategoria.ListIndex = intRegistro
                cboCategoria.Enabled = False
                
                intRegistro = ObtenerItemLista(arrClasificadoraI(), adoRegistro("CodClasificadoraI"))
                If intRegistro > -1 Then cboClasificadoraI.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrClasificadoraII(), adoRegistro("CodClasificadoraII"))
                If intRegistro > -1 Then cboClasificadoraII.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrSubRiesgoI(), Trim(adoRegistro("CodRiesgoI")) + Trim(adoRegistro("CodSubRiesgoI")))
                If intRegistro > -1 Then cboSubRiesgoI.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrSubRiesgoII(), Trim(adoRegistro("CodRiesgoII")) + Trim(adoRegistro("CodSubRiesgoII")))
                If intRegistro > -1 Then cboSubRiesgoII.ListIndex = intRegistro
            
                '*** Clasificación Instrumentos ***
                If cboInstrumentoEmitido.ListCount > 0 Then cboInstrumentoEmitido.ListIndex = 0
                If cboClasificadora1.ListCount > 0 Then cboClasificadora1.ListIndex = 0
                If cboClasificadora2.ListCount > 0 Then cboClasificadora2.ListIndex = 0
                
                If strTipoEntidad = "02" Then
                    tabInstitucion.TabEnabled(3) = True
                    Call CargarTitulosEmitidos
                Else
                    tabInstitucion.TabEnabled(3) = False
                End If
            
            End If
            adoRegistro.Close: Set adoRegistro = Nothing
    
    End Select
    
End Sub

Private Sub InicializarValores()

    '*** Valores Iniciales ***
    strEstado = Reg_Defecto
    tabInstitucion.Tab = 0
    tabInstitucion.TabEnabled(1) = False
    tabInstitucion.TabEnabled(2) = False
    tabInstitucion.TabEnabled(3) = False
    
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgConsulta.Columns(0).Width = tdgConsulta.Width * 0.01 * 10
    tdgConsulta.Columns(1).Width = tdgConsulta.Width * 0.01 * 80
    
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgEmisor.Columns(0).Width = tdgEmisor.Width * 0.01 * 70
    tdgEmisor.Columns(1).Width = tdgEmisor.Width * 0.01 * 20
    
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    Set cmdOpcion.FormularioActivo = Me
    
End Sub


Public Sub Adicionar()
                    
    strEstado = Reg_Adicion
    LlenarFormulario strEstado
    cmdOpcion.Visible = False
    With tabInstitucion
        .TabEnabled(0) = False
        .TabEnabled(1) = True
        .TabEnabled(2) = True
        .TabEnabled(3) = True
        .Tab = 1
    End With
    Call Habilita
            
End Sub

Private Sub CargarListas()

    Dim strsql As String
    
    '*** Tipo Institución ***
    strsql = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPPER' AND ValorParametro='X' ORDER BY DescripParametro"
    CargarControlLista strsql, cboTipoInstitucion, arrTipoInstitucion(), ""
    CargarControlLista strsql, cboTipoEntidad, arrTipoEntidad(), ""
    If cboTipoInstitucion.ListCount > 0 Then cboTipoInstitucion.ListIndex = 0
    
    '*** País ***
    strsql = "{ call up_ACSelDatos(13) }"
    CargarControlLista strsql, cboPais, arrPais(), Sel_Defecto

    '*** Nacionalidad ***
    strsql = "{ call up_ACSelDatos(12) }"
    CargarControlLista strsql, cboExtranjero, arrNacionalidad(), Sel_Defecto
    
    '*** Tipos de documentos juridico ***
    'strSQL = "{ call up_ACSelDatos(27) }"
    'CargarControlLista strSQL, cboTipoIdentidad, arrTipoIdentidad(), Sel_Defecto
    
    '*** Sectores ***
    strsql = "SELECT CodSector CODIGO,DescripSector DESCRIP FROM SectorBursatil ORDER BY DescripSector"
    CargarControlLista strsql, cboSector, arrSector(), Sel_Defecto
    
    '*** Ciiu ***
    strsql = "SELECT CodCiiu CODIGO ,(CodCiiu + Space(1) + DescripCiiu) DESCRIP FROM Ciiu WHERE IndVigente='X'ORDER BY CodCiiu"
    CargarControlLista strsql, cboCiiu, arrCiiu(), Sel_Defecto

    '*** Grupo Económico ***
    strsql = "SELECT CodGrupo CODIGO ,DescripGrupo DESCRIP FROM GrupoEconomico WHERE IndVigente='X' ORDER BY DescripGrupo"
    CargarControlLista strsql, cboGrupo, arrGrupo(), Sel_Defecto

    '*** Categoría Clasificación Riesgo ***
    strsql = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='CATRIE' ORDER BY DescripParametro"
    CargarControlLista strsql, cboCategoria, arrCategoria(), ""
    If cboCategoria.ListCount > 0 Then cboCategoria.ListIndex = 0
    
    strsql = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='CATRIE' AND CodParametro <> '03' ORDER BY DescripParametro"
    CargarControlLista strsql, cboPlazo, arrPlazo(), ""
    If cboPlazo.ListCount > 0 Then cboPlazo.ListIndex = 0
                              
    '*** Empresas Clasificadoras ***
    strsql = "SELECT CodClasificadora CODIGO,DescripClasificadora DESCRIP FROM Clasificadora WHERE CodClasificadora<>'001' ORDER BY DescripClasificadora"
    CargarControlLista strsql, cboClasificadoraI, arrClasificadoraI(), Sel_Defecto
    CargarControlLista strsql, cboClasificadoraII, arrClasificadoraII(), Sel_Defecto
    CargarControlLista strsql, cboClasificadora1, arrClasificadora1(), Sel_Defecto
    CargarControlLista strsql, cboClasificadora2, arrClasificadora2(), Sel_Defecto
    
    '*** Títulos Emitidos ***
    strsql = "SELECT (InversionDetalleFile.CodFile + CodDetalleFile) CODIGO,(RTRIM(DescripFile) + ' ' + DescripDetalleFile) DESCRIP "
    strsql = strsql & "FROM InversionDetalleFile JOIN InversionFile ON(InversionFile.CodFile=InversionDetalleFile.CodFile) WHERE IndEmision='X' ORDER BY DescripFile"
    CargarControlLista strsql, cboInstrumentoEmitido, arrInstrumentoEmitido(), Sel_Defecto
    
    '*** Clase de Persona ***
    strsql = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='CLSPER' ORDER BY DescripParametro"
    CargarControlLista strsql, cboClasePersona, arrClasePersona(), ""
    If cboClasePersona.ListCount > Valor_Numero Then cboClasePersona.ListIndex = Valor_Numero
        
End Sub


Public Sub Buscar()

    Dim adoresultAux1 As ADODB.Recordset
    Set adoConsulta = New ADODB.Recordset
                                                                                    
    Me.MousePointer = vbHourglass
            
    strsql = "SELECT CodPersona,DescripPersona,(CodRiesgo + ' ' + CodSubRiesgo) CodRiesgo FROM InstitucionPersona WHERE TipoPersona='" & strTipoInstitucion & "' AND IndVigente='X' ORDER BY DescripPersona"
        
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

Public Sub Cancelar()

    cmdOpcion.Visible = True
    With tabInstitucion
        .TabEnabled(0) = True
        .TabEnabled(1) = False
        .TabEnabled(2) = False
        .TabEnabled(3) = False
        .Tab = 0
    End With
    Call Buscar
    
End Sub

Private Sub Deshabilita()

    fraDatos.Enabled = False
    fraDireccion.Enabled = False
    fraClasificacion.Enabled = False
    fraEmision.Enabled = False
    
End Sub

Public Sub Eliminar()

    Dim adoresult   As ADODB.Recordset
    Dim sqldel      As String
    
    If strEstado = "CONSULTA" Or strEstado = "EDICION" Then

        If MsgBox("Se procederá a cambiar el estado de la Institución a NO VIGENTE." & vbNewLine & vbNewLine & "ESTE PROCESO ES IRREVERSIBLE." & vbNewLine & vbNewLine & vbNewLine & "Seguro de continuar ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then

            adoComm.CommandText = "UPDATE InstitucionPersona SET IndVigente='' WHERE CodPersona='" & tdgConsulta.Columns(0).Value & "'"
            adoComm.Execute
            tabInstitucion.Tab = 0
            Call Buscar

        End If
    End If

End Sub

Public Sub Grabar()
    
    Dim strClasePersona     As String
    Dim intAccion           As Integer, lngNumError     As Long
    Dim intRegistro         As Integer
    Dim adoRegistro         As ADODB.Recordset
    Dim adoError            As ADODB.Error
    Dim strErrMsg           As String
                
                
    If strEstado = Reg_Consulta Then Exit Sub
    
    On Error GoTo CtrlError
    
    If strEstado = Reg_Adicion Then
        If TodoOK() Then
            Me.MousePointer = vbHourglass
            
            'strClasePersona = "02" 'Por ahora son solo PJ
            
            Set adoRegistro = New ADODB.Recordset
            
            adoComm.CommandText = "SELECT CONVERT(int,MAX(CodPersona)) as SecuencialInstitucion FROM InstitucionPersona WHERE TipoPersona='" & strTipoEntidad & "'"
            Set adoRegistro = adoComm.Execute
            
            If Not adoRegistro.EOF And Not IsNull(adoRegistro("SecuencialInstitucion")) Then
                strCodInstitucion = Format(adoRegistro("SecuencialInstitucion") + 1, "00000000")
            Else
                strCodInstitucion = "00000001"
            End If
            adoRegistro.Close: Set adoRegistro = Nothing
            
            '*** Guardar Entidad ***
            With adoComm
                .CommandText = "{ call up_GNManInstitucionPersona('"
                .CommandText = .CommandText & strCodInstitucion & "','"
                .CommandText = .CommandText & strTipoEntidad & "','"
                .CommandText = .CommandText & strClasePersona & "','"
                .CommandText = .CommandText & strTipoIdentidad & "','"
                .CommandText = .CommandText & Trim(txtNumIdentidad.Text) & "','"
                .CommandText = .CommandText & "','"
                .CommandText = .CommandText & "','"
                .CommandText = .CommandText & "','"
                .CommandText = .CommandText & Trim(txtRazonSocial.Text) & "','"
                .CommandText = .CommandText & Trim(txtRazonSocial.Text) & "','"
                .CommandText = .CommandText & Trim(txtNemonico.Text) & "','"
                .CommandText = .CommandText & Trim(txtDireccion1.Text) & "','"
                .CommandText = .CommandText & Trim(txtDireccion2.Text) & "','"
                .CommandText = .CommandText & "','"
                .CommandText = .CommandText & "','"
                .CommandText = .CommandText & "','"
                .CommandText = .CommandText & strCodPais & "','"
                .CommandText = .CommandText & strCodDepartamento & "','"
                .CommandText = .CommandText & strCodProvincia & "','"
                .CommandText = .CommandText & strCodDistrito & "','"
                .CommandText = .CommandText & strCodNacionalidad & "','"
                .CommandText = .CommandText & Trim(txtTelefono.Text) & "','"
                .CommandText = .CommandText & Trim(txtFax.Text) & "','"
                .CommandText = .CommandText & Trim(txtRuc.Text) & "','"
                .CommandText = .CommandText & strCodGrupo & "','"
                .CommandText = .CommandText & strCodCiiu & "','"
                .CommandText = .CommandText & strCodSector & "',"
                .CommandText = .CommandText & CCur(txtObligaciones.Text) & ","
                .CommandText = .CommandText & CCur(txtCapital.Text) & ","
                .CommandText = .CommandText & CCur(txtLimiteComite.Text) & ","
                .CommandText = .CommandText & CCur(txtInstrumentosEmitidos.Text) & ",'"
                .CommandText = .CommandText & "','"
                .CommandText = .CommandText & "','"
                .CommandText = .CommandText & Convertyyyymmdd(Date) & "','"
                .CommandText = .CommandText & Convertyyyymmdd(Date) & "','"
                .CommandText = .CommandText & "','"
                .CommandText = .CommandText & "','"
                If chkBanco.Value Then
                    .CommandText = .CommandText & "X',"
                    .CommandText = .CommandText & CInt(txtDiasBanco.Text) & ","
                    .CommandText = .CommandText & CInt(txtDiasOtro.Text) & ",'"
                Else
                    .CommandText = .CommandText & "',"
                    .CommandText = .CommandText & "0,"
                    .CommandText = .CommandText & "0,'"
                End If
                .CommandText = .CommandText & "','"
                .CommandText = .CommandText & "','"
                .CommandText = .CommandText & "','"
                .CommandText = .CommandText & "X','"
                .CommandText = .CommandText & strCodCategoria & "','"
                .CommandText = .CommandText & strCodClasificadoraI & "','"
                .CommandText = .CommandText & Trim(strCodRiesgoI) & "','"
                .CommandText = .CommandText & Trim(strCodSubRiesgoI) & "','"
                .CommandText = .CommandText & strCodClasificadoraII & "','"
                .CommandText = .CommandText & Trim(strCodRiesgoII) & "','"
                .CommandText = .CommandText & Trim(strCodSubRiesgoII) & "','"
                .CommandText = .CommandText & Trim(lblRiesgoEntidad.Caption) & "','"
                .CommandText = .CommandText & Trim(lblSubRiesgoEntidad.Caption) & "','"
                .CommandText = .CommandText & gstrLogin & "','"
                .CommandText = .CommandText & Convertyyyymmdd(Date) & "','"
                .CommandText = .CommandText & gstrLogin & "','"
                .CommandText = .CommandText & Convertyyyymmdd(Date) & "','"
                .CommandText = .CommandText & "I') }"
                adoConn.Execute .CommandText
                                            
            End With
            
            Me.MousePointer = vbDefault
            MsgBox Mensaje_Adicion_Exitosa, vbExclamation
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"

            cmdOpcion.Visible = True
            With tabInstitucion
                .TabEnabled(0) = True
                .Tab = 0
            End With
            Call Buscar
        End If
    End If
    
    If strEstado = Reg_Edicion Then
        If TodoOK() Then
            Me.MousePointer = vbHourglass
            
            'strTipoIdentidad = "06"  'RUC
            'strClasePersona = "02"
            
            '*** Guardar Entidad ***
            With adoComm
                .CommandText = "{ call up_GNManInstitucionPersona('"
                .CommandText = .CommandText & strCodInstitucion & "','"
                .CommandText = .CommandText & strTipoEntidad & "','"
                .CommandText = .CommandText & strClasePersona & "','"
                .CommandText = .CommandText & strTipoIdentidad & "','"
                .CommandText = .CommandText & Trim(txtNumIdentidad.Text) & "','"
                .CommandText = .CommandText & "','"
                .CommandText = .CommandText & "','"
                .CommandText = .CommandText & "','"
                .CommandText = .CommandText & Trim(txtRazonSocial.Text) & "','"
                .CommandText = .CommandText & Trim(txtRazonSocial.Text) & "','"
                .CommandText = .CommandText & Trim(txtNemonico.Text) & "','"
                .CommandText = .CommandText & Trim(txtDireccion1.Text) & "','"
                .CommandText = .CommandText & Trim(txtDireccion2.Text) & "','"
                .CommandText = .CommandText & "','"
                .CommandText = .CommandText & "','"
                .CommandText = .CommandText & "','"
                .CommandText = .CommandText & strCodPais & "','"
                .CommandText = .CommandText & strCodDepartamento & "','"
                .CommandText = .CommandText & strCodProvincia & "','"
                .CommandText = .CommandText & strCodDistrito & "','"
                .CommandText = .CommandText & strCodNacionalidad & "','"
                .CommandText = .CommandText & Trim(txtTelefono.Text) & "','"
                .CommandText = .CommandText & Trim(txtFax.Text) & "','"
                .CommandText = .CommandText & Trim(txtRuc.Text) & "','"
                .CommandText = .CommandText & strCodGrupo & "','"
                .CommandText = .CommandText & strCodCiiu & "','"
                .CommandText = .CommandText & strCodSector & "',"
                .CommandText = .CommandText & CCur(txtObligaciones.Text) & ","
                .CommandText = .CommandText & CCur(txtCapital.Text) & ","
                .CommandText = .CommandText & CCur(txtLimiteComite.Text) & ","
                .CommandText = .CommandText & CCur(txtInstrumentosEmitidos.Text) & ",'"
                .CommandText = .CommandText & "','"
                .CommandText = .CommandText & "','"
                .CommandText = .CommandText & Convertyyyymmdd(Date) & "','"
                .CommandText = .CommandText & Convertyyyymmdd(Date) & "','"
                .CommandText = .CommandText & "','"
                .CommandText = .CommandText & "','"
                If chkBanco.Value Then
                    .CommandText = .CommandText & "X',"
                    .CommandText = .CommandText & CInt(txtDiasBanco.Text) & ","
                    .CommandText = .CommandText & CInt(txtDiasOtro.Text) & ",'"
                Else
                    .CommandText = .CommandText & "',"
                    .CommandText = .CommandText & "0,"
                    .CommandText = .CommandText & "0,'"
                End If
                .CommandText = .CommandText & "','"
                .CommandText = .CommandText & "','"
                .CommandText = .CommandText & "','"
                .CommandText = .CommandText & "X','"
                .CommandText = .CommandText & strCodCategoria & "','"
                .CommandText = .CommandText & strCodClasificadoraI & "','"
                .CommandText = .CommandText & Trim(strCodRiesgoI) & "','"
                .CommandText = .CommandText & Trim(strCodSubRiesgoI) & "','"
                .CommandText = .CommandText & strCodClasificadoraII & "','"
                .CommandText = .CommandText & Trim(strCodRiesgoII) & "','"
                .CommandText = .CommandText & Trim(strCodSubRiesgoII) & "','"
                .CommandText = .CommandText & Trim(lblRiesgoEntidad.Caption) & "','"
                .CommandText = .CommandText & Trim(lblSubRiesgoEntidad.Caption) & "','"
                .CommandText = .CommandText & gstrLogin & "','"
                .CommandText = .CommandText & Convertyyyymmdd(Date) & "','"
                .CommandText = .CommandText & gstrLogin & "','"
                .CommandText = .CommandText & Convertyyyymmdd(Date) & "','"
                .CommandText = .CommandText & "U') }"
                adoConn.Execute .CommandText
                                            
            End With
            
            Me.MousePointer = vbDefault
            
            MsgBox Mensaje_Edicion_Exitosa, vbExclamation
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
            cmdOpcion.Visible = True
            With tabInstitucion
                .TabEnabled(0) = True
                .Tab = 0
            End With
            Call Buscar
        End If
    End If
    Exit Sub
                    
CtrlError:
    If adoConn.Errors.Count > 0 Then
        For Each adoError In adoConn.Errors
            strErrMsg = strErrMsg & adoError.Description & " (" & adoError.NativeError & ") "
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

Private Sub Habilita()

    fraDatos.Enabled = True
    fraDireccion.Enabled = True
    fraClasificacion.Enabled = True
    fraEmision.Enabled = True
    
End Sub


Public Sub Imprimir()
    
End Sub

Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Lista de Instituciones"
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    
End Sub


Public Sub Modificar()
    
    If strEstado = Reg_Consulta Then
        strEstado = Reg_Edicion
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabInstitucion
            .TabEnabled(0) = False
            .TabEnabled(1) = True
            .TabEnabled(2) = True
            .TabEnabled(3) = True
            .Tab = 1
        End With
        Call Habilita
    End If
        
End Sub



Public Sub Salir()

    Unload Me
    
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










Public Sub SubImprimir(Index As Integer)

    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()

    Select Case Index
        Case 1
            gstrNameRepo = "InstitucionPersona"
                        
            Set frmReporte = New frmVisorReporte

            ReDim aReportParamS(2)
            ReDim aReportParamFn(3)
            ReDim aReportParamF(3)

            aReportParamFn(0) = "Usuario"
            aReportParamFn(1) = "Hora"
            aReportParamFn(2) = "TituloReporte"
            aReportParamFn(3) = "NombreEmpresa"
            
            aReportParamF(0) = gstrLogin
            aReportParamF(1) = Format(Time(), "hh:mm:ss")
            aReportParamF(3) = gstrNombreEmpresa
                                    
            If tabInstitucion.Tab = 1 And strEstado = Reg_Edicion Then
                aReportParamS(0) = strTipoEntidad
                aReportParamS(1) = tdgConsulta.Columns(0).Value
                aReportParamS(2) = Codigo_Listar_Individual
                aReportParamF(2) = ObtenerDescripcionParametro("TIPPER", strTipoEntidad)
            Else
                aReportParamS(0) = strTipoInstitucion
                aReportParamS(1) = Valor_Caracter
                aReportParamS(2) = Codigo_Listar_Todos
                aReportParamF(2) = ObtenerDescripcionParametro("TIPPER", strTipoInstitucion)
            End If
            
    End Select

    gstrSelFrml = Valor_Caracter
    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"

    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())

    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
    frmReporte.Show vbModal

    Set frmReporte = Nothing

    Screen.MousePointer = vbNormal
    
End Sub

Private Function TodoOK() As Boolean

    TodoOK = False
    
    If Trim(txtRazonSocial) = "" Then
        MsgBox "El Campo Razón Social no es Válido!.", vbCritical
        Exit Function
    End If
    
    If cboSector.ListIndex = -1 Then
        MsgBox "Seleccione Sector.", vbCritical
        Exit Function
    End If
    
    If cboCiiu.ListIndex = -1 Then
        MsgBox "Seleccione CIIU.", vbCritical
        Exit Function
    End If

    If cboGrupo.ListIndex = -1 Then
        MsgBox "Seleccione Grupo.", vbCritical
        Exit Function
    End If

    If chkExtranjero.Value Then
        If cboExtranjero.ListIndex = -1 Then
            MsgBox "Seleccione Nacionalidad.", vbCritical
            Exit Function
        End If
    End If
                                                        
    If Not blnNemotecnicoUnico Then
        MsgBox "El Nemotecnico debe ser Unico!", vbCritical
        Exit Function
    End If
    
    If Trim$(txtNemonico.Text) = Valor_Caracter And strTipoInstitucion = Codigo_Tipo_Persona_Emisor Then
        MsgBox "El Campo Nemotecnico es obligatorio!", vbCritical
        Exit Function
    End If
    
    '*** Si todo paso OK ***
    TodoOK = True

End Function


Private Sub cboCategoria_Click()

    strCodCategoria = ""
    If cboCategoria.ListIndex < 0 Then Exit Sub
    
    strCodCategoria = Trim(arrCategoria(cboCategoria.ListIndex))
    
    If cboClasificadoraI.ListCount > 0 Then cboClasificadoraI.ListIndex = 0
    If cboClasificadoraII.ListCount > 0 Then cboClasificadoraII.ListIndex = 0
        
End Sub


Private Sub cboCiiu_Click()

    strCodCiiu = ""
    If cboCiiu.ListIndex < 0 Then Exit Sub
    
    strCodCiiu = Trim(arrCiiu(cboCiiu.ListIndex))
    
End Sub



Private Sub cboClasePersona_Click()
    
    strClasePersona = ""
    If cboClasePersona.ListIndex < Valor_Numero Then Exit Sub
    
    strClasePersona = Trim(arrClasePersona(cboClasePersona.ListIndex))
    
    If strClasePersona = Codigo_Persona_Juridica Then
        lblInstitucion(2).Caption = "Razón Social"
    Else
        lblInstitucion(2).Caption = "Ap. y Nombres"
    End If

    cboTipoIdentidad.Clear
    
    strsql = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPIDE' AND ValorParametro ='" & strClasePersona & "' ORDER BY DescripParametro"
    CargarControlLista strsql, cboTipoIdentidad, arrTipoIdentidad(), Sel_Defecto
    If cboTipoIdentidad.ListCount > Valor_Numero Then cboTipoIdentidad.ListIndex = Valor_Numero
    
End Sub

Private Sub cboClasificadora1_Click()

    Dim strsql As String
    
    strCodClasificadora1 = Valor_Caracter
    If cboClasificadora1.ListIndex < 0 Then Exit Sub
    
    strCodClasificadora1 = Trim(arrClasificadora1(cboClasificadora1.ListIndex))
    
    strsql = "SELECT (CodRiesgo + CodSubRiesgo) CODIGO, CodSubRiesgo DESCRIP FROM ClasificacionRiesgoDetalle WHERE CodClasificadora='" & strCodClasificadora1 & "' AND CodCategoria='" & strCodPlazo & "'"
    CargarControlLista strsql, cboSubRiesgo1, arrSubRiesgo1(), Sel_Defecto
    
    If cboSubRiesgo1.ListCount > 0 Then cboSubRiesgo1.ListIndex = 0
    
End Sub


Private Sub cboClasificadora2_Click()

    Dim strsql As String
    
    strCodClasificadora2 = Valor_Caracter
    If cboClasificadora2.ListIndex < 0 Then Exit Sub
    
    strCodClasificadora2 = Trim(arrClasificadora2(cboClasificadora2.ListIndex))
    
    strsql = "SELECT (CodRiesgo + CodSubRiesgo) CODIGO, CodSubRiesgo DESCRIP FROM ClasificacionRiesgoDetalle WHERE CodClasificadora='" & strCodClasificadora2 & "' AND CodCategoria='" & strCodPlazo & "'"
    CargarControlLista strsql, cboSubRiesgo2, arrSubRiesgo2(), Sel_Defecto
    
    If cboSubRiesgo2.ListCount > 0 Then cboSubRiesgo2.ListIndex = 0
    
End Sub


Private Sub cboClasificadoraI_Click()

    Dim strsql As String
    
    strCodClasificadoraI = Valor_Caracter
    If cboClasificadoraI.ListIndex < 0 Then Exit Sub
    
    strCodClasificadoraI = Trim(arrClasificadoraI(cboClasificadoraI.ListIndex))
    
    strsql = "SELECT (CodRiesgo + CodSubRiesgo) CODIGO, CodSubRiesgo DESCRIP FROM ClasificacionRiesgoDetalle WHERE CodClasificadora='" & strCodClasificadoraI & "' AND CodCategoria='" & strCodCategoria & "'"
    CargarControlLista strsql, cboSubRiesgoI, arrSubRiesgoI(), Sel_Defecto
    
    If cboSubRiesgoI.ListCount > 0 Then cboSubRiesgoI.ListIndex = 0
    
End Sub


Private Sub cboClasificadoraII_Click()

    Dim strsql As String
    
    strCodClasificadoraII = ""
    If cboClasificadoraII.ListIndex < 0 Then Exit Sub
    
    strCodClasificadoraII = Trim(arrClasificadoraII(cboClasificadoraII.ListIndex))
    
    strsql = "SELECT (CodRiesgo + CodSubRiesgo) CODIGO, CodSubRiesgo DESCRIP FROM ClasificacionRiesgoDetalle WHERE CodClasificadora='" & strCodClasificadoraII & "' AND CodCategoria='" & strCodCategoria & "'"
    CargarControlLista strsql, cboSubRiesgoII, arrSubRiesgoII(), Sel_Defecto
    
    If cboSubRiesgoII.ListCount > 0 Then cboSubRiesgoII.ListIndex = 0
    
End Sub


Private Sub cboDepartamento_Click()

    Dim strsql As String
    
    strCodDepartamento = ""
    If cboDepartamento.ListIndex < 0 Then Exit Sub
    
    strCodDepartamento = Trim(arrDepartamento(cboDepartamento.ListIndex))
    
    strsql = "{ call up_ACSelDatosParametro(2,'" & strCodPais & "','" & strCodDepartamento & "') }"
    CargarControlLista strsql, cboProvincia, arrProvincia(), Sel_Defecto
    
    If cboProvincia.ListCount > -1 Then cboProvincia.ListIndex = 0
    

End Sub


Private Sub cboDistrito_Click()
        
    strCodDistrito = ""
    If cboDistrito.ListIndex < 0 Then Exit Sub
    
    strCodDistrito = Trim(arrDistrito(cboDistrito.ListIndex))
    
End Sub


Private Sub cboExtranjero_Click()

    strCodNacionalidad = ""
    If cboExtranjero.ListIndex < 0 Then Exit Sub
    
    strCodNacionalidad = Trim(arrNacionalidad(cboExtranjero.ListIndex))
    
End Sub


Private Sub cboGrupo_Click()

    strCodGrupo = ""
    If cboGrupo.ListIndex < 0 Then Exit Sub
    
    strCodGrupo = Trim(arrGrupo(cboGrupo.ListIndex))
    
End Sub


Private Sub cboInstrumentoEmitido_Click()

    strCodFile = Valor_Caracter: strCodDetalleFile = Valor_Caracter
    If cboInstrumentoEmitido.ListIndex < 0 Then Exit Sub
    
    strCodFile = Left(arrInstrumentoEmitido(cboInstrumentoEmitido.ListIndex), 3)
    strCodDetalleFile = Right(arrInstrumentoEmitido(cboInstrumentoEmitido.ListIndex), 3)
    
End Sub


Private Sub cboPais_Click()

    Dim strsql As String
    
    strCodPais = ""
    If cboPais.ListIndex < 0 Then Exit Sub
    
    strCodPais = Trim(arrPais(cboPais.ListIndex))
    
    strsql = "{ call up_ACSelDatosParametro(1,'" & strCodPais & "') }"
    CargarControlLista strsql, cboDepartamento, arrDepartamento(), Sel_Defecto
    
    If cboDepartamento.ListCount > -1 Then cboDepartamento.ListIndex = 0
    
End Sub


Private Sub cboPlazo_Click()

    strCodPlazo = ""
    If cboPlazo.ListIndex < 0 Then Exit Sub
    
    strCodPlazo = Trim(arrPlazo(cboPlazo.ListIndex))
    
    If cboClasificadora1.ListCount > 0 Then cboClasificadora1.ListIndex = 0
    If cboClasificadora2.ListCount > 0 Then cboClasificadora2.ListIndex = 0
    
End Sub


Private Sub cboProvincia_Click()

    Dim strsql As String
    
    strCodProvincia = ""
    If cboProvincia.ListIndex < 0 Then Exit Sub
    
    strCodProvincia = Trim(arrProvincia(cboProvincia.ListIndex))
    
    strsql = "{ call up_ACSelDatosParametro(3,'" & strCodPais & "','" & strCodDepartamento & "','" & strCodProvincia & "') }"
    CargarControlLista strsql, cboDistrito, arrDistrito(), Sel_Defecto
    
    If cboDistrito.ListCount > -1 Then cboDistrito.ListIndex = 0
    
End Sub


Private Sub cboSector_Click()

    strCodSector = ""
    If cboSector.ListIndex < 0 Then Exit Sub
    
    strCodSector = Trim(arrSector(cboSector.ListIndex))
    
End Sub




Private Sub cboSubRiesgo1_Click()

    Dim adoRegistro As ADODB.Recordset
    
    strCodRiesgo1 = Valor_Caracter: strCodSubRiesgo1 = Valor_Caracter
    intOrdenClasificadoraI = 0: lblSubRiesgo1.Caption = Valor_Caracter
    lblRiesgo.Caption = Valor_Caracter: lblSubRiesgo.Caption = Valor_Caracter
    If cboSubRiesgo1.ListIndex < 0 Then Exit Sub
    
    strCodRiesgo1 = Left(arrSubRiesgo1(cboSubRiesgo1.ListIndex), 2)
    strCodSubRiesgo1 = Right(arrSubRiesgo1(cboSubRiesgo1.ListIndex), 10)
    
    Set adoRegistro = New ADODB.Recordset
    
    adoComm.CommandText = "SELECT EquivalenciaRiesgo,NumOrden FROM ClasificacionRiesgoDetalle WHERE CodCategoria='" & strCodPlazo & "' AND CodClasificadora='" & strCodClasificadora1 & "' AND CodSubRiesgo='" & strCodSubRiesgo1 & "'"
    Set adoRegistro = adoComm.Execute
    
    If Not adoRegistro.EOF Then
        lblSubRiesgo1.Caption = Trim(adoRegistro("EquivalenciaRiesgo"))
        intOrdenClasificadoraI = adoRegistro("NumOrden")
    End If
    adoRegistro.Close
    
    If intOrdenClasificadoraI >= intOrdenClasificadoraII Then
        adoComm.CommandText = "SELECT CodParametro,ValorParametro FROM AuxiliarParametro WHERE CodTipoParametro='TIPRIE' AND CodParametro='" & strCodRiesgo1 & "'"
        Set adoRegistro = adoComm.Execute

        If Not adoRegistro.EOF Then
            strCodRiesgoFinal = Trim(adoRegistro("CodParametro"))
            lblRiesgo.Caption = Trim(adoRegistro("ValorParametro"))
            lblSubRiesgo.Caption = Trim(lblSubRiesgo1.Caption)
        End If
        adoRegistro.Close
    Else
        adoComm.CommandText = "SELECT CodParametro,ValorParametro FROM AuxiliarParametro WHERE CodTipoParametro='TIPRIE' AND CodParametro='" & strCodRiesgo2 & "'"
        Set adoRegistro = adoComm.Execute

        If Not adoRegistro.EOF Then
            strCodRiesgoFinal = Trim(adoRegistro("CodParametro"))
            lblRiesgo.Caption = Trim(adoRegistro("ValorParametro"))
            lblSubRiesgo.Caption = Trim(lblSubRiesgo2.Caption)
        End If
        adoRegistro.Close
    End If
    Set adoRegistro = Nothing
    
End Sub


Private Sub cboSubRiesgo2_Click()

    Dim adoRegistro As ADODB.Recordset
    
    strCodRiesgo2 = Valor_Caracter: strCodSubRiesgo2 = Valor_Caracter
    intOrdenClasificadoraII = 0: lblSubRiesgo2.Caption = Valor_Caracter
    lblRiesgo.Caption = Valor_Caracter: lblSubRiesgo.Caption = Valor_Caracter
    If cboSubRiesgo2.ListIndex < 0 Then Exit Sub
    
    strCodRiesgo2 = Left(arrSubRiesgo2(cboSubRiesgo2.ListIndex), 2)
    strCodSubRiesgo2 = Right(arrSubRiesgo2(cboSubRiesgo2.ListIndex), 10)
    
    Set adoRegistro = New ADODB.Recordset
    
    adoComm.CommandText = "SELECT EquivalenciaRiesgo,NumOrden FROM ClasificacionRiesgoDetalle WHERE CodCategoria='" & strCodPlazo & "' AND CodClasificadora='" & strCodClasificadora2 & "' AND CodSubRiesgo='" & strCodSubRiesgo2 & "'"
    Set adoRegistro = adoComm.Execute
    
    If Not adoRegistro.EOF Then
        lblSubRiesgo2.Caption = Trim(adoRegistro("EquivalenciaRiesgo"))
        intOrdenClasificadoraII = adoRegistro("NumOrden")
    End If
    adoRegistro.Close
    
    If intOrdenClasificadoraII >= intOrdenClasificadoraI Then
        adoComm.CommandText = "SELECT CodParametro,ValorParametro FROM AuxiliarParametro WHERE CodTipoParametro='TIPRIE' AND CodParametro='" & strCodRiesgo2 & "'"
        Set adoRegistro = adoComm.Execute

        If Not adoRegistro.EOF Then
            strCodRiesgoFinal = Trim(adoRegistro("CodParametro"))
            lblRiesgo.Caption = Trim(adoRegistro("ValorParametro"))
            lblSubRiesgo.Caption = Trim(lblSubRiesgo2.Caption)
        End If
        adoRegistro.Close
    Else
        adoComm.CommandText = "SELECT CodParametro,ValorParametro FROM AuxiliarParametro WHERE CodTipoParametro='TIPRIE' AND CodParametro='" & strCodRiesgo1 & "'"
        Set adoRegistro = adoComm.Execute

        If Not adoRegistro.EOF Then
            strCodRiesgoFinal = Trim(adoRegistro("CodParametro"))
            lblRiesgo.Caption = Trim(adoRegistro("ValorParametro"))
            lblSubRiesgo.Caption = Trim(lblSubRiesgo1.Caption)
        End If
        adoRegistro.Close
    End If
    Set adoRegistro = Nothing
    
End Sub


Private Sub cboSubRiesgoI_Click()

    Dim adoRegistro As ADODB.Recordset
    
    strCodRiesgoI = Valor_Caracter: strCodSubRiesgoI = Valor_Caracter
    intOrdenClasificadoraII = 0: lblSubRiesgoI.Caption = Valor_Caracter
    lblRiesgoEntidad.Caption = Valor_Caracter: lblSubRiesgoEntidad.Caption = Valor_Caracter
    If cboSubRiesgoI.ListIndex < 0 Then Exit Sub
    
    strCodRiesgoI = Left(arrSubRiesgoI(cboSubRiesgoI.ListIndex), 2)
    strCodSubRiesgoI = Right(arrSubRiesgoI(cboSubRiesgoI.ListIndex), 10)
    
    Set adoRegistro = New ADODB.Recordset
    
    adoComm.CommandText = "SELECT EquivalenciaRiesgo,NumOrden FROM ClasificacionRiesgoDetalle WHERE CodCategoria='" & strCodCategoria & "' AND CodClasificadora='" & strCodClasificadoraI & "' AND CodSubRiesgo='" & strCodSubRiesgoI & "'"
    Set adoRegistro = adoComm.Execute
    
    If Not adoRegistro.EOF Then
        lblSubRiesgoI.Caption = Trim(adoRegistro("EquivalenciaRiesgo"))
        intOrdenClasificadoraI = adoRegistro("NumOrden")
    End If
    adoRegistro.Close
                
    If intOrdenClasificadoraI <= intOrdenClasificadoraII Then
        adoComm.CommandText = "SELECT ValorParametro FROM AuxiliarParametro WHERE CodTipoParametro='TIPRIE' AND CodParametro='" & strCodRiesgoI & "'"
        Set adoRegistro = adoComm.Execute

        If Not adoRegistro.EOF Then
            lblRiesgoEntidad.Caption = Trim(adoRegistro("ValorParametro"))
            lblSubRiesgoEntidad.Caption = Trim(lblSubRiesgoI.Caption)
        End If
        adoRegistro.Close
    Else
        adoComm.CommandText = "SELECT ValorParametro FROM AuxiliarParametro WHERE CodTipoParametro='TIPRIE' AND CodParametro='" & strCodRiesgoII & "'"
        Set adoRegistro = adoComm.Execute

        If Not adoRegistro.EOF Then
            lblRiesgoEntidad.Caption = Trim(adoRegistro("ValorParametro"))
            lblSubRiesgoEntidad.Caption = Trim(lblSubRiesgoII.Caption)
        End If
        adoRegistro.Close
    End If
    Set adoRegistro = Nothing
        
End Sub


Private Sub cboSubRiesgoII_Click()

    Dim adoRegistro As ADODB.Recordset
    
    strCodRiesgoII = Valor_Caracter: strCodSubRiesgoII = Valor_Caracter
    intOrdenClasificadoraI = 0: lblSubRiesgoII.Caption = Valor_Caracter
    lblRiesgoEntidad.Caption = Valor_Caracter: lblSubRiesgoEntidad.Caption = Valor_Caracter
    If cboSubRiesgoII.ListIndex < 0 Then Exit Sub
    
    strCodRiesgoII = Left(arrSubRiesgoII(cboSubRiesgoII.ListIndex), 2)
    strCodSubRiesgoII = Right(arrSubRiesgoII(cboSubRiesgoII.ListIndex), 10)
    
    Set adoRegistro = New ADODB.Recordset
    
    adoComm.CommandText = "SELECT EquivalenciaRiesgo,NumOrden FROM ClasificacionRiesgoDetalle WHERE CodCategoria='" & strCodCategoria & "' AND CodClasificadora='" & strCodClasificadoraII & "' AND CodSubRiesgo='" & strCodSubRiesgoII & "'"
    Set adoRegistro = adoComm.Execute
    
    If Not adoRegistro.EOF Then
        lblSubRiesgoII.Caption = Trim(adoRegistro("EquivalenciaRiesgo"))
        intOrdenClasificadoraI = adoRegistro("NumOrden")
    End If
    adoRegistro.Close
    
    If intOrdenClasificadoraII <= intOrdenClasificadoraI Then
        adoComm.CommandText = "SELECT ValorParametro FROM AuxiliarParametro WHERE CodTipoParametro='TIPRIE' AND CodParametro='" & strCodRiesgoII & "'"
        Set adoRegistro = adoComm.Execute
    
        If Not adoRegistro.EOF Then
            lblRiesgoEntidad.Caption = Trim(adoRegistro("ValorParametro"))
            lblSubRiesgoEntidad.Caption = Trim(lblSubRiesgoII.Caption)
        End If
        adoRegistro.Close
    Else
        adoComm.CommandText = "SELECT ValorParametro FROM AuxiliarParametro WHERE CodTipoParametro='TIPRIE' AND CodParametro='" & strCodRiesgoI & "'"
        Set adoRegistro = adoComm.Execute
    
        If Not adoRegistro.EOF Then
            lblRiesgoEntidad.Caption = Trim(adoRegistro("ValorParametro"))
            lblSubRiesgoEntidad.Caption = Trim(lblSubRiesgoI.Caption)
        End If
        adoRegistro.Close
    End If
    Set adoRegistro = Nothing
    
End Sub


Private Sub cboTipoEntidad_Click()

    strTipoEntidad = Valor_Caracter
    If cboTipoEntidad.ListIndex < 0 Then Exit Sub
    
    strTipoEntidad = arrTipoEntidad(cboTipoEntidad.ListIndex)
    
    chkBanco.Value = False
    
    If strTipoEntidad = Codigo_Tipo_Persona_Emisor Then
        frmMainMdi.stbMdi.Panels(3).Text = "Adicionar emisor..."
        chkBanco.Visible = True
        txtCapital.Enabled = True
        txtObligaciones.Enabled = True
        txtLimiteComite.Enabled = True
        txtInstrumentosEmitidos.Enabled = True
        cmdParametros.Visible = True
        


        
    ElseIf strTipoEntidad = Codigo_Tipo_Persona_Agente Then
        frmMainMdi.stbMdi.Panels(3).Text = "Adicionar agente..."
        chkBanco.Visible = False
        txtCapital.Enabled = True
        txtObligaciones.Enabled = True
        txtLimiteComite.Enabled = True
        txtInstrumentosEmitidos.Enabled = False
        cmdParametros.Visible = False
        

    Else
        frmMainMdi.stbMdi.Panels(3).Text = "Adicionar proveedor..."
        chkBanco.Visible = False
        txtCapital.Enabled = True
        txtObligaciones.Enabled = True
        txtLimiteComite.Enabled = True
        txtInstrumentosEmitidos.Enabled = False
        cmdParametros.Visible = False
        

    End If
    
End Sub



Private Sub cboTipoIdentidad_Click()

    strTipoIdentidad = Valor_Caracter
    If cboTipoIdentidad.ListIndex < 0 Then Exit Sub
    
    strTipoIdentidad = arrTipoIdentidad(cboTipoIdentidad.ListIndex)


End Sub

Private Sub cboTipoInstitucion_Click()

    strTipoInstitucion = ""
    If cboTipoInstitucion.ListIndex < 0 Then Exit Sub
                    
    strTipoInstitucion = Trim(arrTipoInstitucion(cboTipoInstitucion.ListIndex))
    
    If strTipoInstitucion = Codigo_Tipo_Persona_Emisor Then
        cmdParametros.Visible = True
    ElseIf strTipoInstitucion = Codigo_Tipo_Persona_Agente Then
        cmdParametros.Visible = False
    End If
    
    Call Buscar
    
End Sub



Private Sub chkBanco_Click()

    If chkBanco.Value Then
        'fraRetencion.Visible = True
        fraClasificacion.Enabled = True
    Else
        'fraRetencion.Visible = False
        fraClasificacion.Enabled = False
    End If
    
End Sub

Private Sub chkExtranjero_Click()

    If chkExtranjero.Value Then
        cboExtranjero.Enabled = True
    Else
        cboExtranjero.Enabled = False
    End If
    
End Sub





























Private Sub cmdAgregar_Click()

    Dim intRegistro As Integer
    
    With adoComm
        .CommandText = "UPDATE EmisionInstitucionPersona SET " & _
            "CodClasificadoraI='" & strCodClasificadora1 & "'," & _
            "CodSubRiesgoI='" & strCodSubRiesgo1 & "'," & _
            "CodEquivalenciaI='" & Trim(lblSubRiesgo1.Caption) & "'," & _
            "CodClasificadoraII='" & strCodClasificadora2 & "'," & _
            "CodSubRiesgoII='" & strCodSubRiesgo2 & "'," & _
            "CodEquivalenciaII='" & Trim(lblSubRiesgo2.Caption) & "'," & _
            "CodRiesgoFinal='" & strCodRiesgoFinal & "'," & _
            "CodSubRiesgoFinal='" & Trim(lblSubRiesgo.Caption) & "' " & _
            "WHERE CodEmisor='" & strCodInstitucion & "' AND CodFile='" & strCodFile & "' AND " & _
            "CodDetalleFile='" & strCodDetalleFile & "' AND CodCategoriaRiesgo='" & strCodPlazo & "'"
        adoConn.Execute .CommandText, intRegistro
        
        If intRegistro = 0 Then
            .CommandText = "INSERT INTO EmisionInstitucionPersona VALUES ('" & _
                strCodInstitucion & "','" & strCodFile & "','" & _
                strCodDetalleFile & "','" & strCodPlazo & "','" & _
                strCodClasificadora1 & "','" & strCodSubRiesgo1 & "','" & _
                Trim(lblSubRiesgo1.Caption) & "','" & strCodClasificadora2 & "','" & _
                strCodSubRiesgo2 & "','" & Trim(lblSubRiesgo2.Caption) & "','" & _
                strCodRiesgoFinal & "','" & Trim(lblSubRiesgo.Caption) & "')"
            adoConn.Execute .CommandText
        End If
        
        '*** Actualizar los Títulos Valores ***
        .CommandText = "UPDATE InversionOrden SET " & _
            "TipoRiesgo='" & strCodRiesgoFinal & "',SubRiesgo='" & Trim(lblSubRiesgo.Caption) & "' " & _
            "WHERE CodEmisor='" & strCodInstitucion & "' AND CodFile='" & strCodFile & "' AND " & _
            "CodDetalleFile='" & strCodDetalleFile & "'"
        adoConn.Execute .CommandText
        
        .CommandText = "UPDATE InversionOperacion SET " & _
            "TipoRiesgo='" & strCodRiesgoFinal & "',SubRiesgo='" & Trim(lblSubRiesgo.Caption) & "' " & _
            "WHERE CodEmisor='" & strCodInstitucion & "' AND CodFile='" & strCodFile & "' AND " & _
            "CodDetalleFile='" & strCodDetalleFile & "'"
        adoConn.Execute .CommandText
        
        .CommandText = "UPDATE InstrumentoInversion SET " & _
            "CodRiesgo='" & strCodRiesgoFinal & "',CodSubRiesgo='" & Trim(lblSubRiesgo.Caption) & "' " & _
            "WHERE CodEmisor='" & strCodInstitucion & "' AND CodFile='" & strCodFile & "' AND " & _
            "CodDetalleFile='" & strCodDetalleFile & "'"
        adoConn.Execute .CommandText
        
        .CommandText = "UPDATE InversionValorizacion SET " & _
            "CodRiesgo='" & strCodRiesgoFinal & "',CodSubRiesgo='" & Trim(lblSubRiesgo.Caption) & "' " & _
            "WHERE CodEmisor='" & strCodInstitucion & "' AND CodFile='" & strCodFile & "' AND " & _
            "CodDetalleFile='" & strCodDetalleFile & "'"
        adoConn.Execute .CommandText
    End With
    
    Call InicializarCalificacion
    Call CargarTitulosEmitidos
    
End Sub


Private Sub cmdParametros_Click()
    frmInstitucionPersonaParametroGeneral.strCodPersonaPrev = strCodInstitucion
    frmInstitucionPersonaParametroGeneral.Show (vbModal)
End Sub

Private Sub cmdQuitar_Click()

    Dim intRegistro As Integer
    
    With adoComm
        .CommandText = "DELETE EmisionInstitucionPersona "
        .CommandText = .CommandText & "WHERE CodEmisor='" & strCodInstitucion & "' AND CodFile='" & Trim(tdgEmisor.Columns(0).Value) & "' AND "
        .CommandText = .CommandText & "CodDetalleFile='" & Trim(tdgEmisor.Columns(1).Value) & "' AND CodCategoriaRiesgo='" & Trim(tdgEmisor.Columns(2).Value) & "'"
        adoConn.Execute .CommandText
                        
    End With
    
    Call CargarTitulosEmitidos
    
End Sub




Private Sub dgdEmisor_DblClick()

    Dim adoRegistro As ADODB.Recordset
    Dim intRegistro As Integer
        
    Set adoRegistro = New ADODB.Recordset
    
    intRegistro = ObtenerItemLista(arrInstrumentoEmitido(), tdgEmisor.Columns(0).Value & tdgEmisor.Columns(1).Value)
    If intRegistro >= 0 Then cboInstrumentoEmitido.ListIndex = intRegistro
    
    intRegistro = ObtenerItemLista(arrPlazo(), Trim(tdgEmisor.Columns(2).Value))
    If intRegistro >= 0 Then cboPlazo.ListIndex = intRegistro
    
    intRegistro = ObtenerItemLista(arrClasificadora1(), Trim(tdgEmisor.Columns(3).Value))
    If intRegistro >= 0 Then cboClasificadora1.ListIndex = intRegistro
    
    adoComm.CommandText = "SELECT CodRiesgo FROM ClasificacionRiesgoDetalle WHERE CodClasificadora='" & strCodClasificadora1 & "' AND CodCategoria='" & strCodPlazo & "' AND CodSubRiesgo='" & Trim(tdgEmisor.Columns(4).Value) & "'"
    Set adoRegistro = adoComm.Execute
    
    If Not adoRegistro.EOF Then
        intRegistro = ObtenerItemLista(arrSubRiesgo1(), Trim(adoRegistro("CodRiesgo")) & Trim(tdgEmisor.Columns(4).Value))
        If intRegistro >= 0 Then cboSubRiesgo1.ListIndex = intRegistro
    End If
    adoRegistro.Close
    
    intRegistro = ObtenerItemLista(arrClasificadora2(), Trim(tdgEmisor.Columns(5).Value))
    If intRegistro >= 0 Then cboClasificadora2.ListIndex = intRegistro
    
    adoComm.CommandText = "SELECT CodRiesgo FROM ClasificacionRiesgoDetalle WHERE CodClasificadora='" & strCodClasificadora2 & "' AND CodCategoria='" & strCodPlazo & "' AND CodSubRiesgo='" & Trim(tdgEmisor.Columns(6).Value) & "'"
    Set adoRegistro = adoComm.Execute
    
    If Not adoRegistro.EOF Then
        intRegistro = ObtenerItemLista(arrSubRiesgo2(), Trim(adoRegistro("CodRiesgo")) & Trim(tdgEmisor.Columns(6).Value))
        If intRegistro >= 0 Then cboSubRiesgo2.ListIndex = intRegistro
    End If
    adoRegistro.Close: Set adoRegistro = Nothing
    
    strCodRiesgoFinal = Trim(tdgEmisor.Columns(7).Value)
    lblRiesgo.Caption = Trim(tdgEmisor.Columns(10).Value)
    lblSubRiesgo.Caption = Trim(tdgEmisor.Columns(8).Value)
    
End Sub


Private Sub Form_Activate()
'
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
    
    For intCont = 0 To (lblInstitucion.Count - 1)
        Call FormatoEtiqueta(lblInstitucion(intCont), vbLeftJustify)
    Next
    
    For Each elemento In Me.Controls
    
        If TypeOf elemento Is TDBGrid Then
            Call FormatoGrilla(elemento)
        End If
    
    Next
            
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call OcultarReportes
    Set frmInstitucion = Nothing
    
End Sub

Private Sub lblSubRiesgo1_DblClick()

    Dim adoRegistro As ADODB.Recordset
    
    If Trim(lblSubRiesgo1.Caption) <> "" Then
        Set adoRegistro = New ADODB.Recordset
        
        adoComm.CommandText = "SELECT CodParametro,ValorParametro FROM AuxiliarParametro WHERE CodTipoParametro='TIPRIE' AND CodParametro='" & strCodRiesgo1 & "'"
        Set adoRegistro = adoComm.Execute
        
        If Not adoRegistro.EOF Then
            strCodRiesgoFinal = Trim(adoRegistro("CodParametro"))
            lblRiesgo.Caption = Trim(adoRegistro("ValorParametro"))
            lblSubRiesgo.Caption = Trim(lblSubRiesgo1.Caption)
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End If
    
End Sub


Private Sub lblSubRiesgo2_DblClick()

    Dim adoRegistro As ADODB.Recordset
    
    If Trim(lblSubRiesgo2.Caption) <> "" Then
        Set adoRegistro = New ADODB.Recordset
        
        adoComm.CommandText = "SELECT CodParametro,ValorParametro FROM AuxiliarParametro WHERE CodTipoParametro='TIPRIE' AND CodParametro='" & strCodRiesgo2 & "'"
        Set adoRegistro = adoComm.Execute
        
        If Not adoRegistro.EOF Then
            strCodRiesgoFinal = Trim(adoRegistro("CodParametro"))
            lblRiesgo.Caption = Trim(adoRegistro("ValorParametro"))
            lblSubRiesgo.Caption = Trim(lblSubRiesgo2.Caption)
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End If
    
End Sub


Private Sub lblSubRiesgoI_DblClick()

    Dim adoRegistro As ADODB.Recordset
    
    If Trim(lblSubRiesgoI.Caption) <> Valor_Caracter Then
        Set adoRegistro = New ADODB.Recordset
        
        adoComm.CommandText = "SELECT ValorParametro FROM AuxiliarParametro WHERE CodTipoParametro='TIPRIE' AND CodParametro='" & strCodRiesgoI & "'"
        Set adoRegistro = adoComm.Execute
        
        If Not adoRegistro.EOF Then
            lblRiesgoEntidad.Caption = Trim(adoRegistro("ValorParametro"))
            lblSubRiesgoEntidad.Caption = Trim(lblSubRiesgoI.Caption)
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End If
        
End Sub


Private Sub lblSubRiesgoII_DblClick()

    Dim adoRegistro As ADODB.Recordset
    
    If Trim(lblSubRiesgoII.Caption) <> "" Then
        Set adoRegistro = New ADODB.Recordset
        
        adoComm.CommandText = "SELECT ValorParametro FROM AuxiliarParametro WHERE CodTipoParametro='TIPRIE' AND CodParametro='" & strCodRiesgoII & "'"
        Set adoRegistro = adoComm.Execute
        
        If Not adoRegistro.EOF Then
            lblRiesgoEntidad.Caption = Trim(adoRegistro("ValorParametro"))
            lblSubRiesgoEntidad.Caption = Trim(lblSubRiesgoII.Caption)
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End If
    
End Sub


Private Sub tabInstitucion_Click(PreviousTab As Integer)

    Select Case tabInstitucion.Tab
        Case 1, 2, 3
            If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vModify)
            If strEstado = Reg_Defecto Then tabInstitucion.Tab = 0
            
    End Select
    
End Sub

Private Sub txtCapital_Change()

    Call FormatoCajaTexto(txtCapital, Decimales_Monto)
    
End Sub

Private Sub txtCapital_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtCapital, Decimales_Monto)
        
End Sub


Private Sub txtDiasBanco_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "N")
    
End Sub


Private Sub txtDiasOtro_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "N")
    
End Sub


Private Sub txtDireccion1_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
End Sub


Private Sub txtDireccion2_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
End Sub


Private Sub txtFax_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "N")
    
End Sub


Private Sub txtInstrumentosEmitidos_Change()

    Call FormatoCajaTexto(txtInstrumentosEmitidos, Decimales_Monto)
    
End Sub

Private Sub txtInstrumentosEmitidos_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtInstrumentosEmitidos, Decimales_Monto)
    
End Sub


Private Sub txtLimiteComite_Change()

    Call FormatoCajaTexto(txtLimiteComite, Decimales_Monto)
    
End Sub

Private Sub txtLimiteComite_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtLimiteComite, Decimales_Monto)
    
End Sub

Private Sub txtNemonico_Change()
    Dim adoAuxiliar As ADODB.Recordset
    If Trim$(txtNemonico.Text) <> Valor_Caracter Then
        adoComm.CommandText = "select COUNT(*) as Count from InstitucionPersona where DescripNemonico = '" & _
                                txtNemonico.Text & "' and TipoPersona = '" & strTipoInstitucion & "' and CodPersona <> '" & strCodInstitucion & "'"
        Set adoAuxiliar = adoComm.Execute
        If adoAuxiliar("Count") > 0 Then
            MsgBox "El Nemotecnico debe ser Unico!", vbCritical
            blnNemotecnicoUnico = False
        Else
            blnNemotecnicoUnico = True
        End If
    End If
End Sub

Private Sub txtNemonico_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub

Private Sub txtObligaciones_Change()

    Call FormatoCajaTexto(txtObligaciones, Decimales_Monto)
    
End Sub


Private Sub txtObligaciones_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtObligaciones, Decimales_Monto)
    
End Sub


Private Sub txtRazonSocial_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
End Sub


Private Sub txtRuc_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "N")
    
End Sub


Private Sub txtTelefono_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "N")
    
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
