VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{830D5A06-9B70-4F8C-98B6-7A19C4E7760B}#1.0#0"; "TAMControls.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmDesembolsoAcreencias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Desembolsos - Operaciones de Acreencias"
   ClientHeight    =   9180
   ClientLeft      =   1500
   ClientTop       =   1680
   ClientWidth     =   13650
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "frmDesembolsoAcreencias.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9180
   ScaleWidth      =   13650
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
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
      Left            =   5490
      Picture         =   "frmDesembolsoAcreencias.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   8430
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
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
      Left            =   6990
      Picture         =   "frmDesembolsoAcreencias.frx":0600
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   8430
      Visible         =   0   'False
      Width           =   1200
   End
   Begin TAMControls2.ucBotonEdicion2 cmdOpcion 
      Height          =   735
      Left            =   600
      TabIndex        =   67
      Top             =   8430
      Width           =   4200
      _ExtentX        =   7408
      _ExtentY        =   1296
      Buttons         =   3
      Caption0        =   "&Nuevo"
      Tag0            =   "0"
      Visible0        =   0   'False
      ToolTipText0    =   "Nuevo"
      Caption1        =   "&Eliminar"
      Tag1            =   "4"
      Visible1        =   0   'False
      ToolTipText1    =   "Eliminar"
      Caption2        =   "&Buscar"
      Tag2            =   "5"
      Visible2        =   0   'False
      ToolTipText2    =   "Buscar"
      UserControlWidth=   4200
   End
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   11760
      TabIndex        =   66
      Top             =   8430
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
      Caption0        =   "&Salir"
      Tag0            =   "9"
      Visible0        =   0   'False
      ToolTipText0    =   "Salir"
      UserControlWidth=   1200
   End
   Begin MSAdodcLib.Adodc adoConsulta 
      Height          =   330
      Left            =   9000
      Top             =   8430
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
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin TabDlg.SSTab tabRFCortoPlazo 
      Height          =   8325
      Left            =   30
      TabIndex        =   43
      Top             =   60
      Width           =   13605
      _ExtentX        =   23998
      _ExtentY        =   14684
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabsPerRow      =   5
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
      TabPicture(0)   =   "frmDesembolsoAcreencias.frx":0B62
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "tdgConsulta"
      Tab(0).Control(1)=   "fraCriterio"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Datos Orden Inversión"
      TabPicture(1)   =   "frmDesembolsoAcreencias.frx":0B7E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraResumen"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraDatosTitulo"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "fraDatosAnexo"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "fraDatosBasicos"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Negociación"
      TabPicture(2)   =   "frmDesembolsoAcreencias.frx":0B9A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraDatosNegociacion"
      Tab(2).Control(1)=   "fraComisionMontoFL1"
      Tab(2).Control(2)=   "fraComisiones"
      Tab(2).ControlCount=   3
      Begin VB.Frame fraDatosBasicos 
         Caption         =   "Datos Básicos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1875
         Left            =   120
         TabIndex        =   88
         Top             =   450
         Width           =   13305
         Begin VB.ComboBox cboComisionista 
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
            Left            =   8805
            Style           =   2  'Dropdown List
            TabIndex        =   207
            Top             =   1440
            Width           =   4185
         End
         Begin VB.ComboBox cboSubClaseInstrumento 
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
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   199
            Top             =   1440
            Width           =   4185
         End
         Begin VB.ComboBox cboGestor 
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
            Left            =   8805
            Style           =   2  'Dropdown List
            TabIndex        =   198
            Top             =   1080
            Width           =   4185
         End
         Begin VB.ComboBox cboObligado 
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
            Left            =   8805
            Style           =   2  'Dropdown List
            TabIndex        =   195
            Top             =   720
            Width           =   4185
         End
         Begin VB.ComboBox cboEmisor 
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
            Left            =   8805
            Style           =   2  'Dropdown List
            TabIndex        =   194
            Top             =   360
            Width           =   4185
         End
         Begin VB.ComboBox cboLineaCliente 
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
            Left            =   2060
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1920
            Visible         =   0   'False
            Width           =   4185
         End
         Begin VB.ComboBox cboFondoOrden 
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
            Left            =   2060
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   360
            Width           =   4185
         End
         Begin VB.ComboBox cboTipoInstrumentoOrden 
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
            Left            =   2060
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   720
            Width           =   4185
         End
         Begin VB.ComboBox cboClaseInstrumento 
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
            Left            =   2060
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1080
            Width           =   4185
         End
         Begin VB.Label lblComisionistaInversion 
            Caption         =   "Comisionista"
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
            Left            =   6690
            TabIndex        =   208
            Top             =   1470
            Width           =   1095
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Obligado"
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
            Index           =   87
            Left            =   6675
            TabIndex        =   197
            Top             =   780
            Width           =   765
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Gestor"
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
            Index           =   88
            Left            =   6675
            TabIndex        =   196
            Top             =   1140
            Width           =   570
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Especificar Línea"
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
            Index           =   115
            Left            =   330
            TabIndex        =   193
            Top             =   1980
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Clase"
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
            Index           =   29
            Left            =   300
            TabIndex        =   93
            Top             =   1140
            Width           =   480
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fondo"
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
            Index           =   0
            Left            =   300
            TabIndex        =   92
            Top             =   420
            Width           =   540
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Instrumento"
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
            Left            =   300
            TabIndex        =   91
            Top             =   780
            Width           =   1005
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Emisor"
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
            Index           =   6
            Left            =   6675
            TabIndex        =   90
            Top             =   420
            Width           =   570
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "SubClase"
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
            Index           =   42
            Left            =   300
            TabIndex        =   89
            Top             =   1500
            Width           =   810
         End
      End
      Begin VB.Frame fraComisiones 
         Height          =   1005
         Left            =   -74880
         TabIndex        =   166
         Top             =   5400
         Width           =   6615
         Begin VB.TextBox txtComisionAgente 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   4260
            Locked          =   -1  'True
            MaxLength       =   45
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   240
            Width           =   2025
         End
         Begin TAMControls.TAMTextBox txtPorcenIgv 
            Height          =   285
            Index           =   0
            Left            =   2860
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   570
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
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
            Alignment       =   1
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Locked          =   -1  'True
            Container       =   "frmDesembolsoAcreencias.frx":0BB6
            Text            =   "0.000000"
            Decimales       =   6
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            EnterTab        =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   6
         End
         Begin TAMControls.TAMTextBox txtPorcenAgente 
            Height          =   285
            Index           =   0
            Left            =   2860
            TabIndex        =   60
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
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
            Alignment       =   1
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "frmDesembolsoAcreencias.frx":0BD2
            Text            =   "0.000000"
            Decimales       =   6
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            EnterTab        =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   6
         End
         Begin VB.TextBox txtComisionBolsa 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   4260
            Locked          =   -1  'True
            MaxLength       =   45
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   1140
            Visible         =   0   'False
            Width           =   2025
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisión de Desembolso"
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
            Index           =   94
            Left            =   240
            TabIndex        =   173
            Top             =   300
            Width           =   2100
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Otras comisiones"
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
            Index           =   95
            Left            =   390
            TabIndex        =   172
            Top             =   1185
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "IGV"
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
            Index           =   34
            Left            =   240
            TabIndex        =   171
            Top             =   630
            Width           =   360
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisión BVL"
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
            Index           =   26
            Left            =   720
            TabIndex        =   170
            Top             =   1185
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.Label lblComisionIgv 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Index           =   0
            Left            =   4260
            TabIndex        =   169
            Tag             =   "0.00"
            ToolTipText     =   "Monto de Comisión IGV"
            Top             =   570
            Width           =   2025
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisiones"
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
            Index           =   91
            Left            =   660
            TabIndex        =   168
            Top             =   1155
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label lblPorcenBolsa 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Index           =   0
            Left            =   2760
            TabIndex        =   167
            Tag             =   "0"
            ToolTipText     =   "Porcentaje de Comisión BVL"
            Top             =   1140
            Visible         =   0   'False
            Width           =   1335
         End
      End
      Begin VB.Frame fraComisionMontoFL1 
         Caption         =   "Comisiones y Montos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4095
         Left            =   -74880
         TabIndex        =   96
         Top             =   3615
         Width           =   6615
         Begin TAMControls.TAMTextBox txtDiasCobroMinimoInteres 
            Height          =   285
            Left            =   2865
            TabIndex        =   202
            TabStop         =   0   'False
            Top             =   1110
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   503
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
            Alignment       =   1
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Locked          =   -1  'True
            Container       =   "frmDesembolsoAcreencias.frx":0BEE
            Text            =   "0"
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            Apariencia      =   1
            Borde           =   1
            MaximoValor     =   999999
         End
         Begin VB.TextBox txtInteresCorrido 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   8280
            MaxLength       =   45
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   2970
            Width           =   2025
         End
         Begin VB.CheckBox chkInteresCorrido 
            Caption         =   "Interés Corrido"
            Enabled         =   0   'False
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
            Index           =   0
            Left            =   8010
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   2520
            Width           =   1935
         End
         Begin VB.TextBox txtIntAdicional 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   10980
            Locked          =   -1  'True
            MaxLength       =   45
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   690
            Width           =   2025
         End
         Begin TAMControls.TAMTextBox txtTirBruta1 
            Height          =   315
            Left            =   960
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   3630
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
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
            Alignment       =   1
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "frmDesembolsoAcreencias.frx":0C0A
            Text            =   "0.000000"
            Decimales       =   6
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            EnterTab        =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   6
         End
         Begin VB.TextBox txtTirNeta 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   2610
            MaxLength       =   45
            TabIndex        =   47
            Top             =   5580
            Width           =   1365
         End
         Begin VB.TextBox txtPrecioUnitario 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   1080
            MaxLength       =   45
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   6060
            Width           =   1340
         End
         Begin VB.CommandButton cmdCalculo 
            Caption         =   "#"
            Height          =   375
            Left            =   510
            TabIndex        =   42
            TabStop         =   0   'False
            ToolTipText     =   "Calcular Valor al Vencimiento y TIRs de la orden"
            Top             =   5550
            Width           =   375
         End
         Begin TAMControls.TAMTextBox txtMontoVencimiento1 
            Height          =   315
            Left            =   4260
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   3600
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   556
            BackColor       =   16777215
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "frmDesembolsoAcreencias.frx":0C26
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            EnterTab        =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   2
         End
         Begin TAMControls.TAMTextBox txtPorcenIgvInt 
            Height          =   285
            Index           =   0
            Left            =   2860
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   750
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
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
            Alignment       =   1
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Locked          =   -1  'True
            Container       =   "frmDesembolsoAcreencias.frx":0C42
            Text            =   "0.000000"
            Decimales       =   6
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            EnterTab        =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   6
         End
         Begin TAMControls.TAMTextBox txtImptoInteresCorrido 
            Height          =   315
            Index           =   0
            Left            =   7800
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   2700
            Width           =   1335
            _ExtentX        =   2355
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
            Alignment       =   1
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "frmDesembolsoAcreencias.frx":0C5E
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            EnterTab        =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   2
         End
         Begin TAMControls.TAMTextBox txtImptoInteres 
            Height          =   285
            Index           =   0
            Left            =   2860
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   390
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
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
            Alignment       =   1
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Locked          =   -1  'True
            Container       =   "frmDesembolsoAcreencias.frx":0C7A
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            EnterTab        =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   2
         End
         Begin TAMControls.TAMTextBox txtImptoInteresAdic 
            Height          =   315
            Index           =   0
            Left            =   9480
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   690
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
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
            Alignment       =   1
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "frmDesembolsoAcreencias.frx":0C96
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            EnterTab        =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   2
         End
         Begin TAMControls.TAMTextBox txtCobroMinimoInteres 
            Height          =   285
            Left            =   4260
            TabIndex        =   204
            TabStop         =   0   'False
            Top             =   1110
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   503
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
            Alignment       =   1
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Locked          =   -1  'True
            Container       =   "frmDesembolsoAcreencias.frx":0CB2
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   2
         End
         Begin TAMControls.TAMTextBox txtPorcenIgvInt2 
            Height          =   285
            Index           =   1
            Left            =   2865
            TabIndex        =   205
            TabStop         =   0   'False
            Top             =   1470
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
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
            Alignment       =   1
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Locked          =   -1  'True
            Container       =   "frmDesembolsoAcreencias.frx":0CCE
            Text            =   "0.000000"
            Decimales       =   6
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            EnterTab        =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   6
         End
         Begin TAMControls.TAMTextBox txtIGVCobroMinimoInteres 
            Height          =   285
            Left            =   4260
            TabIndex        =   206
            TabStop         =   0   'False
            Top             =   1470
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   503
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
            Alignment       =   1
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Locked          =   -1  'True
            Container       =   "frmDesembolsoAcreencias.frx":0CEA
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   2
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Días"
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
            Index           =   5
            Left            =   3750
            TabIndex        =   203
            Top             =   1170
            Width           =   450
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IGV Cobro Mínimo de Interés"
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
            Left            =   240
            TabIndex        =   201
            Top             =   1560
            Width           =   2700
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cobro Mínimo de Interés"
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
            Left            =   240
            TabIndex        =   200
            Top             =   1170
            Width           =   2310
         End
         Begin VB.Label lblMontoMinimoComision 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Left            =   4290
            TabIndex        =   191
            Tag             =   "0.00"
            ToolTipText     =   "Fecha de Liquidación"
            Top             =   1860
            Width           =   2025
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto Min Comisión"
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
            Index           =   186
            Left            =   2400
            TabIndex        =   190
            Top             =   1890
            Width           =   1695
         End
         Begin VB.Label lblFechaVencimientoAdic 
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6750
            TabIndex        =   183
            Tag             =   "0.00"
            ToolTipText     =   "Fecha de Vencimiento del Título de la Orden"
            Top             =   1530
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label lblDiasAdic 
            AutoSize        =   -1  'True
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   2430
            TabIndex        =   176
            Top             =   1770
            Width           =   1005
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "IGV Intereses"
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
            Index           =   111
            Left            =   240
            TabIndex        =   163
            Top             =   810
            Width           =   1170
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Interés Prov.Protesto"
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
            Index           =   93
            Left            =   6960
            TabIndex        =   161
            Top             =   765
            Width           =   1800
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Interés"
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
            Index           =   92
            Left            =   240
            TabIndex        =   160
            Top             =   450
            Width           =   585
         End
         Begin VB.Label lblIntAdelantado 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Index           =   0
            Left            =   4260
            TabIndex        =   159
            Tag             =   "0.00"
            ToolTipText     =   "SubTotal de la Orden"
            Top             =   390
            Width           =   2025
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto Total"
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
            Index           =   36
            Left            =   270
            TabIndex        =   105
            Top             =   2910
            Width           =   1035
         End
         Begin VB.Label lblDescripMoneda 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Nuevos Soles (S/.)"
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
            Index           =   0
            Left            =   4455
            TabIndex        =   104
            Top             =   120
            Width           =   1665
         End
         Begin VB.Label lblMontoVencimiento 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4320
            TabIndex        =   103
            Tag             =   "0.00"
            Top             =   5610
            Width           =   2025
         End
         Begin VB.Label lblTirBruta 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1080
            TabIndex        =   102
            Tag             =   "0.00"
            Top             =   5580
            Width           =   1335
         End
         Begin VB.Label lblTirNeta 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Left            =   2860
            TabIndex        =   101
            Tag             =   "0.00"
            Top             =   3630
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   1
            X1              =   360
            X2              =   6300
            Y1              =   2760
            Y2              =   2760
         End
         Begin VB.Label lblMontoTotal 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Index           =   0
            Left            =   4260
            TabIndex        =   100
            Tag             =   "0.00"
            ToolTipText     =   "Monto Total de la Orden"
            Top             =   2835
            Width           =   2025
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Valor al Vencimiento"
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
            Index           =   37
            Left            =   4440
            TabIndex        =   99
            Top             =   3315
            Width           =   1755
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "TIR Bruta"
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
            Index           =   38
            Left            =   1230
            TabIndex        =   98
            Top             =   3315
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "TIR Neta"
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
            Index           =   39
            Left            =   3100
            TabIndex        =   97
            Top             =   3315
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   2
            Visible         =   0   'False
            X1              =   240
            X2              =   6300
            Y1              =   3180
            Y2              =   3180
         End
         Begin VB.Label lblComisionIgvInt 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Index           =   0
            Left            =   4260
            TabIndex        =   164
            Tag             =   "0.00"
            ToolTipText     =   "Monto de Comisión IGV"
            Top             =   750
            Width           =   2025
         End
      End
      Begin VB.Frame fraDatosNegociacion 
         Caption         =   "Negociación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3225
         Left            =   -74880
         TabIndex        =   94
         Top             =   360
         Width           =   9495
         Begin VB.ComboBox cboPeriodoCapitalizacion 
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
            Left            =   6720
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   240
            Width           =   2295
         End
         Begin VB.ComboBox cboPeriodoTasa 
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
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   940
            Width           =   1900
         End
         Begin VB.ComboBox cboCobroInteres 
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
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   1640
            Width           =   1900
         End
         Begin VB.TextBox txtValorNominalDcto 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2160
            Locked          =   -1  'True
            MaxLength       =   45
            TabIndex        =   38
            Text            =   "0.00"
            Top             =   2690
            Width           =   1900
         End
         Begin VB.ComboBox cboNegociacion 
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
            Left            =   6720
            Style           =   2  'Dropdown List
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   600
            Width           =   2295
         End
         Begin VB.TextBox txtTasa 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2160
            Locked          =   -1  'True
            MaxLength       =   45
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   240
            Width           =   1900
         End
         Begin VB.ComboBox cboBaseAnual 
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
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   1290
            Width           =   1900
         End
         Begin VB.ComboBox cboTipoTasa 
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
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   590
            Width           =   1900
         End
         Begin VB.TextBox txtTipoCambio 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   6720
            MaxLength       =   45
            TabIndex        =   39
            TabStop         =   0   'False
            Text            =   "0.00"
            Top             =   2430
            Visible         =   0   'False
            Width           =   1830
         End
         Begin TAMControls.TAMTextBox txtPorcenDctoValorNominal 
            Height          =   315
            Left            =   2160
            TabIndex        =   37
            Top             =   2340
            Width           =   1900
            _ExtentX        =   3360
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
            Alignment       =   1
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "frmDesembolsoAcreencias.frx":0D06
            Text            =   "0.000000"
            Decimales       =   6
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            EnterTab        =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   6
         End
         Begin TAMControls.TAMTextBox txtValorNominal 
            Height          =   315
            Left            =   2160
            TabIndex        =   36
            Top             =   1990
            Width           =   1900
            _ExtentX        =   3360
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
            Alignment       =   1
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Locked          =   -1  'True
            Container       =   "frmDesembolsoAcreencias.frx":0D22
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            EnterTab        =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   2
         End
         Begin VB.Label lblPeriodoCapitalizacion 
            AutoSize        =   -1  'True
            Caption         =   "Periodo Capitaliza"
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
            Left            =   4680
            TabIndex        =   189
            Top             =   315
            Width           =   1545
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Periodo Tasa"
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
            Index           =   181
            Left            =   360
            TabIndex        =   188
            Top             =   1000
            Width           =   1140
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Modo cobro Interés"
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
            Index           =   110
            Left            =   360
            TabIndex        =   162
            Top             =   1700
            Width           =   1650
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Valor Nominal Dcto."
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
            Index           =   90
            Left            =   360
            TabIndex        =   158
            Top             =   2750
            Width           =   1710
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "% V.Nominal Dcto."
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
            Index           =   89
            Left            =   360
            TabIndex        =   157
            Top             =   2400
            Width           =   1590
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Mecanismo"
            Enabled         =   0   'False
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
            Index           =   31
            Left            =   4680
            TabIndex        =   155
            Top             =   690
            Width           =   960
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Plazo (Días)"
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
            Index           =   83
            Left            =   4680
            TabIndex        =   150
            Top             =   2145
            Width           =   1050
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Vencimiento"
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
            Index           =   50
            Left            =   4680
            TabIndex        =   149
            Top             =   1785
            Width           =   1050
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Emisión"
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
            Index           =   48
            Left            =   4680
            TabIndex        =   148
            Top             =   1425
            Width           =   645
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Liquidación"
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
            Index           =   47
            Left            =   4680
            TabIndex        =   147
            Top             =   1065
            Width           =   975
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Valor Nominal"
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
            Index           =   12
            Left            =   360
            TabIndex        =   146
            Top             =   2050
            Width           =   1185
         End
         Begin VB.Label lblDiasPlazo 
            BorderStyle     =   1  'Fixed Single
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
            Left            =   6720
            TabIndex        =   112
            Tag             =   "0.00"
            ToolTipText     =   "Días de Plazo del Título de la Orden"
            Top             =   2070
            Width           =   1815
         End
         Begin VB.Label lblFechaVencimiento 
            BorderStyle     =   1  'Fixed Single
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
            Left            =   6720
            TabIndex        =   111
            Tag             =   "0.00"
            ToolTipText     =   "Fecha de Vencimiento del Título de la Orden"
            Top             =   1710
            Width           =   1815
         End
         Begin VB.Label lblFechaEmision 
            BorderStyle     =   1  'Fixed Single
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
            Left            =   6720
            TabIndex        =   110
            Tag             =   "0.00"
            ToolTipText     =   "Fecha Emisión"
            Top             =   1350
            Width           =   1815
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000015&
            X1              =   4470
            X2              =   4470
            Y1              =   240
            Y2              =   3090
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tasa Facial"
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
            Index           =   11
            Left            =   360
            TabIndex        =   109
            Top             =   300
            Width           =   1005
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Base Anual"
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
            Index           =   40
            Left            =   360
            TabIndex        =   108
            Top             =   1350
            Width           =   975
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Tasa"
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
            Index           =   41
            Left            =   360
            TabIndex        =   107
            Top             =   650
            Width           =   1300
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Cambio"
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
            Index           =   17
            Left            =   4680
            TabIndex        =   106
            Top             =   2475
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.Label lblFechaLiquidacion 
            BorderStyle     =   1  'Fixed Single
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
            Left            =   6720
            TabIndex        =   95
            Tag             =   "0.00"
            ToolTipText     =   "Fecha de Liquidación"
            Top             =   990
            Width           =   1815
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmDesembolsoAcreencias.frx":0D3E
         Height          =   5355
         Left            =   -74850
         OleObjectBlob   =   "frmDesembolsoAcreencias.frx":0D58
         TabIndex        =   44
         Top             =   2520
         Width           =   13305
      End
      Begin VB.Frame fraCriterio 
         Caption         =   "Criterios de Búsqueda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2025
         Left            =   -74880
         TabIndex        =   78
         Top             =   390
         Width           =   13305
         Begin VB.ComboBox cboLineaClienteLista 
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
            Left            =   8520
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   1140
            Width           =   3105
         End
         Begin VB.CommandButton cmdEnviar 
            Caption         =   "En&viar"
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
            Left            =   11820
            Picture         =   "frmDesembolsoAcreencias.frx":BA1C
            Style           =   1  'Graphical
            TabIndex        =   45
            ToolTipText     =   "Enviar a BackOffice"
            Top             =   1110
            Width           =   1200
         End
         Begin VB.ComboBox cboEstado 
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
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1080
            Width           =   4605
         End
         Begin VB.ComboBox cboTipoInstrumento 
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
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   720
            Width           =   4605
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
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   360
            Width           =   4605
         End
         Begin MSComCtl2.DTPicker dtpFechaOrdenDesde 
            Height          =   285
            Left            =   9240
            TabIndex        =   3
            Top             =   390
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
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
            CheckBox        =   -1  'True
            Format          =   180617217
            CurrentDate     =   38785
         End
         Begin MSComCtl2.DTPicker dtpFechaOrdenHasta 
            Height          =   285
            Left            =   11595
            TabIndex        =   4
            Top             =   390
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
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
            CheckBox        =   -1  'True
            Format          =   180617217
            CurrentDate     =   38785
         End
         Begin MSComCtl2.DTPicker dtpFechaLiquidacionDesde 
            Height          =   285
            Left            =   9240
            TabIndex        =   5
            Top             =   750
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
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
            CheckBox        =   -1  'True
            Format          =   180617217
            CurrentDate     =   38785
         End
         Begin MSComCtl2.DTPicker dtpFechaLiquidacionHasta 
            Height          =   285
            Left            =   11595
            TabIndex        =   6
            Top             =   750
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
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
            CheckBox        =   -1  'True
            Format          =   180617217
            CurrentDate     =   38785
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Especificar Línea"
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
            Index           =   116
            Left            =   6450
            TabIndex        =   174
            Top             =   1185
            Width           =   1500
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Estado"
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
            Index           =   23
            Left            =   360
            TabIndex        =   87
            Top             =   1155
            Width           =   600
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Instrumento"
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
            Index           =   22
            Left            =   360
            TabIndex        =   86
            Top             =   795
            Width           =   1005
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
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
            Height          =   195
            Index           =   21
            Left            =   10920
            TabIndex        =   85
            Top             =   405
            Width           =   510
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
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
            Height          =   195
            Index           =   20
            Left            =   8520
            TabIndex        =   84
            Top             =   435
            Width           =   555
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fondo"
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
            Index           =   19
            Left            =   360
            TabIndex        =   83
            Top             =   435
            Width           =   540
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Orden"
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
            Index           =   43
            Left            =   6450
            TabIndex        =   82
            Top             =   435
            Width           =   1110
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Liquidación"
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
            Index           =   44
            Left            =   6450
            TabIndex        =   81
            Top             =   795
            Width           =   1560
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
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
            Height          =   195
            Index           =   45
            Left            =   8520
            TabIndex        =   80
            Top             =   795
            Width           =   555
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
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
            Height          =   195
            Index           =   46
            Left            =   10920
            TabIndex        =   79
            Top             =   795
            Width           =   510
         End
      End
      Begin VB.Frame fraDatosAnexo 
         Caption         =   "Datos del Anexo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   178
         Top             =   2340
         Width           =   13305
         Begin VB.TextBox txtNumAnexo 
            BackColor       =   &H00FFFFFF&
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2060
            MaxLength       =   45
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   600
            Width           =   2010
         End
         Begin VB.CheckBox cbxGenerarLetra 
            Caption         =   "Generar Letra"
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
            Left            =   8790
            TabIndex        =   12
            Top             =   660
            Value           =   1  'Checked
            Width           =   2565
         End
         Begin VB.TextBox txtNumContrato 
            BackColor       =   &H00FFFFFF&
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2060
            MaxLength       =   45
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   240
            Width           =   2010
         End
         Begin TAMControls.TAMTextBox txtTotalMNAnexo 
            Height          =   315
            Left            =   6870
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
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
            Alignment       =   1
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Locked          =   -1  'True
            Container       =   "frmDesembolsoAcreencias.frx":BF77
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            EnterTab        =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   2
         End
         Begin TAMControls.TAMTextBox txtTotalMEAnexo 
            Height          =   315
            Left            =   6870
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
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
            Alignment       =   1
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Locked          =   -1  'True
            Container       =   "frmDesembolsoAcreencias.frx":BF93
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            EnterTab        =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   2
         End
         Begin TAMControls.TAMTextBox txtTotalDctosAnexo 
            Height          =   315
            Left            =   11580
            TabIndex        =   11
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
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
            Alignment       =   1
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "frmDesembolsoAcreencias.frx":BFAF
            Text            =   "0"
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            EnterTab        =   -1  'True
            Apariencia      =   1
            Borde           =   1
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Número de Anexo"
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
            Index           =   114
            Left            =   210
            TabIndex        =   192
            Top             =   660
            Width           =   1500
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Total Anexo Descontado"
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
            Index           =   178
            Left            =   4470
            TabIndex        =   182
            Top             =   660
            Width           =   2115
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Total Anexo"
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
            Index           =   176
            Left            =   4470
            TabIndex        =   181
            Top             =   300
            Width           =   1035
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Cant. Documentos a descontar"
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
            Height          =   165
            Index           =   177
            Left            =   8790
            TabIndex        =   180
            Top             =   300
            Width           =   2700
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Número de Contrato"
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
            Index           =   163
            Left            =   210
            TabIndex        =   179
            Top             =   300
            Width           =   1695
         End
      End
      Begin VB.Frame fraDatosTitulo 
         Caption         =   "Datos de la Orden"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3045
         Left            =   120
         TabIndex        =   70
         Top             =   3390
         Width           =   13305
         Begin VB.ComboBox cboMonedaDocumento 
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
            Left            =   2880
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   630
            Width           =   2130
         End
         Begin VB.ComboBox cboResponsablePago 
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
            Left            =   5910
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   2250
            Width           =   2265
         End
         Begin VB.TextBox txtNumDocDscto 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2880
            MaxLength       =   10
            TabIndex        =   13
            Top             =   270
            Width           =   2130
         End
         Begin VB.TextBox txtObservacion 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   645
            Left            =   10620
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   29
            Top             =   1920
            Width           =   2280
         End
         Begin VB.TextBox txtNemonico 
            BackColor       =   &H00FFFFFF&
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
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1410
            MaxLength       =   15
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   1200
            Width           =   2055
         End
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
            Left            =   10620
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   1200
            Width           =   2280
         End
         Begin VB.TextBox txtDescripOrden 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   5910
            MaxLength       =   45
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   1200
            Width           =   2850
         End
         Begin MSComCtl2.DTPicker dtpFechaOrden 
            Height          =   315
            Left            =   1410
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   1530
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   180617217
            CurrentDate     =   38776
         End
         Begin MSComCtl2.DTPicker dtpFechaLiquidacion 
            Height          =   315
            Left            =   1410
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   1890
            Width           =   1455
            _ExtentX        =   2566
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
            Format          =   180617217
            CurrentDate     =   38776
         End
         Begin MSComCtl2.DTPicker dtpFechaEmision 
            Height          =   315
            Left            =   11700
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   630
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
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
            Format          =   180617217
            CurrentDate     =   38776
         End
         Begin MSComCtl2.DTPicker dtpFechaVencimiento 
            Height          =   315
            Left            =   5910
            TabIndex        =   24
            Top             =   1530
            Width           =   1575
            _ExtentX        =   2778
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
            Format          =   180617217
            CurrentDate     =   38776
         End
         Begin MSComCtl2.DTPicker dtpFechaPago 
            Height          =   315
            Left            =   5910
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   1890
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   180617217
            CurrentDate     =   38776
         End
         Begin MSComCtl2.DTPicker dtpFechaVencimientoDcto 
            Height          =   315
            Left            =   11700
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
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
            Format          =   180617217
            CurrentDate     =   38776
         End
         Begin TAMControls.TAMTextBox txtDiasPlazo 
            Height          =   315
            Left            =   1410
            TabIndex        =   25
            Top             =   2250
            Width           =   975
            _ExtentX        =   1720
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
            Alignment       =   1
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "frmDesembolsoAcreencias.frx":BFCB
            Text            =   "0"
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            EnterTab        =   -1  'True
            Apariencia      =   1
            Borde           =   1
         End
         Begin TAMControls.TAMTextBox txtValorNominalDescuento 
            Height          =   315
            Left            =   10620
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   1560
            Width           =   2265
            _ExtentX        =   3995
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
            Alignment       =   1
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Locked          =   -1  'True
            Container       =   "frmDesembolsoAcreencias.frx":BFE7
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            EnterTab        =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   2
         End
         Begin TAMControls.TAMTextBox txtValorNominalDocumento 
            Height          =   315
            Left            =   7800
            TabIndex        =   15
            Top             =   270
            Width           =   1455
            _ExtentX        =   2566
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
            Alignment       =   1
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "frmDesembolsoAcreencias.frx":C003
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            EnterTab        =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   2
         End
         Begin TAMControls.TAMTextBox txtTipoCambioDescuento 
            Height          =   315
            Left            =   7800
            TabIndex        =   16
            Top             =   630
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
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
            Alignment       =   1
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "frmDesembolsoAcreencias.frx":C01F
            Text            =   "0.000000"
            Decimales       =   6
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            EnterTab        =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   6
         End
         Begin VB.CheckBox chkDiasAdicional 
            Caption         =   "Adicionar días protesto"
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
            Left            =   8250
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   2310
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   2595
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Moneda Documento"
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
            Index           =   182
            Left            =   210
            TabIndex        =   187
            Top             =   690
            Width           =   1710
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Valor Nominal Documento"
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
            Index           =   183
            Left            =   5520
            TabIndex        =   186
            Top             =   330
            Width           =   2205
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Cambio"
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
            Index           =   185
            Left            =   5520
            TabIndex        =   185
            Top             =   690
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Valor Nominal"
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
            Index           =   184
            Left            =   9210
            TabIndex        =   184
            Top             =   1590
            Width           =   1185
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Pagador"
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
            Index           =   138
            Left            =   3810
            TabIndex        =   177
            Top             =   2310
            Width           =   720
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Vencimiento Operación"
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
            Index           =   117
            Left            =   3810
            TabIndex        =   175
            Top             =   1590
            Width           =   1965
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Nro.  Documento a descontar"
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
            Index           =   113
            Left            =   210
            TabIndex        =   165
            Top             =   360
            Width           =   2520
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Instrucciones"
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
            Index           =   30
            Left            =   9210
            TabIndex        =   156
            Top             =   1980
            Width           =   1155
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Pago"
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
            Index           =   86
            Left            =   3810
            TabIndex        =   154
            Top             =   1950
            Width           =   450
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Nemónico"
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
            Index           =   85
            Left            =   210
            TabIndex        =   153
            Top             =   1260
            Width           =   840
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Emisión Documento"
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
            Index           =   8
            Left            =   9720
            TabIndex        =   77
            Top             =   660
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
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
            Height          =   195
            Index           =   3
            Left            =   9210
            TabIndex        =   76
            Top             =   1260
            Width           =   690
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Plazo (DIAS)"
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
            Index           =   10
            Left            =   210
            TabIndex        =   75
            Top             =   2310
            Width           =   1095
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Vcto.Documento"
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
            Index           =   9
            Left            =   9720
            TabIndex        =   74
            Top             =   330
            Visible         =   0   'False
            Width           =   1425
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
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
            Index           =   7
            Left            =   3810
            TabIndex        =   73
            Top             =   1260
            Width           =   1005
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Liquidación"
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
            Index           =   14
            Left            =   210
            TabIndex        =   72
            Top             =   1935
            Width           =   975
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Orden"
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
            Index           =   13
            Left            =   210
            TabIndex        =   71
            Top             =   1605
            Width           =   525
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000010&
            BorderStyle     =   6  'Inside Solid
            Index           =   0
            X1              =   150
            X2              =   13110
            Y1              =   1050
            Y2              =   1050
         End
      End
      Begin VB.Frame fraResumen 
         Caption         =   "Resumen Negociación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1785
         Left            =   120
         TabIndex        =   113
         Top             =   6450
         Width           =   13305
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Analítica"
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
            Index           =   84
            Left            =   8730
            TabIndex        =   152
            Top             =   255
            Width           =   765
         End
         Begin VB.Label lblAnalitica 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "???-????????"
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
            Left            =   10800
            TabIndex        =   151
            Tag             =   "0.00"
            Top             =   240
            Width           =   2025
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Precio"
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
            Index           =   54
            Left            =   4530
            TabIndex        =   145
            Top             =   960
            Width           =   555
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Precio"
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
            Index           =   53
            Left            =   210
            TabIndex        =   144
            Top             =   930
            Width           =   555
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "SubTotal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   61
            Left            =   390
            TabIndex        =   143
            Top             =   3000
            Width           =   645
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisiones"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   63
            Left            =   390
            TabIndex        =   142
            Top             =   3330
            Width           =   795
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Intereses Corridos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   64
            Left            =   390
            TabIndex        =   141
            Top             =   3660
            Width           =   1260
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto Total"
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
            Index           =   65
            Left            =   210
            TabIndex        =   140
            Top             =   1320
            Width           =   1035
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "SubTotal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   66
            Left            =   5190
            TabIndex        =   139
            Top             =   3000
            Width           =   645
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisiones"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   67
            Left            =   5190
            TabIndex        =   138
            Top             =   3330
            Width           =   795
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Intereses Corridos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   68
            Left            =   5190
            TabIndex        =   137
            Top             =   3660
            Width           =   1545
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto Total"
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
            Index           =   69
            Left            =   4530
            TabIndex        =   136
            Top             =   1320
            Width           =   1035
         End
         Begin VB.Label lblPrecioResumen 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Index           =   0
            Left            =   2010
            TabIndex        =   135
            Tag             =   "0.00"
            Top             =   915
            Width           =   2025
         End
         Begin VB.Label lblSubTotalResumen 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   2310
            TabIndex        =   134
            Tag             =   "0.00"
            Top             =   2985
            Width           =   2025
         End
         Begin VB.Label lblComisionesResumen 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   2310
            TabIndex        =   133
            Tag             =   "0.00"
            Top             =   3315
            Width           =   2025
         End
         Begin VB.Label lblInteresesResumen 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   2310
            TabIndex        =   132
            Tag             =   "0.00"
            Top             =   3645
            Width           =   2025
         End
         Begin VB.Label lblTotalResumen 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Index           =   0
            Left            =   2010
            TabIndex        =   131
            Tag             =   "0.00"
            Top             =   1305
            Width           =   2025
         End
         Begin VB.Label lblPrecioResumen 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Index           =   1
            Left            =   6060
            TabIndex        =   130
            Tag             =   "0.00"
            Top             =   915
            Width           =   2025
         End
         Begin VB.Label lblSubTotalResumen 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   7200
            TabIndex        =   129
            Tag             =   "0.00"
            Top             =   2985
            Width           =   2025
         End
         Begin VB.Label lblComisionesResumen 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   7230
            TabIndex        =   128
            Tag             =   "0.00"
            Top             =   3315
            Width           =   2025
         End
         Begin VB.Label lblInteresesResumen 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   7230
            TabIndex        =   127
            Tag             =   "0.00"
            Top             =   3645
            Width           =   2025
         End
         Begin VB.Label lblTotalResumen 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Index           =   1
            Left            =   6060
            TabIndex        =   126
            Tag             =   "0.00"
            Top             =   1305
            Width           =   2025
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Contado"
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
            Index           =   70
            Left            =   210
            TabIndex        =   125
            Top             =   600
            Width           =   720
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Plazo"
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
            Index           =   71
            Left            =   4530
            TabIndex        =   124
            Top             =   600
            Width           =   480
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tir Bruta"
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
            Index           =   74
            Left            =   8730
            TabIndex        =   123
            Top             =   960
            Width           =   750
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tir Neta"
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
            Index           =   75
            Left            =   8730
            TabIndex        =   122
            Top             =   1320
            Width           =   705
         End
         Begin VB.Label lblTirBrutaResumen 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Left            =   10800
            TabIndex        =   121
            Tag             =   "0.00"
            Top             =   960
            Width           =   2025
         End
         Begin VB.Label lblTirNetaResumen 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Left            =   10800
            TabIndex        =   120
            Tag             =   "0.00"
            Top             =   1290
            Width           =   2025
         End
         Begin VB.Line Line4 
            BorderColor     =   &H80000015&
            X1              =   4230
            X2              =   4230
            Y1              =   240
            Y2              =   1620
         End
         Begin VB.Label lblCantidadResumen 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Left            =   2010
            TabIndex        =   119
            Tag             =   "0.00"
            Top             =   240
            Width           =   2025
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad Facial"
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
            Index           =   77
            Left            =   210
            TabIndex        =   118
            Top             =   270
            Width           =   1335
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Vencimiento"
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
            Index           =   78
            Left            =   4530
            TabIndex        =   117
            Top             =   255
            Width           =   1050
         End
         Begin VB.Label lblVencimientoResumen 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Left            =   6060
            TabIndex        =   116
            Tag             =   "0.00"
            Top             =   240
            Width           =   2025
         End
         Begin VB.Label lblDescripMonedaResumen 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Nuevos Soles (S/.)"
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
            Index           =   0
            Left            =   2190
            TabIndex        =   115
            Top             =   600
            Width           =   1665
         End
         Begin VB.Label lblDescripMonedaResumen 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Nuevos Soles (S/.)"
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
            Left            =   6210
            TabIndex        =   114
            Top             =   600
            Width           =   1665
         End
         Begin VB.Line Line3 
            BorderColor     =   &H80000015&
            X1              =   8250
            X2              =   8250
            Y1              =   150
            Y2              =   1590
         End
      End
   End
End
Attribute VB_Name = "frmDesembolsoAcreencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Ordenes de Instrumentos de Renta Fija Corto Plazo"
Option Explicit

Dim arrFondo()                As String, arrFondoOrden()                As String
Dim arrTipoInstrumento()      As String, arrTipoInstrumentoOrden()      As String
Dim arrEstado()               As String, arrNegociacion()               As String
Dim arrEmisor()               As String, arrMoneda()                    As String
Dim arrObligado()             As String, arrGestor()                    As String
Dim arrBaseAnual()            As String, arrTipoTasa()                  As String
Dim arrOrigen()               As String, arrClaseInstrumento()          As String
Dim arrSubClaseInstrumento()  As String, arrConceptoCosto()             As String
Dim arrResponsablePago()      As String, arrMonedaDocumento()           As String
Dim arrPeriodoTasa()          As String, arrPeriodoCapitalizacion()     As String
Dim arrBaseCalculo()          As String, arrLineaClienteLista()         As String
Dim arrLineaCliente()         As String, arrComisionista()              As String
Dim strCodFondo               As String, strCodFondoOrden               As String
Dim strCodTipoInstrumento     As String, strCodTipoInstrumentoOrden     As String
Dim strCodEstado              As String, strCodTipoOrden                As String
Dim strCodOperacion           As String, strCodNegociacion              As String
Dim strCodEmisor              As String, strCodMoneda                   As String
Dim strCodMonedaDocumento     As String, strPeriodoTasa                 As String
Dim strIndCapitalizable       As String, strPeriodoCapitalizable        As String
Dim strIndTipoCambio          As String, strIndGeneraLetra              As String
Dim strCodObligado            As String, strCodGestor                   As String
Dim strCodBaseAnual           As String, strCodTipoTasa                 As String
Dim strCodOrigen              As String, strCodClaseInstrumento         As String
Dim strCodTitulo              As String, strCodSubClaseInstrumento      As String
Dim strCodConcepto            As String, strCodReportado                As String
Dim strCodGarantia            As String, strCodAgente                  As String
Dim strEstado                 As String, strSQL                         As String
Dim strCodFiador              As String
Dim strLineaClienteLista      As String, strResponsablePago             As String
Dim arrPagoInteres()          As String, strCodFile                     As String
Dim strCodAnalitica           As String

Dim strEstadoOrden            As String, strCodRiesgo                   As String
Dim strCodSubRiesgo           As String, strCalcVcto                    As String
Dim strCodTipoCostoBolsa      As String, strCodComisionista             As String
Dim numSecCondicion           As Integer

Dim strLineaCliente           As String
Dim strTipoPersonaLim         As String
Dim strCodPersonaLim          As String

Dim strCodigosFile            As String, strCodCobroInteres             As String
Dim strViaCobranza            As String, dblTipoCambio                  As Double

Dim dblTasaInteres            As Double, dblPorcentajeComision          As Double
Dim dblMontoMinComisiones     As Double, intUltNumAnexo                 As Integer

Dim dblComisionBolsa          As Double
Dim intBaseCalculo            As Integer

Dim rsg                       As New ADODB.Recordset
Dim rsgVcto                   As New ADODB.Recordset

Dim strCodMonedaParEvaluacion As String, strCodMonedaParPorDefecto      As String
Dim strNumAnexo               As String, intDiasAdicionales             As Integer
Dim datFechaVctoAdicional     As Date

Dim blnCargadoDesdeCartera    As Boolean, blnCargarCabeceraAnexo        As Boolean
Dim blnCancelaPrepago         As Boolean

Dim dblComisionOperacion      As Double
Dim strCodMonedaComision      As String
Dim strPersonalizaComision    As String
Dim dblPorcenDescuento        As Double
Dim strResponsablePagoCancel  As String
Dim dblTotalMNAnexo           As Double
Dim dblTotalMEAnexo           As Double
Dim dblTotalAnexoDscto        As Double


Dim blnEmisorReady            As Boolean
Dim blnPreInfoReady           As Boolean
Dim blnLockEvents             As Boolean


'JAFR: Días mínimos de cobro de intereses
Dim intDiasInteresMinimo           As Integer


Public Sub Adicionar()
    Dim adoAuxiliar As ADODB.Recordset
    
    Dim intCantidadOperaciones     As Integer
    Dim intNumeroDocumentosEnAnexo As Integer
    Dim adoRegistro                As ADODB.Recordset
    
    strNumAnexo = Space$(10)
    txtNumDocDscto.Text = ""

    If Not EsDiaUtil(gdatFechaActual) Then
        MsgBox "No se puede negociar en un día no útil !", vbCritical, Me.Caption
        Exit Sub
    End If
    
    'Obteniendo Días minimos de cobro de interés
    With adoComm
        .CommandText = "select ValorParametro from FondoParametroGeneral where CodFondo = '" & strCodFondo & "' and CodAdministradora = '" & gstrCodAdministradora & "' and CodParametro = '26' "
        Set adoAuxiliar = .Execute
        
        If adoAuxiliar.EOF Then
            intDiasInteresMinimo = 0
        Else
            intDiasInteresMinimo = adoAuxiliar("ValorParametro")
            If intDiasInteresMinimo < 0 Then intDiasInteresMinimo = 0
        End If
        adoAuxiliar.Close: Set adoAuxiliar = Nothing
        txtDiasCobroMinimoInteres.Text = intDiasInteresMinimo
        '        datFechaInteresMinimo = DateAdd("d", DiasInteresMinimo, datFechaEmision)
        '        datFechaValorDeuda = Max(datFechaInteresMinimo, Convertddmmyyyy(gstrFechaActual))
        '        strFechaValorDeuda = Convertyyyymmdd(datFechaValorDeuda)

    End With

    
    If cboTipoInstrumento.ListCount > 1 Then
        frmMainMdi.stbMdi.Panels(3).Text = "Adicionar Orden..."
        
        blnCargarCabeceraAnexo = False

        If tdgConsulta.Row >= 0 Then
            'Se comprueba que el anexo seleccionado esté completo:
            Set adoRegistro = New ADODB.Recordset
            strNumAnexo = adoConsulta.Recordset("NumAnexo")
            strCodEmisor = adoConsulta.Recordset("CodGirador")
            adoComm.CommandText = "select COUNT(NumOrden) as CantidadOperaciones, CantDocumAnexo from InversionOrden where " & "CodFondo = '" & strCodFondo & "' and CodGirador = '" & strCodEmisor & "' and NumAnexo = '" & strNumAnexo & "' and EstadoOrden <> '01' " & "GROUP BY CantDocumAnexo"
                                    
            Set adoRegistro = adoComm.Execute

            If Not (adoRegistro.EOF) Then
                intCantidadOperaciones = adoRegistro("CantidadOperaciones")
                intNumeroDocumentosEnAnexo = adoRegistro("CantDocumAnexo")
            Else
                intCantidadOperaciones = 0
                intNumeroDocumentosEnAnexo = 0
            End If
           
            If blnCancelaPrepago = False And adoConsulta.Recordset("TipoOrden") = Codigo_Orden_Compra And adoConsulta.Recordset("FechaOrden") = gdatFechaActual Then
                If intCantidadOperaciones < intNumeroDocumentosEnAnexo Then
                    If MsgBox("¿Desea continuar con el Anexo Nro. " & Trim$(strNumAnexo) & " de " & Trim$(adoConsulta.Recordset("DesGirador")) & "?", vbQuestion + vbYesNo + vbDefaultButton1, Me.Caption) = vbYes Then
                        'Si la respuesta es afirmativa cargar algunos datos de la orden
                        Call CargarCabeceraAnexo(strCodFondo, gstrCodAdministradora, strNumAnexo, adoConsulta.Recordset("NumOrden"))
                    Else
                        strNumAnexo = Space$(10)
                        Call HabilitaCabeceraAnexo(True)
                    End If

                Else
                    strNumAnexo = Space$(10)
                    Call HabilitaCabeceraAnexo(True)
                End If
            End If
            
        End If
        
        strEstado = Reg_Adicion
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        cmdGuardar.Visible = True
        cmdCancelar.Visible = True
       
        tabRFCortoPlazo.TabEnabled(0) = False
        tabRFCortoPlazo.TabEnabled(1) = True
        tabRFCortoPlazo.TabEnabled(2) = True
        tabRFCortoPlazo.Tab = 1
    
    Else
        MsgBox "Acceso a Negociación Denegada", vbCritical, Me.Caption
    End If
    
End Sub

Private Sub CalcularTirBruta()

'    Dim dblTasaCalculada As Double
'
'    If CDbl(txtPrecioUnitario(0).Text) = 0 Then
'        MsgBox "Por favor ingrese el Precio.", vbCritical, Me.Caption
'        Exit Sub
'    End If
'
'    Me.MousePointer = vbHourglass
'
'    If CDbl(txtPrecioUnitario(0).Text) > 0 Then
'        ReDim Array_Monto(1): ReDim Array_Dias(1)
'        Array_Monto(0) = CDec((CCur(lblSubTotal(0).Caption) + txtInteresCorrido(0).Text) * -1)
'        Array_Dias(0) = dtpFechaLiquidacion.Value
'
'        If strCodBaseAnual = Codigo_Base_Actual_Actual Or strCodBaseAnual = Codigo_Base_Actual_365 Or strCodBaseAnual = Codigo_Base_30_365 Then
'            If strCodTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
'                dblTasaCalculada = ((1 + (CDbl(txtTasa.Text) / 100)) ^ (CDbl(txtDiasPlazo) / 365)) - 1
'            Else
'                dblTasaCalculada = 1 + ((CDbl(txtTasa.Text) / 100 / 365) * CDbl(txtDiasPlazo))
'            End If
'
'        Else
'
'            If strCodTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
'                dblTasaCalculada = ((1 + (CDbl(txtTasa.Text) / 100)) ^ (CDbl(txtDiasPlazo) / 360)) - 1
'            Else
'                dblTasaCalculada = 1 + ((CDbl(txtTasa.Text) / 100 / 360) * CDbl(txtDiasPlazo))
'            End If
'        End If
'
'        If strCalcVcto = "D" Then
'            Array_Monto(1) = CDec(txtCantidad.Text)
'        Else
'            Array_Monto(1) = CDbl(txtCantidad.Text) * (1 + dblTasaCalculada)
'        End If
'
'        Array_Dias(1) = dtpFechaVencimiento.Value
'        lblTirBruta.Caption = CStr(TIR(Array_Monto(), Array_Dias(), (10 / 100)) * 100)
'        lblTirBrutaResumen.Caption = lblTirBruta.Caption
'
'        If strCodTipoOrden = Codigo_Orden_Pacto Then lblTirBrutaResumen.Caption = "0"
'    End If
'
'    Me.MousePointer = vbDefault

End Sub

Private Sub CalcularTirNeta()

'    Dim dblTir           As Double
'    Dim dblTasaCalculada As Double
'
'    If CDbl(lblSubTotal(0).Caption) <= 0 Then
'        MsgBox "Por favor ingrese los datos necesarios para hallar la TIR Neta", vbCritical, Me.Caption
'        Exit Sub
'    End If
'
'    Me.MousePointer = vbHourglass
'
'    ReDim Array_Monto(1): ReDim Array_Dias(1)
'
'    'Array_Monto(0) = CDec((CCur(lblSubTotal(0).Caption) + CCur(txtInteresCorrido(0).Text) + CCur(txtComisionAgente(0).Text) + CCur(txtComisionBolsa(0).Text) + CCur(txtComisionConasev(0).Text) + CCur(lblComisionIgv(0).Caption)) * -1)
'    Array_Monto(0) = CDec((CCur(lblSubTotal(0).Caption) + CCur(txtInteresCorrido(0).Text) + CCur(txtComisionAgente(0).Text) + CCur(txtComisionBolsa(0).Text) + CCur(lblComisionIgv(0).Caption)) * -1)
'    Array_Dias(0) = dtpFechaLiquidacion.Value
'
'    If strCodBaseAnual = Codigo_Base_Actual_Actual Or strCodBaseAnual = Codigo_Base_Actual_365 Or strCodBaseAnual = Codigo_Base_30_365 Then
'        If strCodTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
'            dblTasaCalculada = ((1 + (CDbl(txtTasa.Text) / 100)) ^ (CDbl(txtDiasPlazo) / 365)) - 1
'        Else
'            dblTasaCalculada = 1 + ((CDbl(txtTasa.Text) / 100 / 365) * CDbl(txtDiasPlazo))
'        End If
'
'    Else
'
'        If strCodTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
'            dblTasaCalculada = ((1 + (CDbl(txtTasa.Text) / 100)) ^ (CDbl(txtDiasPlazo) / 360)) - 1
'        Else
'            dblTasaCalculada = 1 + ((CDbl(txtTasa.Text) / 100 / 360) * CDbl(txtDiasPlazo))
'        End If
'    End If
'
'    If strCalcVcto = "D" Then
'        Array_Monto(1) = CDec(txtCantidad.Text)
'    Else
'        Array_Monto(1) = CDbl(txtCantidad.Text) * (1 + dblTasaCalculada)
'    End If
'
'    Array_Dias(1) = dtpFechaVencimiento.Value
'
'    dblTir = TIR(Array_Monto(), Array_Dias(), (10 / 100)) * 100
'
'    lblTirNeta.Caption = CStr(dblTir)
'    lblTirNetaResumen.Caption = CStr(dblTir)
'
'    If strCodTipoOrden = Codigo_Orden_Pacto Then lblTirNetaResumen.Caption = "0"
'    Me.MousePointer = vbDefault

End Sub

Private Sub IniciarComisiones()

    txtComisionAgente(0).Text = "0"
    txtComisionBolsa(0).Text = "0"
    lblComisionIgv(0).Caption = "0"
    
    txtPorcenAgente(0).Text = "0"
    lblPorcenBolsa(0).Caption = "0"
    
    lblPrecioResumen(0).Caption = "0"
    lblSubTotalResumen(0).Caption = "0"
    lblComisionesResumen(0).Caption = "0"
    lblInteresesResumen(0).Caption = "0"
    lblTotalResumen(0).Caption = "0"
        
End Sub

Private Sub LlenarFormulario(strModo As String)

    Dim intRegistro As Integer
  
    Select Case strModo

        Case Reg_Adicion
        
            If blnCargarCabeceraAnexo = False Then  'si no he precargado datos
            
                chkInteresCorrido(0).Value = vbUnchecked

                lblFechaVencimientoAdic.Visible = False

                intRegistro = ObtenerItemLista(arrFondoOrden(), strCodFondo)

                If intRegistro >= 0 Then cboFondoOrden.ListIndex = intRegistro
            
                cboTipoInstrumentoOrden.ListIndex = -1

                If cboTipoInstrumentoOrden.ListCount > 0 Then cboTipoInstrumentoOrden.ListIndex = 0
                                        
                cboNegociacion.ListIndex = -1

                If cboNegociacion.ListCount > 0 Then cboNegociacion.ListIndex = 0
                
                cboEmisor.ListIndex = -1

                If cboEmisor.ListCount > 0 Then cboEmisor.ListIndex = 0
                
                txtTasa.Text = "0"
                
                cboBaseAnual.ListIndex = -1

                If cboBaseAnual.ListCount > 0 Then cboBaseAnual.ListIndex = 0
                
                cboTipoTasa.ListIndex = -1

                If cboTipoTasa.ListCount > 0 Then cboTipoTasa.ListIndex = 0
            
                intRegistro = ObtenerItemLista(arrOrigen(), Codigo_Negociacion_Local)
                
                txtPorcenDctoValorNominal.Text = dblPorcenDescuento
                txtTotalMNAnexo.Text = 0#
                txtTotalMEAnexo.Text = 0#
                txtTotalDctosAnexo.Text = 1
                txtPorcenAgente(0).Text = 0#
                
                
                
                
            End If
            
            If cboResponsablePago.ListCount > 0 Then cboResponsablePago.ListIndex = 2
           
            cboObligado.ListIndex = -1

            If cboObligado.ListCount > 0 Then cboObligado.ListIndex = 0

            intRegistro = ObtenerItemLista(arrMoneda(), strCodMoneda)

            If intRegistro >= 0 Then cboMoneda.ListIndex = intRegistro
            
            cboMonedaDocumento.ListIndex = cboMoneda.ListIndex
            
            dtpFechaOrden.Value = gdatFechaActual
            dtpFechaLiquidacion.Value = dtpFechaOrden.Value
            lblFechaLiquidacion.Caption = CStr(dtpFechaOrden.Value)
            
            txtDiasPlazo.Text = "0"
            lblDiasPlazo.Caption = "0"
            
            txtInteresCorrido(0).Text = "0"
            txtImptoInteresCorrido(0).Text = "0"
            
            txtDescripOrden.Text = Valor_Caracter
            txtNemonico.Text = Valor_Caracter
            txtObservacion.Text = Valor_Caracter
            txtPrecioUnitario(0).Text = "100"

            txtValorNominal.Text = "1"

           ' If blnCargarCabeceraAnexo = False Then txtCantidad.Text = "0"

            lblAnalitica.Caption = "??? - ????????"

            dtpFechaEmision.Value = gdatFechaActual
            dtpFechaVencimientoDcto.Value = dtpFechaEmision.Value
            dtpFechaVencimiento.Value = dtpFechaEmision.Value
            dtpFechaPago.Value = dtpFechaVencimiento.Value
            lblFechaEmision.Caption = CStr(dtpFechaEmision.Value)
            lblFechaVencimiento.Caption = CStr(dtpFechaVencimiento.Value)
            
            lblIntAdelantado(0).Caption = "0"
            
            txtCobroMinimoInteres.Text = "0"
            txtIGVCobroMinimoInteres.Text = "0"
            
            txtIntAdicional(0).Text = "0"
            fraComisiones.Enabled = True
           ' lblSubTotal(0).Caption = "0"
           
            txtCobroMinimoInteres.Text = "0"
            
            txtIGVCobroMinimoInteres.Text = "0"
            
            Call IniciarComisiones
            
            txtInteresCorrido(0).Text = "0"
            txtImptoInteresCorrido(0).Text = "0"
            
            lblMontoTotal(0).Caption = "0"

            lblTirBruta.Caption = "0"
            lblTirNeta.Caption = "0"
            lblMontoVencimiento.Caption = "0"
            lblVencimientoResumen.Caption = "0"

            lblCantidadResumen.Caption = "0"
                                                
            lblTirBrutaResumen.Caption = "0"
            lblTirNetaResumen.Caption = "0"
                        
            txtMontoVencimiento1.Text = "0"
            txtTirBruta1.Text = "0"

            txtTirBruta1.Tag = 0
            txtPrecioUnitario(0).Tag = 0
            txtMontoVencimiento1.Tag = 0

            txtValorNominalDocumento.Text = "0"
            txtValorNominalDescuento.Text = "0"
    
    End Select
  
End Sub

Public Sub Cancelar()

    cmdOpcion.Visible = True
    cmdCancelar.Visible = False
    cmdGuardar.Visible = False

    With tabRFCortoPlazo
        .TabEnabled(0) = True
        .Tab = 0
    End With

    Call Buscar
    
End Sub

Public Sub Eliminar()

    Dim strNumOrden  As String
    Dim strCodTitulo As String
    Dim intRegistro  As Integer

    For intRegistro = 0 To tdgConsulta.SelBookmarks.Count - 1
        adoConsulta.Recordset.MoveFirst
        adoConsulta.Recordset.Move CLng(tdgConsulta.SelBookmarks(intRegistro) - 1), 0
        tdgConsulta.Refresh
        
        strNumOrden = Trim$(adoConsulta.Recordset("NumOrden"))
        strCodTitulo = Valor_Caracter
        strCodEstado = Trim$(adoConsulta.Recordset("EstadoOrden"))
        
        If strEstado = Reg_Consulta Or strEstado = Reg_Edicion Then
            Dim strMensaje As String
            
            'verificar si la orden no está ya anulada
            
            If strCodEstado <> Estado_Orden_Anulada And strCodEstado <> Estado_Orden_Procesada Then
            
                strMensaje = "Se procederá a eliminar la ORDEN " & strNumOrden & " por la " & tdgConsulta.Columns(3) & vbNewLine & vbNewLine & vbNewLine & "¿ Seguro de continuar ?"
                
                If MsgBox(strMensaje, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
            
                    '*** Anular Orden ***
                    adoComm.CommandText = "UPDATE InversionOrden SET EstadoOrden='" & Estado_Orden_Anulada & "' WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & "CodTitulo='" & strCodTitulo & "' AND NumOrden='" & strNumOrden & "'"
                        
                    adoConn.Execute adoComm.CommandText
                    
                    MsgBox Mensaje_Eliminacion_Exitosa, vbExclamation, Me.Caption
                    
                    tabRFCortoPlazo.TabEnabled(0) = True
                    tabRFCortoPlazo.Tab = 0
                    Call Buscar
                    
                    Exit Sub
                End If
                
            Else
    
                If strCodEstado = Estado_Orden_Anulada Then
                    MsgBox "La orden " & strNumOrden & " ya ha sido anulada.", vbExclamation, "Anular Orden"
                Else
                    MsgBox "La orden " & strNumOrden & " ya ha sido procesada." & vbNewLine & "No se puede anular.", vbCritical, "Anular Orden"
                End If
            End If
        
        End If
    Next
    
End Sub

Public Sub Grabar()

    Call Accion(vSave)

End Sub

Public Sub GrabarNew()

    Dim adoRegistro     As ADODB.Recordset
    Dim strFechaOrden   As String, strFechaLiquidacion      As String
    Dim strFechaEmision As String, strFechaVencimiento      As String
    Dim strFechaPago    As String
    Dim strMensaje      As String, strIndTitulo             As String
    Dim intAccion       As Integer
    Dim lngNumError     As Long
    Dim dblTasaInteres  As Double
 
    Dim strMsgError     As String
    
    On Error GoTo CtrlError
   
    If strEstado = Reg_Consulta Then Exit Sub
    
    If (strEstado = Reg_Adicion) And (TodoOK()) Then
        
        If TodoOK() Then
            strEstadoOrden = Estado_Orden_Ingresada
            strMensaje = "_____________________________________________________" & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & _
               Space$(8) & "<<<<< " & Trim$(UCase$(cboFondoOrden.Text)) & " >>>>>" & Chr$(vbKeyReturn) & _
               "_____________________________________________________" & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & _
               "Para proceder al Registro de la Orden Confirme lo siquiente : " & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & _
               "Fecha de Emisión          " & Space$(3) & ">" & Space$(2) & CStr(dtpFechaEmision.Value) & Chr$(vbKeyReturn) & _
               "Fecha de Vencimiento      " & Space$(3) & ">" & Space$(2) & CStr(dtpFechaVencimiento.Value) & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & _
               "Fecha de Operación        " & Space$(3) & ">" & Space$(2) & CStr(dtpFechaOrden.Value) & Chr$(vbKeyReturn) & _
               "Fecha de Liquidación      " & Space$(3) & ">" & Space$(2) & CStr(dtpFechaLiquidacion.Value) & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & _
               "Fecha de Pago             " & Space$(3) & ">" & Space$(2) & CStr(dtpFechaPago.Value) & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & _
               "Comisión                  " & Space$(3) & ">" & Space$(2) & txtComisionAgente(0).Text & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & _
               "Tasa                      " & Space$(3) & ">" & Space$(2) & txtTasa.Text & Chr$(vbKeyReturn) & _
               "Interés TOTAL calculado   " & Space$(3) & ">" & Space$(2) & CStr(CDbl(lblIntAdelantado(0).Caption) + CDbl(txtIntAdicional(0).Text)) & Chr$(vbKeyReturn) & _
               "Días por protesto         " & Space$(3) & ">" & Space$(2) & IIf(chkDiasAdicional.Visible = True And chkDiasAdicional.Value = Checked, DateDiff("d", dtpFechaVencimiento.Value, datFechaVctoAdicional), 0) & Chr$(vbKeyReturn) & _
               "Interés Prov. por Protesto" & Space$(3) & ">" & Space$(2) & txtIntAdicional(0).Text & Chr$(vbKeyReturn) & _
               "Cobro de Intereses        " & Space$(3) & ">" & Space$(2) & cboCobroInteres.Text & Chr$(vbKeyReturn) & _
               "Nominal                   " & Space$(3) & ">" & Space$(2) & txtValorNominal.Text & Chr$(vbKeyReturn) & _
               "Porcentaje de Descuento   " & Space$(3) & ">" & Space$(2) & txtPorcenDctoValorNominal.Text & Chr$(vbKeyReturn) & _
               "Valor Nominal Descontado  " & Space$(3) & ">" & Space$(2) & txtValorNominalDcto.Text & Chr$(vbKeyReturn) & _
               "Cantidad                  " & Space$(3) & ">" & Space$(2) & "1" & Chr$(vbKeyReturn) & _
               "Precio Unitario (%)       " & Space$(3) & ">" & Space$(2) & txtPrecioUnitario(0).Text & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & _
               "Monto Total Desembolsado  " & Space$(3) & ">" & Space$(2) & Trim$(lblDescripMoneda(0).Caption) & Space$(1) & lblMontoTotal(0).Caption & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & _
               "Tir Neta                  " & Space$(3) & ">" & Space$(2) & lblTirNeta.Caption & Chr$(vbKeyReturn) & _
               Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & _
               "¿ Seguro de continuar ?"

            If MsgBox(strMensaje, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                txtPorcenAgente(0).Text = 0#
                Me.Refresh: Exit Sub
            End If
        
            Me.MousePointer = vbHourglass
            
            strFechaOrden = Convertyyyymmdd(dtpFechaOrden.Value)
            strFechaLiquidacion = Convertyyyymmdd(dtpFechaLiquidacion.Value)
            strFechaEmision = Convertyyyymmdd(dtpFechaEmision.Value)
            strFechaVencimiento = Convertyyyymmdd(dtpFechaVencimiento.Value)
            strFechaPago = Convertyyyymmdd(dtpFechaPago.Value)
           
            Set adoRegistro = New ADODB.Recordset

            '*** Guardar Orden de Inversion ***
            With adoComm
                strIndTitulo = Valor_Caracter
                
                If strCodTipoOrden = Codigo_Orden_Pacto Then
                    strIndTitulo = Valor_Caracter
                    'En el sp de grabación de órdenes se cambiará el valor random asignado a la analítica para Depósitos a plazo.  ACC 13/01/2010
                    strCodAnalitica = NumAleatorio(8)
                    strCodTitulo = NumAleatorio(15)
                    strCodTipoTasa = Codigo_Tipo_Tasa_Efectiva
                    strCodBaseAnual = Codigo_Base_Actual_365
                    strCodRiesgo = "00" ' Sin Clasificacion
                    strCodReportado = Valor_Caracter
                    strCodFile = Left$(Trim$(lblAnalitica.Caption), 3)
                ElseIf (strCodTipoOrden = Codigo_Orden_Compra Or strCodTipoOrden = Codigo_Orden_Renovacion) Then

                    If strCodFile <> "003" And strCodFile <> "010" And strCodFile <> "014" And strCodFile <> "015" And strCodFile <> "016" Then
                        strCodAnalitica = ObtenerNuevaAnalitica(strCodFile)
                    End If   'caso contrario: strCodAnalitica Se obtiene en el sp de Grabación de la orden

                    strCodTitulo = NumAleatorio(15)
                        
                Else
                    strIndTitulo = Valor_Caracter
                    strCodTitulo = strCodGarantia
                    strCodGarantia = Valor_Caracter
                    strCodReportado = Valor_Caracter
                End If
                
                If strCalcVcto = "V" Then  'Con tasa de interés
                    dblTasaInteres = CDbl(txtTasa.Text)
                Else 'Con Precio
                    dblTasaInteres = 0#
                End If
                
                Dim dblMontoTotalMFL1 As Double
                dblMontoTotalMFL1 = CDec(lblMontoTotal(0).Caption)
          
                .CommandText = "{ call up_IVAdicInversionOrden('" & strCodFondoOrden & "','" & gstrCodAdministradora & "','','" & strFechaOrden & "','" & strCodTitulo & "','" & Trim$(txtNemonico.Text) & "','" & gstrPeriodoActual & "','" & gstrMesActual & "','','" & _
                   strEstadoOrden & "','" & strCodAnalitica & "','" & strCodFile & "','" & strCodAnalitica & "','" & strCodClaseInstrumento & "','" & strCodSubClaseInstrumento & "','01','" & _
                   strCodOperacion & "','" & strCodNegociacion & "','" & strCodOrigen & "','" & Trim$(txtDescripOrden.Text) & "','" & strCodEmisor & "','" & strCodAgente & "','" & strCodGarantia & "','" & strCodComisionista & "'," & numSecCondicion & ",'" & _
                   strFechaPago & "','" & strFechaVencimiento & "','" & strFechaLiquidacion & "','" & strFechaEmision & "','" & strCodMonedaDocumento & "'," & CDec(txtValorNominalDocumento.Text) & ",'" & _
                   strIndTipoCambio & "','" & strCodMoneda & "','" & strCodMoneda & "'," & CDec(txtValorNominalDcto.Text) & "," & CDec(txtTipoCambioDescuento.Text) & "," & CDec(txtTipoCambioDescuento.Text) & "," & _
                   txtValorNominal.Value & "," & txtPorcenDctoValorNominal.Value & "," & CDec(txtValorNominalDcto.Text) & ",1,1," & CDec(txtValorNominalDcto.Text) & "," & _
                   "0," & CDec(txtComisionAgente(0).Text) & ",0,0,0,0,0,0,0," & CDec(lblComisionIgvInt(0).Caption) & "," & dblMontoTotalMFL1 & ",0,0,0,0,0,0,0,0,0,0," & _
                   "0,0,0,0," & CDec(txtMontoVencimiento1.Value) & "," & CInt(txtDiasPlazo.Text) & ",'X','','','','','" & strCodReportado & "','" & strCodEmisor & "','" & _
                   strCodObligado & "','" & strCodObligado & "','" & strCodGestor & "','" & strCodFiador & "',0,'','X','" & strIndTitulo & "','" & strCodTipoTasa & "','" & strCodBaseAnual & _
                   "'," & CDec(dblTasaInteres) & ",'" & strPeriodoTasa & "','" & strIndCapitalizable & "','" & strPeriodoCapitalizable & "','" & strIndGeneraLetra & "'," & CDec(dblTasaInteres) & "," & _
                   CDec(dblTasaInteres) & "," & CDec(dblTasaInteres) & ",'" & strCodRiesgo & "','" & strCodSubRiesgo & "','" & Trim$(txtObservacion.Text) & "','" & gstrLogin & "','" & gstrFechaActual & "','" & _
                   gstrLogin & "','" & gstrFechaActual & "','" & strCodTitulo & "','" & strCodCobroInteres & "'," & CDec(lblIntAdelantado(0).Caption) & "," & CDec(txtCobroMinimoInteres.Text) & ",0,0," & CDec(txtDiasCobroMinimoInteres.Text) & ",0,'01'," & CDec(txtPorcenIgvInt(0).Text) & "," & _
                   CDec(lblComisionIgvInt(0).Caption) & ",0,0," & CDec(txtPorcenIgv(0).Text) & "," & CDec(lblComisionIgv(0).Caption) & ",0,0,0,0,0,0,'" & Trim$(txtNumAnexo.Text) & "','" & txtNumContrato.Text & "','" & _
                   Trim$(txtNumDocDscto.Text) & "','','" & strLineaCliente & "','" & Codigo_LimiteRE_Cliente & "','" & strCodPersonaLim & "','" & strTipoPersonaLim & "','" & strResponsablePago & "','" & _
                   strViaCobranza & "'," & CDec(txtTotalMNAnexo.Text) & "," & CDec(txtTotalDctosAnexo.Text) & "," & CDec(txtPorcenAgente(0).Text) & ",0) }"

                adoConn.Execute .CommandText
                
            End With
            
            Me.MousePointer = vbDefault
            
            MsgBox Mensaje_Adicion_Exitosa, vbExclamation
            adoComm.CommandText = "Update InstitucionPersonaParametroGeneral set ValorParametro = '" & CInt(txtNumAnexo.Text) & "' where CodPersona = '" & strCodEmisor & "' and CodParametro = '99'"
            adoConn.Execute adoComm.CommandText
            
            adoComm.CommandText = "Update InversionOrden set MontoTotalAnexo = '" & CDbl(txtTotalMNAnexo.Text) & _
                                "' where CodFondo = '" & strCodFondoOrden & "' and CodAdministradora = '" & gstrCodAdministradora & _
                                "' AND CodEmisor = '" & strCodEmisor & "' AND NumAnexo = '" & strNumAnexo & "'"
            adoConn.Execute adoComm.CommandText

            adoComm.CommandText = "Update InversionOperacion set MontoTotalAnexo = '" & CDbl(txtTotalMNAnexo.Text) & _
                                "' where CodFondo = '" & strCodFondoOrden & "' and CodAdministradora = '" & gstrCodAdministradora & _
                                "' AND CodEmisor = '" & strCodEmisor & "' AND NumAnexo = '" & strNumAnexo & "'"
            adoConn.Execute adoComm.CommandText

            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
            cmdOpcion.Visible = True

            With tabRFCortoPlazo
                .TabEnabled(0) = True
                .Tab = 0
                cmdOpcion.Visible = True
                cmdCancelar.Visible = False
                cmdGuardar.Visible = False
            End With

            Call Buscar
        End If
    End If

    Exit Sub
        
CtrlError:
  
    Me.MousePointer = vbDefault

    If Left$(err.Description, 14) <> "Excede Limites" Then
        strMsgError = "Error " & Str$(err.Number) & vbNewLine
        strMsgError = strMsgError & err.Description
        MsgBox strMsgError, vbCritical, "Error"
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

    Else
        strMsgError = strMsgError & err.Description
        MsgBox strMsgError, vbCritical, "Limites"
        
    End If
        
End Sub

Private Function TodoOK() As Boolean
        
    Dim adoRegistro   As ADODB.Recordset
    Dim strFechaDesde As String, strFechaHasta        As String
    
    TodoOK = False
          
    If cboTipoInstrumentoOrden.ListIndex <= 0 Then
        MsgBox "Debe seleccionar el Tipo de Instrumento de Corto Plazo.", vbCritical, Me.Caption

        If cboTipoInstrumentoOrden.Enabled Then cboTipoInstrumentoOrden.SetFocus
        Exit Function
    End If
    
    If cboClaseInstrumento.ListIndex <= 0 Then
        MsgBox "Debe seleccionar la Clase de Instrumento de Corto Plazo.", vbCritical, Me.Caption

        If cboClaseInstrumento.Enabled Then cboClaseInstrumento.SetFocus
        Exit Function
    End If

    If cboEmisor.ListIndex <= 0 Then
        MsgBox "Debe seleccionar el Emisor.", vbCritical, Me.Caption

        If cboEmisor.Enabled Then cboEmisor.SetFocus
        Exit Function
    End If
    
    Set adoRegistro = New ADODB.Recordset
    
    '*** Buscar en Títulos ***
    adoComm.CommandText = "SELECT Nemotecnico FROM InstrumentoInversion " & "WHERE CodFile='" & strCodFile & "' AND Nemotecnico='" & Trim$(txtNemonico.Text) & "' AND IndVigente='X'"
    Set adoRegistro = adoComm.Execute
    
    If Not adoRegistro.EOF Then
        MsgBox "Nemónico YA EXISTE...por favor verificar.", vbCritical, Me.Caption

        If txtNemonico.Enabled Then txtNemonico.SetFocus
        adoRegistro.Close: Set adoRegistro = Nothing
        Exit Function
    End If

    adoRegistro.Close
    
     If cboLineaCliente.ListIndex < 0 Then
        MsgBox "Debe seleccionar la Línea a afectar.", vbCritical, Me.Caption
        If cboLineaCliente.Enabled Then cboLineaCliente.SetFocus
        Exit Function
    End If
        

    strFechaDesde = Convertyyyymmdd(dtpFechaOrden.Value)
    strFechaHasta = Convertyyyymmdd(DateAdd("d", 1, dtpFechaOrden.Value))
    
    '*** Buscar en Ordenes del día ***
    adoComm.CommandText = "SELECT Nemotecnico FROM InversionOrden " & "WHERE (FechaOrden>='" & strFechaDesde & "' AND FechaOrden<'" & strFechaHasta & "') AND " & "CodFile='" & strCodFile & "' AND Nemotecnico='" & Trim$(txtNemonico.Text) & "' AND EstadoOrden<>'" & Estado_Orden_Anulada & "'"
    Set adoRegistro = adoComm.Execute
    
    If Not adoRegistro.EOF Then
        MsgBox "Nemónico YA EXISTE...por favor verificar.", vbCritical, Me.Caption

        If txtNemonico.Enabled Then txtNemonico.SetFocus
        adoRegistro.Close: Set adoRegistro = Nothing
        Exit Function
    End If

    adoRegistro.Close: Set adoRegistro = Nothing
        
'    If cboLineaCliente.ListIndex < 0 Then
'        MsgBox "Debe seleccionar la Línea a afectar.", vbCritical, Me.Caption
'
'        If cboLineaCliente.Enabled Then cboLineaCliente.SetFocus
'        Exit Function
'    End If
'
    If Trim$(txtNumDocDscto.Text) = Valor_Caracter Then
        MsgBox "Debe indicar el número del documento físico a descontar.", vbCritical, Me.Caption

        If txtNumDocDscto.Enabled Then txtNumDocDscto.SetFocus
        Exit Function
    End If
        
    If Trim$(txtDescripOrden.Text) = Valor_Caracter Then
        MsgBox "Debe indicar la Descripción de la ORDEN.", vbCritical, Me.Caption

        If txtDescripOrden.Enabled Then txtDescripOrden.SetFocus
        Exit Function
    End If
        
    If CVDate(dtpFechaEmision.Value) > CVDate(dtpFechaVencimiento.Value) Then
        MsgBox "La Fecha de Vencimiento debe ser mayor a la Fecha de Emisión.", vbCritical, Me.Caption

        If dtpFechaVencimiento.Enabled Then dtpFechaVencimiento.SetFocus
        Exit Function
    End If
    
    If CInt(txtDiasPlazo.Text) = 0 Then
        MsgBox "Debe indicar el número de días de plazo.", vbCritical, Me.Caption

        If txtDiasPlazo.Enabled Then txtDiasPlazo.SetFocus
        Exit Function
    End If
    
    If CDbl(txtTasa.Text) = 0 And strCalcVcto = "V" Then
        MsgBox "Debe indicar la Tasa Facial.", vbCritical, Me.Caption

        If txtTasa.Enabled Then txtTasa.SetFocus
        Exit Function
    End If
    
    If CCur(txtValorNominal.Text) = 0 Then
        MsgBox "Debe indicar el Valor Nominal.", vbCritical, Me.Caption

        If txtValorNominal.Enabled Then txtValorNominal.SetFocus
        Exit Function
    End If
    
    If CVDate(dtpFechaOrden.Value) > CVDate(dtpFechaLiquidacion.Value) Then
        MsgBox "La Fecha de Liquidación debe ser mayor o igual a la Fecha de la ORDEN.", vbCritical, Me.Caption

        If dtpFechaLiquidacion.Enabled Then dtpFechaLiquidacion.SetFocus
        Exit Function
    End If
            
    If CDbl(txtPrecioUnitario(0).Text) = 0 Then
        MsgBox "Debe indicar el Precio.", vbCritical, Me.Caption

        If txtPrecioUnitario(0).Enabled Then txtPrecioUnitario(0).SetFocus
        Exit Function
    End If

    '*** Si todo paso OK ***
    TodoOK = True
  
End Function

Public Sub SubImprimir(Index As Integer)

    Dim frmReporte As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    Dim strSeleccionRegistro As String

    If tabRFCortoPlazo.Tab = 1 Then Exit Sub
    
    Select Case Index

        Case 1
            gstrNameRepo = "InversionOrden"
            
            strSeleccionRegistro = "{InversionOrden.FechaOrden} IN 'Fch1' TO 'Fch2'"
            gstrSelFrml = strSeleccionRegistro
            frmRangoFecha.Show vbModal
                        
            If gstrSelFrml <> "0" Then
                Set frmReporte = New frmVisorReporte
    
                ReDim aReportParamS(5)
                ReDim aReportParamFn(5)
                ReDim aReportParamF(5)
                            
                aReportParamFn(0) = "Usuario"
                aReportParamFn(1) = "FechaDesde"
                aReportParamFn(2) = "FechaHasta"
                aReportParamFn(3) = "Hora"
                aReportParamFn(4) = "Fondo"
                aReportParamFn(5) = "NombreEmpresa"
                            
                aReportParamF(0) = gstrLogin
                aReportParamF(1) = Mid$(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10)
                aReportParamF(2) = Mid$(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10)
                aReportParamF(3) = Format$(Time$(), "hh:mm:ss")
                aReportParamF(4) = Trim$(cboFondo.Text)
                aReportParamF(5) = gstrNombreEmpresa & Space$(1)
                            
                aReportParamS(0) = strCodFondo
                aReportParamS(1) = gstrCodAdministradora
                aReportParamS(2) = Convertyyyymmdd(Mid$(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10))
                aReportParamS(3) = Convertyyyymmdd(DateAdd("d", 1, CVDate(Mid$(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10))))
                aReportParamS(4) = strCodMoneda
                aReportParamS(5) = strCodTipoInstrumento
            End If

        Case 2
            gstrNameRepo = "PapeletaInversion"
            
            strSeleccionRegistro = "{InversionOrden.FechaOrden} IN 'Fch1' TO 'Fch2'"
            gstrSelFrml = strSeleccionRegistro
            frmRangoFecha.Show vbModal
                        
            If gstrSelFrml <> "0" Then
                Set frmReporte = New frmVisorReporte
    
                ReDim aReportParamS(5)
                ReDim aReportParamFn(1)
                ReDim aReportParamF(1)
                            
                aReportParamFn(0) = "Fondo"
                aReportParamFn(1) = "NombreEmpresa"
                            
                aReportParamF(0) = Trim$(cboFondo.Text)
                aReportParamF(1) = gstrNombreEmpresa & Space$(1)
                            
                aReportParamS(0) = strCodFondo
                aReportParamS(1) = gstrCodAdministradora
                aReportParamS(2) = Convertyyyymmdd(Mid$(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10))
                aReportParamS(3) = Convertyyyymmdd(DateAdd("d", 1, CVDate(Mid$(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10))))
                aReportParamS(4) = strCodMoneda
                aReportParamS(5) = strCodTipoInstrumento
            End If
            
        Case 3, 4

            If Index = 3 Then
                gstrNameRepo = "Anexo"
            Else
                gstrNameRepo = "AnexoCliente"
            End If
            
            Set frmReporte = New frmVisorReporte
    
            ReDim aReportParamS(9)
            ReDim aReportParamFn(3)
            ReDim aReportParamF(3)
                
            aReportParamFn(0) = "Usuario"
            aReportParamFn(1) = "Hora"
            aReportParamFn(2) = "Fondo"
            aReportParamFn(3) = "NombreEmpresa"
                            
            aReportParamF(0) = gstrLogin
            aReportParamF(1) = Format$(Time$(), "hh:mm:ss")
            aReportParamF(2) = Trim$(cboFondo.Text)
            aReportParamF(3) = gstrNombreEmpresa & Space$(1)

            aReportParamS(0) = strCodFondo                              'CodFondo
            aReportParamS(1) = gstrCodAdministradora                    'CodAdministradora
            aReportParamS(2) = tdgConsulta.Columns("CodFile")           'CodFile
            aReportParamS(3) = tdgConsulta.Columns("CodDetalleFile")    'CodDetalleFile
            aReportParamS(4) = tdgConsulta.Columns("CodSubDetalleFile") 'Subdetallefile'
            aReportParamS(5) = tdgConsulta.Columns("CodLimiteCli")      'CodLimiteCli'
            aReportParamS(6) = tdgConsulta.Columns("CodEstructura")     'CodEstructura
            aReportParamS(7) = tdgConsulta.Columns("CodPersonaLim")     'CodPersona
            aReportParamS(8) = tdgConsulta.Columns("TipoPersonaLim")    'TipoPersona
            aReportParamS(9) = tdgConsulta.Columns("NumAnexo")          'Nro de Anexo
            
            'End If
            
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

Public Sub Salir()

    Unload Me
    
End Sub

Private Sub cboComisionista_Click()
    
    strCodComisionista = Valor_Caracter
    numSecCondicion = 0
    
    If cboComisionista.ListIndex < 0 Then Exit Sub
    
    strCodComisionista = Mid$(arrComisionista(cboComisionista.ListIndex), 1, 8)
    numSecCondicion = Mid$(arrComisionista(cboComisionista.ListIndex), 9)

End Sub

Private Sub cboLineaCliente_Click()

    Dim adoRegistro As New ADODB.Recordset

    strLineaCliente = Valor_Caracter
    strTipoPersonaLim = Valor_Caracter
    strCodPersonaLim = Valor_Caracter
    
    If cboLineaCliente.ListIndex < 0 Then Exit Sub
    
    strLineaCliente = Trim(arrLineaCliente(cboLineaCliente.ListIndex))
    
    strTipoPersonaLim = Codigo_Tipo_Persona_Emisor
    If strLineaCliente = Linea_Financiamiento_Proveedores Then
        strCodPersonaLim = strCodObligado
        strTipoPersonaLim = Codigo_Tipo_Persona_Obligado
    Else
        strCodPersonaLim = strCodEmisor
        strTipoPersonaLim = Codigo_Tipo_Persona_Emisor
    End If
    
End Sub

Private Sub cboLineaClienteLista_Click()

    strLineaClienteLista = Valor_Caracter
    If cboLineaClienteLista.ListIndex < 0 Then Exit Sub
    
    strLineaClienteLista = Trim(arrLineaClienteLista(cboLineaClienteLista.ListIndex))

    Call Buscar

End Sub
Private Sub cbxGenerarLetra_Click()

    If (cbxGenerarLetra.Value = vbChecked) Then
        strIndGeneraLetra = Valor_Indicador
    Else
        strIndGeneraLetra = Valor_Caracter
    End If

End Sub

Private Sub cboMonedaDocumento_Click()
    strCodMonedaDocumento = Valor_Caracter
    strCodMoneda = Valor_Caracter
    
    If cboMonedaDocumento.ListIndex < 0 Then Exit Sub
    
    strCodMonedaDocumento = Trim$(arrMonedaDocumento(cboMonedaDocumento.ListIndex))
    strCodMoneda = Trim$(arrMoneda(cboMoneda.ListIndex))

    If (strCodMonedaDocumento <> strCodMoneda) Then
        lblDescrip.Item(185).Visible = True
        txtTipoCambioDescuento.Visible = True
        strIndTipoCambio = Valor_Indicador
    Else
        lblDescrip.Item(185).Visible = False
        txtTipoCambioDescuento.Visible = False
        txtTipoCambioDescuento.Text = "1.00000"
        strIndTipoCambio = Valor_Caracter
    End If
    
End Sub

Private Sub cboPeriodoTasa_Click()
    strPeriodoTasa = Valor_Caracter

    If cboPeriodoTasa.ListIndex < 0 Then Exit Sub
    
    strPeriodoTasa = Trim$(arrPeriodoTasa(cboPeriodoTasa.ListIndex))
    
    lblIntAdelantado(0).Caption = CalculoInteresDescuento(CDbl(txtTasa.Text), strCodTipoTasa, strPeriodoTasa, strCodBaseAnual, txtValorNominalDcto, CStr(dtpFechaOrden.Value), CStr(dtpFechaPago.Value))
    txtCobroMinimoInteres.Text = CalculoInteresDescuento(CDbl(txtTasa.Text), strCodTipoTasa, strPeriodoTasa, strCodBaseAnual, txtValorNominalDcto, CStr(dtpFechaOrden.Value), CStr(DateAdd("d", intDiasInteresMinimo, dtpFechaOrden.Value)))
    Call CalculoTotal(0)
End Sub

Private Sub cboPeriodoCapitalizacion_Click()
    strPeriodoCapitalizable = Valor_Caracter

    If cboPeriodoCapitalizacion.ListIndex < 0 Then Exit Sub
    
    strPeriodoCapitalizable = Trim$(arrPeriodoCapitalizacion(cboPeriodoCapitalizacion.ListIndex))
    lblIntAdelantado(0).Caption = CalculoInteresDescuento(CDbl(txtTasa.Text), strCodTipoTasa, strPeriodoTasa, strCodBaseAnual, txtValorNominalDcto, CStr(dtpFechaOrden.Value), CStr(dtpFechaPago.Value))
    txtCobroMinimoInteres.Text = CalculoInteresDescuento(CDbl(txtTasa.Text), strCodTipoTasa, strPeriodoTasa, strCodBaseAnual, txtValorNominalDcto, CStr(dtpFechaOrden.Value), CStr(DateAdd("d", intDiasInteresMinimo, dtpFechaOrden.Value)))
    Call CalculoTotal(0)

End Sub

Private Sub chkCapitalizable_Click()

    If strCodTipoTasa = Codigo_Tipo_Tasa_Nominal Then
        lblPeriodoCapitalizacion.Enabled = True
        cboPeriodoCapitalizacion.Enabled = True
        strIndCapitalizable = Valor_Indicador
    Else
        lblPeriodoCapitalizacion.Enabled = False
        cboPeriodoCapitalizacion.Enabled = False
        strIndCapitalizable = Valor_Caracter
    End If
    
End Sub

Private Sub cboBaseAnual_Click()

    strCodBaseAnual = Valor_Caracter

    If cboBaseAnual.ListIndex < 0 Then Exit Sub
    
    strCodBaseAnual = Trim$(arrBaseAnual(cboBaseAnual.ListIndex))
    
    '*** Base de Cálculo ***
    intBaseCalculo = 365

    Select Case strCodBaseAnual

        Case Codigo_Base_30_360: intBaseCalculo = 360

        Case Codigo_Base_Actual_365: intBaseCalculo = 365

        Case Codigo_Base_Actual_360: intBaseCalculo = 360

        Case Codigo_Base_30_365: intBaseCalculo = 365
    End Select
    
    txtValorNominal_Change
    lblIntAdelantado(0).Caption = CalculoInteresDescuento(CDbl(txtTasa.Text), strCodTipoTasa, strPeriodoTasa, strCodBaseAnual, txtValorNominalDcto, CStr(dtpFechaOrden.Value), CStr(dtpFechaPago.Value))
    txtCobroMinimoInteres.Text = CalculoInteresDescuento(CDbl(txtTasa.Text), strCodTipoTasa, strPeriodoTasa, strCodBaseAnual, txtValorNominalDcto, CStr(dtpFechaOrden.Value), CStr(DateAdd("d", intDiasInteresMinimo, dtpFechaOrden.Value)))
    
End Sub

Private Sub cboClaseInstrumento_Click()

    strCodClaseInstrumento = Valor_Caracter

    If cboClaseInstrumento.ListIndex < 0 Then Exit Sub
    
    strCodClaseInstrumento = Trim$(arrClaseInstrumento(cboClaseInstrumento.ListIndex))
        
    strSQL = "SELECT CodLimite CODIGO, DescripLimite DESCRIP FROM LimiteReglamentoEstructuraDetalle where CodEstructura = '" & Codigo_LimiteRE_Cliente & "'" & _
             " and CodLimite = '" & Linea_Descuento_Letras_Facturas & "' and Estado  = '01' "
    CargarControlLista strSQL, cboLineaCliente, arrLineaCliente(), ""
    Call cboLineaCliente_Click       'Para obligar a que se seleccione el único elemento de la lista
    
    If cboLineaCliente.ListCount > 0 Then cboLineaCliente.ListIndex = 0
    
    strSQL = "SELECT CodSubDetalleFile CODIGO,DescripSubDetalleFile DESCRIP FROM InversionSubDetalleFile WHERE " & "CodDetalleFile='" & strCodClaseInstrumento & "' AND CodFile='" & strCodTipoInstrumentoOrden & "' AND IndVigente='X' ORDER BY CodSubDetalleFile"
        
    CargarControlLista strSQL, cboSubClaseInstrumento, arrSubClaseInstrumento(), Sel_Defecto
    
    If cboSubClaseInstrumento.ListCount > 1 Then
        cboSubClaseInstrumento.ListIndex = 1
    Else

        If cboSubClaseInstrumento.ListCount > 0 Then cboSubClaseInstrumento.ListIndex = 0
    End If
    
    cboSubClaseInstrumento.Enabled = True

    If strCodClaseInstrumento = "001" Then strCalcVcto = "V"   'tasa de interés
    If strCodClaseInstrumento = "002" Then strCalcVcto = "D"    'Al descuento

    txtNemonico.Text = GenerarNemonico(strCodTipoInstrumentoOrden, strCodClaseInstrumento, strCodEmisor, txtNumDocDscto.Text)

End Sub

Private Function GenerarNemonico(strTipoInstrumento As String, _
                                 strClaseOperacion As String, _
                                 strCodEmisor As String, _
                                 strNumDocumento) As String

    Dim adoTemporal       As ADODB.Recordset
    Dim strNemotecnico    As String
    Dim strValorParametro As String
    Dim strCodParametro   As String

    GenerarNemonico = Valor_Caracter

    If Trim$(strTipoInstrumento) = "" Or Trim$(strClaseOperacion) = "" Or Trim$(strCodEmisor) = "" Or Trim$(strNumDocumento) = "" Then
        Exit Function
    End If
    
    Set adoTemporal = New ADODB.Recordset

    With adoComm
        .CommandText = "SELECT DescripNemonico FROM InstitucionPersona WHERE CodPersona='" & strCodEmisor & "' AND TipoPersona = '" & Codigo_Tipo_Persona_Emisor & "'"
        
        Set adoTemporal = .Execute

        If Not adoTemporal.EOF Then
            strNemotecnico = Trim$(adoTemporal("DescripNemonico"))

            If strNemotecnico = "" Then
                adoTemporal.Close: Set adoTemporal = Nothing
                Exit Function
            End If

        Else
            adoTemporal.Close: Set adoTemporal = Nothing
            Exit Function
        End If
      
        Select Case strTipoInstrumento
        
            Case CodFile_Descuento_Comprobantes_Pago   '"014"

                Select Case strClaseOperacion

                    Case "001"   'Factura
                        strCodParametro = "02"

                    Case "002"   'Recibos por honorarios
                        strCodParametro = "05"

                    Case "003"   'Coleta de venta
                        strCodParametro = "06"
                End Select
                
            Case CodFile_Descuento_Documentos_Cambiario  '"015"

                Select Case strClaseOperacion

                    Case "001"  'Letra
                        strCodParametro = "01"

                    Case "002"  'Pagaré
                        strCodParametro = "08"

                    Case "003"  'Cheque
                        strCodParametro = "07"
                End Select
            
            Case CodFile_Descuento_Flujos_Dinerarios   '"016"

                Select Case strClaseOperacion

                    Case "001"  'Contratos
                        strCodParametro = "03"
                End Select
            
            Case "010" 'Letras
                strCodParametro = "04"
            
        End Select
        
        .CommandText = "SELECT ValorParametro FROM AuxiliarParametro WHERE CodParametro='" & strCodParametro & "' AND CodTipoParametro = 'NEMODF' "
        Set adoTemporal = .Execute

        If Not adoTemporal.EOF Then
            strValorParametro = Trim$(adoTemporal("ValorParametro"))
        End If
        
        adoTemporal.Close: Set adoTemporal = Nothing
        
    End With

    GenerarNemonico = Trim$(strValorParametro) & Trim$(strNemotecnico) & "-" & Trim$(strNumDocumento)

End Function

'Private Sub cboConceptoCosto_Click()
'
'    Dim adoRegistro As ADODB.Recordset
'
'    strCodConcepto = Valor_Caracter
'
'    If cboConceptoCosto.ListIndex < 0 Then Exit Sub
'
'    strCodConcepto = Trim$(arrConceptoCosto(cboConceptoCosto.ListIndex))
'
'    strCodTipoCostoBolsa = Valor_Caracter
'
'    dblComisionBolsa = 0
'
'    With adoComm
'        Set adoRegistro = New ADODB.Recordset
'
'        .CommandText = "SELECT CodCosto,TipoCosto,ValorCosto FROM CostoNegociacion WHERE TipoOperacion='" & strCodConcepto & "' AND TipoValor='" & Codigo_Valor_RentaFija & "' ORDER BY CodCosto"
'        Set adoRegistro = .Execute
'
'        Do Until adoRegistro.EOF
'
'            Select Case Trim$(adoRegistro("CodCosto"))
'
'                Case Codigo_Costo_Bolsa
'                    strCodTipoCostoBolsa = Trim$(adoRegistro("TipoCosto"))
'                    dblComisionBolsa = CDbl(adoRegistro("ValorCosto"))
'
'            End Select
'
'            adoRegistro.MoveNext
'        Loop
'
'        adoRegistro.Close: Set adoRegistro = Nothing
'    End With
'
'End Sub

Private Sub ObtenerParametrosGeneralesEmisor()
    Dim adoRegistro As ADODB.Recordset
    
    With adoComm
        .CommandText = "SELECT CodParametro, CodSubParametro, TipoValor, ValorParametro from InstitucionPersonaParametroGeneral " & _
                        "where CodPersona = '" & strCodEmisor & "' and Estado = '01'"
        Set adoRegistro = .Execute
    End With
    
    dblTasaInteres = 0
    dblPorcentajeComision = 0
    dblMontoMinComisiones = 0
    intUltNumAnexo = 0
    txtNumContrato.Text = "0"
    txtPorcenDctoValorNominal.Text = "85.000000"

    While Not adoRegistro.EOF
        Select Case adoRegistro("CodParametro")
            Case "00" 'Numero de Contrato
                txtNumContrato.Text = Trim$(adoRegistro("ValorParametro"))
            Case "01" 'Tasa de interes
                dblTasaInteres = CDbl(adoRegistro("ValorParametro"))
            Case "02" 'Porcentaje de comision
                dblPorcentajeComision = CDbl(adoRegistro("ValorParametro"))
            Case "03" 'MontoMinimoComision
                If adoRegistro("CodSubParametro") = strCodMoneda Then
                    dblMontoMinComisiones = CDbl(adoRegistro("ValorParametro"))
                End If
            Case "04" 'Porcentaje de Descuento del Valor Nominal
                txtPorcenDctoValorNominal.Text = adoRegistro("ValorParametro")
            Case "99" 'UltNumAnexo
                intUltNumAnexo = CInt(adoRegistro("ValorParametro"))
        End Select
        adoRegistro.MoveNext
    Wend
    
    txtTasa.Text = dblTasaInteres
    If Not blnCargarCabeceraAnexo Then
        txtNumAnexo.Text = Format(intUltNumAnexo + 1, "0000000000")
    Else
        txtNumAnexo.Text = strNumAnexo
    End If
    strNumAnexo = txtNumAnexo.Text
    
End Sub

Private Sub cboEmisor_Click()
    blnEmisorReady = False
    Dim adoRegistro As ADODB.Recordset
    
    strCodTitulo = Valor_Caracter
    strCodEmisor = Valor_Caracter: strCodAnalitica = Valor_Caracter
    lblAnalitica = strCodTipoInstrumentoOrden & "-??????": 'txtValorNominal.Text = "1"
    
    If cboEmisor.ListIndex < 0 Then
        blnEmisorReady = True
        Exit Sub
    End If

    strCodEmisor = Left$(Trim$(arrEmisor(cboEmisor.ListIndex)), 8)
        
    cboLineaCliente_Click
    
    'Asignando el nemónico
    txtNemonico.Text = GenerarNemonico(strCodTipoInstrumentoOrden, strCodClaseInstrumento, strCodEmisor, txtNumDocDscto.Text)
    
    'Obteniendo los parámetros generales del EMISOR
    Call ObtenerParametrosGeneralesEmisor
    
    'Obtener lista de comisionistas
    If strCodEmisor <> Valor_Caracter Then
        strSQL = "{ call up_ACLstFondoComisionistaContraparte('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                    Codigo_Tipo_Comisionista_Inversion & "','" & Codigo_Tipo_Persona_Emisor & "','" & strCodEmisor & "','" & _
                    strCodMoneda & "','" & gstrFechaActual & "') }"
        CargarControlLista strSQL, cboComisionista, arrComisionista(), Valor_Caracter
    Else
        cboComisionista.Clear
    End If
    
    If cboComisionista.ListCount = 1 Then
        cboComisionista.ListIndex = 0
    End If
    
    '*** Validar Limites ***
    If strCodTipoInstrumentoOrden = Valor_Caracter Then
        blnEmisorReady = True
        Exit Sub
    End If
    
    If Not PosicionLimites() Then
        blnEmisorReady = True
        Exit Sub
    End If

    If blnCancelaPrepago = False Then
        strCodTitulo = strCodFondoOrden & strCodFile & strCodAnalitica
    End If
    
    With adoComm
        Set adoRegistro = New ADODB.Recordset
                        
        '*** Categoría del instrumento emitido por el emisor ***
        .CommandText = "SELECT CodCategoriaRiesgo,CodRiesgoFinal,CodSubRiesgoFinal FROM EmisionInstitucionPersona " & "WHERE CodEmisor='" & strCodEmisor & "' AND CodFile='" & strCodTipoInstrumentoOrden & "' AND " & "CodDetalleFile='" & strCodClaseInstrumento & "'"
            
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            strCodRiesgo = Trim$(adoRegistro("CodRiesgoFinal"))
            strCodSubRiesgo = Trim$(adoRegistro("CodSubRiesgoFinal"))
        Else

            If strCodEmisor <> Valor_Caracter Then
                'MsgBox "La Clasificación de Riesgo no está definida...", vbCritical, Me.Caption
                cboLineaCliente_Click
                blnEmisorReady = True
                Exit Sub
            End If
        End If

        adoRegistro.Close
        
        '*** Obtener el Riesgo ***
        .CommandText = "SELECT ValorParametro FROM AuxiliarParametro WHERE CodTipoParametro='TIPRIE' AND CodParametro='" & strCodRiesgo & "'"
        Set adoRegistro = .Execute
        
        adoRegistro.Close: Set adoRegistro = Nothing
        
    End With
    blnEmisorReady = True
End Sub

Private Function PosicionLimites() As Boolean

    PosicionLimites = False
        
    If cboTipoInstrumentoOrden.ListIndex <= 0 Then
        MsgBox "Debe seleccionar el Tipo de Instrumento.", vbCritical, Me.Caption
        cboEmisor.ListIndex = -1 ': cboTitulo.ListIndex = -1

        If cboTipoInstrumentoOrden.Enabled Then cboTipoInstrumentoOrden.SetFocus
        Exit Function
    End If

    '*** Si todo pasó OK ***
    PosicionLimites = True
    
End Function

Private Sub cboEstado_Click()

    strCodEstado = Valor_Caracter

    If cboEstado.ListIndex < 0 Then Exit Sub
    
    strCodEstado = Trim$(arrEstado(cboEstado.ListIndex))
    
    Call Buscar
End Sub

Private Sub cboFondo_Click()
        
    On Error GoTo cboFondo_Click_Err
    
    Dim adoRegistro As ADODB.Recordset
    
    strCodFondo = Valor_Caracter
    cboFondoOrden.ListIndex = cboFondo.ListIndex

    If cboFondo.ListIndex < 0 Then Exit Sub
    strCodFondo = Trim$(arrFondo(cboFondo.ListIndex))
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Fecha Vigente, Moneda ***
        .CommandText = "{ call up_ACSelDatosParametro(23,'" & strCodFondo & "','" & gstrCodAdministradora & "','000') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            gdatFechaActual = CVDate(adoRegistro("FechaCuota"))
            dtpFechaOrdenDesde.Value = gdatFechaActual
            dtpFechaOrdenHasta.Value = dtpFechaOrdenDesde.Value
            strCodMoneda = Trim$(adoRegistro("CodMoneda"))
            
            frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
        End If

        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    '*** Tipo de Instrumento ***
    strSQL = "SELECT FIF.CodFile CODIGO,DescripFile DESCRIP " & "FROM FondoInversionFile FIF JOIN InversionFile IVF ON(IVF.CodFile=FIF.CodFile) " & "WHERE TipoValor='" & Codigo_Valor_RentaFija & "' AND TipoPlazo='" & Codigo_Valor_CortoPlazo & "' AND IndInstrumento='X' AND IndVigente='X' AND " & "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND FIF.CodFile <> '" & CodFile_Descuento_Flujos_Dinerarios & "' ORDER BY DescripFile"
    CargarControlLista strSQL, cboTipoInstrumento, arrTipoInstrumento(), Sel_Todos
    
    If cboTipoInstrumento.ListCount > 0 Then cboTipoInstrumento.ListIndex = 0
    
    Exit Sub

cboFondo_Click_Err:
    MsgBox err.Description & vbCrLf & "in Inversion.frmDesembolsoAcreencias.cboFondo_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"
    Resume Next
  
End Sub

Private Sub cboFondoOrden_Click()

    Dim adoRegistro As ADODB.Recordset
    
    strCodFondoOrden = Valor_Caracter

    If cboFondoOrden.ListIndex < 0 Then Exit Sub
    
    cboFondo.ListIndex = cboFondoOrden.ListIndex
    
    strCodFondoOrden = Trim$(arrFondoOrden(cboFondoOrden.ListIndex))

    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Fecha Vigente, Moneda, Tipo de Cambio ***
        .CommandText = "{ call up_ACSelDatosParametro(23,'" & strCodFondoOrden & "','" & gstrCodAdministradora & "','000') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            gdatFechaActual = CVDate(adoRegistro("FechaCuota"))
            gstrFechaActual = Convertyyyymmdd(gdatFechaActual)
            dtpFechaOrden.Value = gdatFechaActual
            dtpFechaLiquidacion.Value = dtpFechaOrden.Value
            dtpFechaEmision.Value = dtpFechaOrden.Value
            dtpFechaVencimiento.Value = dtpFechaEmision.Value
            strCodMoneda = Trim$(adoRegistro("CodMoneda"))
            txtTipoCambio.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, dtpFechaOrden.Value, Codigo_Moneda_Local, strCodMoneda))

            If CDbl(txtTipoCambio.Text) = 0 Then txtTipoCambio.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, DateAdd("d", -1, dtpFechaOrden.Value), Codigo_Moneda_Local, strCodMoneda))
            dblTipoCambio = CDbl(txtTipoCambio.Text)
                        
            frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
            Me.Refresh
        End If

        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    '*** Tipo de Instrumento ***
    strSQL = "SELECT FIF.CodFile CODIGO,DescripFile DESCRIP " & "FROM FondoInversionFile FIF JOIN InversionFile IVF ON(IVF.CodFile=FIF.CodFile) " & "WHERE TipoValor='" & Codigo_Valor_RentaFija & "' AND TipoPlazo='" & Codigo_Valor_CortoPlazo & "' AND IndInstrumento='X' AND IndVigente='X' AND " & "CodFondo='" & strCodFondoOrden & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND FIF.CodFile <> '" & CodFile_Descuento_Flujos_Dinerarios & "' ORDER BY CODIGO"
    CargarControlLista strSQL, cboTipoInstrumentoOrden, arrTipoInstrumentoOrden(), Sel_Defecto
        
    cboTipoInstrumentoOrden.ListIndex = 0
  
End Sub

Private Sub cboGestor_Click()
    
    strCodGestor = Valor_Caracter

    If cboGestor.ListIndex < 0 Then Exit Sub
    
    strCodGestor = Trim$(arrGestor(cboGestor.ListIndex))

End Sub

Private Sub cboNegociacion_Click()

    strCodNegociacion = Valor_Caracter

    If cboNegociacion.ListIndex < 0 Then Exit Sub
    
    strCodNegociacion = Trim$(arrNegociacion(cboNegociacion.ListIndex))
            
End Sub

Private Sub cboObligado_Click()

    strCodObligado = Valor_Caracter

    If cboObligado.ListIndex < 0 Then Exit Sub
    
    strCodObligado = Trim$(arrObligado(cboObligado.ListIndex))
    

    
 
End Sub

Private Sub cboCobroInteres_Click()

    strCodCobroInteres = Valor_Caracter

    If cboCobroInteres.ListIndex < 0 Then Exit Sub

    strCodCobroInteres = Mid$(Trim$(arrPagoInteres(cboCobroInteres.ListIndex)), 7, 2)

    If strCodCobroInteres = Codigo_Modalidad_Pago_Adelantado Then   'Si es pago de intereses adelantados permitir la edición de int. adicionales
        If (strCodTipoInstrumentoOrden = "015" And strCodClaseInstrumento = "001" And strCodSubClaseInstrumento = "001") Then   'Sòlo en caso de letras
        
            txtIntAdicional(0).Enabled = True
            chkDiasAdicional.Visible = True
            chkDiasAdicional.Value = Checked
        
            'Los días adicionales se suman a la fecha de vencimieno del documento
            datFechaVctoAdicional = DateAdd("d", intDiasAdicionales, CVDate(dtpFechaVencimiento.Value))

            If Not EsDiaUtil(datFechaVctoAdicional) Then
                datFechaVctoAdicional = ProximoDiaUtil(datFechaVctoAdicional)
            End If

            lblFechaVencimientoAdic.Caption = datFechaVctoAdicional
            lblDiasAdic(0).Caption = "( " & CStr(DateDiff("d", dtpFechaVencimiento.Value, datFechaVctoAdicional)) & " días )"
            lblFechaVencimientoAdic.Visible = True
           
            txtDiasPlazo.Text = CStr(DateDiff("d", dtpFechaOrden.Value, dtpFechaVencimiento.Value))
            lblDiasPlazo.Caption = txtDiasPlazo.Text
            Call CalculoTotal(0)
            
        End If

        If (strCodTipoInstrumentoOrden = "014" And strCodClaseInstrumento = "001" And strCodSubClaseInstrumento = "001") Then   'Sòlo en caso de facturas
        
            txtIntAdicional(0).Enabled = True
            chkDiasAdicional.Visible = True
            chkDiasAdicional.Value = Checked
        
            'Los días adicionales se suman a la fecha de vencimieno del documento
            datFechaVctoAdicional = DateAdd("d", intDiasAdicionales, CVDate(dtpFechaVencimiento.Value))

            If Not EsDiaUtil(datFechaVctoAdicional) Then
                datFechaVctoAdicional = ProximoDiaUtil(datFechaVctoAdicional)
            End If
            
            lblFechaVencimientoAdic.Caption = datFechaVctoAdicional
            lblDiasAdic(0).Caption = "( " & CStr(DateDiff("d", dtpFechaVencimiento.Value, datFechaVctoAdicional)) & " días )"
            lblFechaVencimientoAdic.Visible = True
           
            txtDiasPlazo.Text = CStr(DateDiff("d", dtpFechaOrden.Value, dtpFechaVencimiento.Value))
            lblDiasPlazo.Caption = txtDiasPlazo.Text
            Call CalculoTotal(0)
                        
        End If

    Else

        If (strCodTipoInstrumentoOrden = "015" And strCodClaseInstrumento = "001" And strCodSubClaseInstrumento = "001") Then   'Sòlo en caso de letras
            
            txtIntAdicional(0).Text = 0#
            txtIntAdicional(0).Enabled = False
            chkDiasAdicional.Value = Unchecked
            
            datFechaVctoAdicional = dtpFechaVencimiento.Value
            lblFechaVencimientoAdic.Caption = datFechaVctoAdicional
            lblFechaVencimientoAdic.Visible = False
            lblDiasAdic(0).Caption = 0
           
            txtDiasPlazo.Text = CStr(DateDiff("d", dtpFechaOrden.Value, dtpFechaVencimiento.Value))
            lblDiasPlazo.Caption = txtDiasPlazo.Text
            Call CalculoTotal(0)
        
        End If

        If (strCodTipoInstrumentoOrden = "014" And strCodClaseInstrumento = "001" And strCodSubClaseInstrumento = "001") Then   'caso de facturas
            
            txtIntAdicional(0).Text = 0#
            txtIntAdicional(0).Enabled = False
            chkDiasAdicional.Value = Unchecked
            
            datFechaVctoAdicional = dtpFechaVencimiento.Value
            lblFechaVencimientoAdic.Caption = datFechaVctoAdicional
            lblFechaVencimientoAdic.Visible = False
            lblDiasAdic(0).Caption = 0
           
            txtDiasPlazo.Text = CStr(DateDiff("d", dtpFechaOrden.Value, dtpFechaVencimiento.Value))
            lblDiasPlazo.Caption = txtDiasPlazo.Text
            Call CalculoTotal(0)
            lblIntAdelantado(0).Caption = 0
            txtCobroMinimoInteres.Text = "0"

            lblComisionIgvInt(0).Caption = 0
            
        End If
    End If
            
End Sub

Private Sub cboResponsablePago_Click()

    strResponsablePago = Valor_Caracter

    If cboResponsablePago.ListIndex < 0 Then Exit Sub

    strResponsablePago = Trim$(arrResponsablePago(cboResponsablePago.ListIndex))

End Sub

Private Sub cboSubClaseInstrumento_Click()

    Dim adoRegistro As ADODB.Recordset          'ACC 12/03/2010  Agregado

    strCodSubClaseInstrumento = Valor_Caracter

    If cboSubClaseInstrumento.ListIndex < 0 Then Exit Sub
    
    strCodSubClaseInstrumento = Trim$(arrSubClaseInstrumento(cboSubClaseInstrumento.ListIndex))
    
    intDiasAdicionales = 0

    If strCodTipoInstrumentoOrden = "015" And strCodClaseInstrumento = "001" And strCodSubClaseInstrumento = "001" Then

        chkDiasAdicional.Visible = True
        chkDiasAdicional.Value = Checked
        chkDiasAdicional_Click
        cboCobroInteres.Enabled = True
    
        'Obteniendo los dìas adicionales
        Set adoRegistro = New ADODB.Recordset
        adoComm.CommandText = "SELECT CONVERT(int,ValorParametro) AS DiasAdicionales FROM ParametroGeneral WHERE CodParametro = '24'"
        Set adoRegistro = adoComm.Execute

        If Not (adoRegistro.EOF) Then
            intDiasAdicionales = adoRegistro("DiasAdicionales")
        End If

        If intDiasAdicionales = Null Then
            intDiasAdicionales = 0
        End If

        adoRegistro.Close: Set adoRegistro = Nothing
    
    Else

        If strCodTipoInstrumentoOrden <> "010" Then
            cboCobroInteres.ListIndex = ObtenerItemLista(arrPagoInteres(), "MODPAG" + Codigo_Modalidad_Pago_Adelantado)
            chkDiasAdicional.Value = Unchecked
            chkDiasAdicional.Visible = False
        End If
    End If
    
    Call cboTipoOrden_Click
End Sub

Private Sub cboTipoInstrumento_Click()

    strCodTipoInstrumento = Valor_Caracter

    If cboTipoInstrumento.ListIndex < 0 Then Exit Sub
    
    strCodTipoInstrumento = Trim$(arrTipoInstrumento(cboTipoInstrumento.ListIndex))
    
    strSQL = "SELECT CodLimite CODIGO, DescripLimite DESCRIP FROM LimiteReglamentoEstructuraDetalle where CodEstructura = '" & Codigo_LimiteRE_Cliente & "'" & " and CodLimite = '" & Linea_Descuento_Letras_Facturas & "' and Estado  = '01' "
    CargarControlLista strSQL, cboLineaClienteLista, arrLineaClienteLista(), ""

    If cboLineaClienteLista.ListCount > 0 Then cboLineaClienteLista.ListIndex = 0
    
    Call Buscar
    
End Sub

Private Sub cboTipoInstrumentoOrden_Click()
    
    strCodTipoInstrumentoOrden = Valor_Caracter

    If cboTipoInstrumentoOrden.ListIndex < 0 Then Exit Sub
    
    strCodTipoInstrumentoOrden = Trim$(arrTipoInstrumentoOrden(cboTipoInstrumentoOrden.ListIndex))

    If strCodTipoInstrumentoOrden = "010" Then   'Letras
        cboCobroInteres.ListIndex = ObtenerItemLista(arrPagoInteres(), "MODPAG" + Codigo_Modalidad_Pago_Vencimiento)
    Else
        cboCobroInteres.ListIndex = ObtenerItemLista(arrPagoInteres(), "MODPAG" + Codigo_Modalidad_Pago_Adelantado)
        cboCobroInteres.Enabled = True
    End If
    
    'Asignar nemónico
    txtNemonico.Text = GenerarNemonico(strCodTipoInstrumentoOrden, strCodClaseInstrumento, strCodEmisor, txtNumDocDscto.Text)
    
    strSQL = "SELECT IFTON.CodTipoOperacion CODIGO,DescripParametro DESCRIP " & "FROM InversionFileTipoOperacionNegociacion IFTON JOIN AuxiliarParametro AUX ON(AUX.CodParametro=IFTON.CodTipoOperacion AND AUX.CodTipoParametro = 'OPECAJ') " & "WHERE IFTON.CodFile='" & strCodTipoInstrumentoOrden & "' ORDER BY DescripParametro"

    lblAnalitica.Caption = strCodTipoInstrumentoOrden & " - ????????"
    strCodFile = strCodTipoInstrumentoOrden

    '*** Clase de Instrumento ***
    strSQL = "SELECT CodDetalleFile CODIGO,DescripDetalleFile DESCRIP FROM InversionDetalleFile WHERE CodFile='" & strCodTipoInstrumentoOrden & "' AND IndVigente='X' ORDER BY DescripDetalleFile"
    CargarControlLista strSQL, cboClaseInstrumento, arrClaseInstrumento(), Sel_Defecto
    
    If cboClaseInstrumento.ListCount > 0 Then
        cboClaseInstrumento.ListIndex = 0
        cboClaseInstrumento.Enabled = True
    End If
    
End Sub

Private Sub cboMoneda_Click()
    
    lblDescripMoneda(0).Caption = "S/.": lblDescripMoneda(0).Tag = Codigo_Moneda_Local
   
    lblDescripMonedaResumen(0) = "S/.": lblDescripMonedaResumen(0).Tag = Codigo_Moneda_Local
    lblDescripMonedaResumen(1) = "S/.": lblDescripMonedaResumen(1).Tag = Codigo_Moneda_Local
    
    If cboMoneda.ListIndex < 0 Then Exit Sub
    
    strCodMoneda = Trim$(arrMoneda(cboMoneda.ListIndex))
        
    lblDescripMoneda(0).Caption = ObtenerSignoMoneda(strCodMoneda)
    lblDescripMoneda(0).Tag = strCodMoneda

    lblDescripMonedaResumen(0).Caption = ObtenerSignoMoneda(strCodMoneda)
    lblDescripMonedaResumen(0).Tag = strCodMoneda
    lblDescripMonedaResumen(1).Caption = ObtenerSignoMoneda(strCodMoneda)
    lblDescripMonedaResumen(1).Tag = strCodMoneda
    
    Call AsignarComisionOperacion
    
    If cboMonedaDocumento.ListIndex < 0 Then Exit Sub
    
    strCodMonedaDocumento = Trim$(arrMonedaDocumento(cboMonedaDocumento.ListIndex))
    
    If (strCodMonedaDocumento <> strCodMoneda) Then
        lblDescrip.Item(185).Visible = True
        txtTipoCambioDescuento.Visible = True
        strIndTipoCambio = Valor_Indicador
        txtTipoCambioDescuento.Text = 1
    Else
        lblDescrip.Item(185).Visible = False
        txtTipoCambioDescuento.Visible = False
        strIndTipoCambio = Valor_Caracter
    End If
    
    Call ActualizarValorNominalTC
    
    'Obtener Comisionistas de la moneda elegida con el obligado seleccionado
    
    strSQL = "{ call up_ACLstFondoComisionistaContraparte('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                Codigo_Tipo_Comisionista_Inversion & "','" & Codigo_Tipo_Persona_Emisor & "','" & strCodEmisor & "','" & _
                strCodMoneda & "','" & gstrFechaActual & "') }"
    CargarControlLista strSQL, cboComisionista, arrComisionista(), Valor_Caracter
    
    If cboComisionista.ListCount = 1 Then
        cboComisionista.ListIndex = 0
    End If
    
End Sub

Public Sub CargarComisiones(ByVal strCodComision As String, Index As Integer)
     
    Call AplicarCostos(Index)
     
End Sub

Private Sub AplicarCostos(Index As Integer)
        
    If strCodTipoCostoBolsa = Codigo_Tipo_Costo_Monto Then
        txtComisionBolsa(Index).Text = CStr(dblComisionBolsa)
    Else
        AsignaComision strCodTipoCostoBolsa, dblComisionBolsa, txtComisionBolsa(Index)
    End If
                         
    Call CalculoTotal(Index)
    
End Sub

Private Sub cboTipoOrden_Click()

    cboEmisor.Visible = True
    lblDescrip(6) = "Emisor"
    fraDatosAnexo.Visible = True
    fraDatosTitulo.Visible = True
    fraResumen.Visible = True

    'Obtener la clase de forma de pago (recibido o entregado según el tipo de operación
    'Se asume que en Inversiones es una sóla clase a la vez a diferencia de Cambios que se da recibidos y entregados
    
    strCodFile = strCodTipoInstrumentoOrden
    
    Me.MousePointer = vbDefault
    
End Sub

Private Sub cboTipoTasa_Click()

    strCodTipoTasa = Valor_Caracter

    If cboTipoTasa.ListIndex < 0 Then Exit Sub
    
    strCodTipoTasa = Trim$(arrTipoTasa(cboTipoTasa.ListIndex))
    lblIntAdelantado(0).Caption = CalculoInteresDescuento(CDbl(txtTasa.Text), strCodTipoTasa, strPeriodoTasa, strCodBaseAnual, txtValorNominalDcto, CStr(dtpFechaOrden.Value), CStr(dtpFechaPago.Value))
    txtCobroMinimoInteres.Text = CalculoInteresDescuento(CDbl(txtTasa.Text), strCodTipoTasa, strPeriodoTasa, strCodBaseAnual, txtValorNominalDcto, CStr(dtpFechaOrden.Value), CStr(DateAdd("d", intDiasInteresMinimo, dtpFechaOrden.Value)))

    Call CalculoTotal(0)
    
    If strCodTipoTasa = Codigo_Tipo_Tasa_Nominal Then
        lblPeriodoCapitalizacion.Enabled = True
        cboPeriodoCapitalizacion.Enabled = True
        strIndCapitalizable = Valor_Indicador
    Else
        lblPeriodoCapitalizacion.Enabled = False
        cboPeriodoCapitalizacion.Enabled = False
        strIndCapitalizable = Valor_Caracter
    End If
    
End Sub

Private Sub chkDiasAdicional_Click()

    If chkDiasAdicional.Visible = True Then

        If chkDiasAdicional.Value = Checked Then
        
            datFechaVctoAdicional = DateAdd("d", intDiasAdicionales, CVDate(dtpFechaVencimiento.Value))

            If Not EsDiaUtil(datFechaVctoAdicional) Then
                datFechaVctoAdicional = ProximoDiaUtil(datFechaVctoAdicional)
            End If

            lblFechaVencimientoAdic.Caption = datFechaVctoAdicional
            lblDiasAdic(0).Caption = "( " & CStr(DateDiff("d", dtpFechaVencimiento.Value, datFechaVctoAdicional)) & " días )"
            lblFechaVencimientoAdic.Visible = True
                        
        Else

            datFechaVctoAdicional = dtpFechaVencimiento.Value
            lblFechaVencimientoAdic.Caption = datFechaVctoAdicional
            lblFechaVencimientoAdic.Visible = False
            lblDiasAdic(0).Caption = 0
            
        End If
    
        txtDiasPlazo.Text = CStr(DateDiff("d", dtpFechaOrden.Value, dtpFechaVencimiento.Value))
        lblDiasPlazo.Caption = txtDiasPlazo.Text
        Call CalculoTotal(0)
    
    End If

End Sub

Private Sub chkInteresCorrido_Click(Index As Integer)

    If chkInteresCorrido(Index).Value Then
        txtInteresCorrido(Index).Text = CalculoInteresDescuento(CDbl(txtTasa.Text), strCodTipoTasa, "01", strCodBaseAnual, txtValorNominalDcto, CStr(dtpFechaEmision.Value), CStr(dtpFechaOrden.Value))
    Else
        txtInteresCorrido(Index).Text = 0#
    End If

    txtImptoInteresCorrido(Index).Text = (txtInteresCorrido(Index).Text * txtPorcenIgvInt(Index).Text / 100)
    Call CalculoTotal(Index)

End Sub

Private Sub cmdCalculo_Click()

    Call CalcularTirBruta
    
End Sub

Private Sub cmdCancelar_Click()
            blnCancelaPrepago = False
            Call Cancelar
End Sub

Private Sub cmdEnviar_Click()

    Dim strFechaDesde As String
    Dim intRegistro   As Integer, intContador         As Integer
    
    If adoConsulta.Recordset.RecordCount = 0 Then Exit Sub
    
    strFechaDesde = Convertyyyymmdd(dtpFechaOrdenDesde.Value)
    
    intContador = tdgConsulta.SelBookmarks.Count - 1
    
    If intContador < 0 Then
        MsgBox "No se ha seleccionado ningún registro", vbCritical, Me.Caption
        Exit Sub
    End If
        
    For intRegistro = 0 To intContador
        tdgConsulta.Row = tdgConsulta.SelBookmarks(intRegistro) - 1
               
        If strCodEstado = Estado_Orden_Ingresada Then
            adoComm.CommandText = "UPDATE InversionOrden SET EstadoOrden='" & Estado_Orden_Enviada & "'," & "UsuarioEdicion='" & gstrLogin & "',FechaEdicion='" & strFechaDesde & Space$(1) & Format$(Time, "hh:mm") & "' " & "WHERE NumOrden='" & Trim$(tdgConsulta.Columns(0)) & "' AND CodFondo='" & strCodFondo & "' AND " & "CodAdministradora='" & gstrCodAdministradora & "' AND EstadoOrden='" & Estado_Orden_Ingresada & "'"
        ElseIf strCodEstado = Estado_Orden_Enviada Then
            adoComm.CommandText = "UPDATE InversionOrden SET EstadoOrden='" & Estado_Orden_Ingresada & "'," & "UsuarioEdicion='" & gstrLogin & "',FechaEdicion='" & strFechaDesde & Space$(1) & Format$(Time, "hh:mm") & "' " & "WHERE NumOrden='" & Trim$(tdgConsulta.Columns(0)) & "' AND CodFondo='" & strCodFondo & "' AND " & "CodAdministradora='" & gstrCodAdministradora & "' AND EstadoOrden='" & Estado_Orden_Enviada & "'"
        End If

        adoConn.Execute adoComm.CommandText
    Next
    
    If strCodEstado = Estado_Orden_Ingresada Then
        MsgBox Mensaje_Envio_Exitoso, vbExclamation, gstrNombreEmpresa
    ElseIf strCodEstado = Estado_Orden_Enviada Then
        MsgBox Mensaje_Desenvio_Exitoso, vbExclamation, gstrNombreEmpresa
    ElseIf strCodEstado = Estado_Orden_Procesada Or strCodEstado = "" Then
        MsgBox "Las órdenes seleccionadas ya han sido confirmadas.", vbExclamation, gstrNombreEmpresa
    ElseIf strCodEstado = Estado_Orden_Anulada Then
        MsgBox "No puede enviarse a backoffice una orden anulada.", vbExclamation, gstrNombreEmpresa
    End If

    Call Buscar
    
End Sub

Private Sub cmdGuardar_Click()
    Call GrabarNew
End Sub

Private Sub dtpFechaEmision_Change()

    lblFechaEmision.Caption = CStr(dtpFechaEmision.Value)
    
    If chkInteresCorrido(0).Value = Checked Then
        txtInteresCorrido(0).Text = CalculoInteresDescuento(CDbl(txtTasa.Text), strCodTipoTasa, "01", strCodBaseAnual, txtValorNominalDcto, CStr(dtpFechaEmision.Value), CStr(dtpFechaOrden.Value))
        txtImptoInteresCorrido(0).Text = (txtInteresCorrido(0).Text * txtPorcenIgvInt(0).Text / 100)
    End If
    
End Sub

Private Sub dtpFechaLiquidacion_Change()

    If dtpFechaLiquidacion.Value < dtpFechaOrden.Value Then
        dtpFechaLiquidacion.Value = dtpFechaOrden.Value
    End If
        
    If Not EsDiaUtil(dtpFechaLiquidacion.Value) Then
        MsgBox "La Fecha no es un día útil...se cambiará por una fecha correcta !", vbInformation, Me.Caption
        dtpFechaLiquidacion.Value = ProximoDiaUtil(dtpFechaLiquidacion.Value)
    End If

    lblFechaLiquidacion.Caption = CStr(dtpFechaLiquidacion.Value)
    
    If strCodTipoInstrumentoOrden = "015" Then
        txtDiasPlazo.Text = CStr(DateDiff("d", dtpFechaOrden.Value, dtpFechaVencimiento.Value))
        lblDiasPlazo.Caption = txtDiasPlazo.Text
        Call CalculoTotal(0)
    End If
    
End Sub

Private Sub dtpFechaLiquidacionDesde_Click()

    If IsNull(dtpFechaLiquidacionDesde.Value) Then
        dtpFechaLiquidacionHasta.Value = Null
    Else
        dtpFechaLiquidacionDesde.Value = gdatFechaActual
        dtpFechaLiquidacionHasta.Value = dtpFechaLiquidacionDesde.Value
    End If
    
End Sub

Private Sub dtpFechaLiquidacionHasta_Click()

    If IsNull(dtpFechaLiquidacionHasta.Value) Then
        dtpFechaLiquidacionDesde.Value = Null
    Else
        dtpFechaLiquidacionDesde.Value = gdatFechaActual
        dtpFechaLiquidacionHasta.Value = dtpFechaLiquidacionDesde.Value
    End If
    
End Sub

Private Sub dtpFechaOrdenDesde_Click()

    If IsNull(dtpFechaOrdenDesde.Value) Then
        dtpFechaOrdenHasta.Value = Null
    Else
        dtpFechaOrdenDesde.Value = gdatFechaActual
        dtpFechaOrdenHasta.Value = dtpFechaOrdenDesde.Value
    End If
    
End Sub

Private Sub dtpFechaOrdenHasta_Click()

    If IsNull(dtpFechaOrdenHasta.Value) Then
        dtpFechaOrdenDesde.Value = Null
    Else
        dtpFechaOrdenDesde.Value = gdatFechaActual
        dtpFechaOrdenHasta.Value = dtpFechaOrdenDesde.Value
    End If
    
End Sub

Private Sub dtpFechaPago_Change()

    If dtpFechaPago.Value < dtpFechaVencimiento.Value Then
        dtpFechaPago.Value = dtpFechaVencimiento.Value
    End If
    
    If Not EsDiaUtil(dtpFechaPago.Value) Then
        MsgBox "La Fecha de Pago no es un día útil...se cambiará por una fecha correcta !", vbInformation, Me.Caption
        dtpFechaPago.Value = ProximoDiaUtil(dtpFechaPago.Value)
        dtpFechaVencimiento.Value = dtpFechaPago.Value
    End If
    
    Call CalculoTotal(0)
    
End Sub

Private Sub dtpFechaVencimiento_Change()
    '<EhHeader>
    On Error GoTo dtpFechaVencimiento_Change_Err
    '</EhHeader>

    If dtpFechaVencimiento.Value < dtpFechaOrden.Value Then
        dtpFechaVencimiento.Value = dtpFechaOrden.Value
    End If
    
    If dtpFechaVencimiento.Value < dtpFechaEmision.Value Then
        dtpFechaVencimiento.Value = dtpFechaEmision.Value
    End If
    
    If dtpFechaVencimiento.Value < dtpFechaLiquidacion.Value Then
        dtpFechaVencimiento.Value = dtpFechaLiquidacion.Value
    End If
    
    dtpFechaPago.Value = dtpFechaVencimiento.Value
    txtDiasPlazo.Text = CStr(DateDiff("d", dtpFechaOrden.Value, dtpFechaVencimiento.Value))
    lblDiasPlazo.Caption = txtDiasPlazo.Text
    
    lblFechaVencimiento.Caption = CStr(dtpFechaVencimiento.Value)
    
    '<EhFooter>
    Exit Sub

dtpFechaVencimiento_Change_Err:
    MsgBox err.Description & vbCrLf & "in Inversion.frmDesembolsoAcreencias.dtpFechaVencimiento_Change " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"
    Resume Next
    '</EhFooter>
End Sub

Private Sub dtpFechaVencimientoDcto_Change()

    If dtpFechaVencimientoDcto.Value < dtpFechaOrden.Value Then
        dtpFechaVencimientoDcto.Value = dtpFechaOrden.Value
    End If

    If dtpFechaVencimientoDcto.Value < dtpFechaEmision.Value Then
        dtpFechaVencimientoDcto.Value = dtpFechaEmision.Value
    End If

    dtpFechaVencimiento.Value = dtpFechaVencimientoDcto.Value
    dtpFechaPago.Value = dtpFechaVencimiento.Value

    chkDiasAdicional_Click

    Call dtpFechaVencimiento_Change
    
    lblFechaVencimiento.Caption = CStr(dtpFechaVencimiento.Value)

End Sub

Private Sub dtpFechaVencimientoDcto_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then

        If dtpFechaVencimientoDcto.Value < dtpFechaOrden.Value Then
            dtpFechaVencimientoDcto.Value = dtpFechaOrden.Value
        End If
    
        If dtpFechaVencimientoDcto.Value < dtpFechaEmision.Value Then
            dtpFechaVencimientoDcto.Value = dtpFechaEmision.Value
        End If
    
        If dtpFechaVencimientoDcto.Value < dtpFechaLiquidacion.Value Then
            dtpFechaVencimientoDcto.Value = dtpFechaLiquidacion.Value
        End If

    End If

End Sub

Private Sub Form_Activate()

    frmMainMdi.stbMdi.Panels(3).Text = Me.Caption
    Call CargarReportes
        
    
End Sub

Private Sub Form_Deactivate()

    Call OcultarReportes
    
End Sub

Private Sub Form_Load()
    
    blnCargadoDesdeCartera = False
    
    Call InicializarValores
    Call CargarListas
    Call CargarReportes
    Call Buscar

    Call ValidarPermisoUsoControl(Trim$(gstrLogin), Me, Trim$(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)

    CentrarForm Me
    
    Call ValidaExisteTipoCambio(Codigo_TipoCambio_SBS, gstrFechaActual)
            
End Sub

Public Sub Buscar()

    Dim strFechaOrdenDesde       As String, strFechaOrdenHasta        As String
    Dim strFechaLiquidacionDesde As String, strFechaLiquidacionHasta  As String
    Dim datFechaSiguiente        As Date
    Dim adoAuxiliar              As ADODB.Recordset
    
    Me.MousePointer = vbHourglass
    
    '*** Fecha Vigente, Moneda ***
    adoComm.CommandText = "{ call up_ACSelDatosParametro(23,'" & strCodFondo & "','" & gstrCodAdministradora & "','000') }"
    Set adoAuxiliar = adoComm.Execute
    
    If Not adoAuxiliar.EOF Then
        gdatFechaActual = CVDate(adoAuxiliar("FechaCuota"))
        strCodMoneda = Trim$(adoAuxiliar("CodMoneda"))
        frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
    End If

    adoAuxiliar.Close: Set adoAuxiliar = Nothing
    
    If Not IsNull(dtpFechaOrdenDesde.Value) And Not IsNull(dtpFechaOrdenHasta.Value) Then
        strFechaOrdenDesde = Convertyyyymmdd(dtpFechaOrdenDesde.Value)
        datFechaSiguiente = DateAdd("d", 1, dtpFechaOrdenHasta.Value)
        strFechaOrdenHasta = Convertyyyymmdd(datFechaSiguiente)
    End If
    
    If Not IsNull(dtpFechaLiquidacionDesde.Value) And Not IsNull(dtpFechaLiquidacionHasta.Value) Then
        strFechaLiquidacionDesde = Convertyyyymmdd(dtpFechaLiquidacionDesde.Value)
        datFechaSiguiente = DateAdd("d", 1, dtpFechaLiquidacionHasta.Value)
        strFechaLiquidacionHasta = Convertyyyymmdd(datFechaSiguiente)
    End If
            
    strSQL = "SELECT IOR.NumOrden,FechaOrden,FechaLiquidacion,CodTitulo,Nemotecnico,EstadoOrden,IOR.CodFile,CodAnalitica,TipoOrden,IOR.CodMoneda," & _
       "(RTRIM(DescripParametro) + SPACE(1) + DescripOrden) DescripOrden,CantOrden,ValorNominal,IOR.PorcenDsctoValorNominal,MontoTotalMFL1, " & _
       "CodSigno DescripMoneda, IOR.NumAnexo, NumDocumentoFisico,IOR.CodDetalleFile, IOR.CodSubDetalleFile, IOR.CodFondo, IOR.CodGirador, " & _
       "IP1.DescripPersona DesGirador, IOR.CodObligado, IP2.DescripPersona DesObligado, IOR.CodGestor, IP3.DescripPersona DesGestor, " & _
       "IOR.CodLimiteCli, IOR.CodEstructura, IOR.CodPersonaLim, IOR.TipoPersonaLim " & _
       "FROM InversionOrden IOR JOIN AuxiliarParametro AUX ON(AUX.CodParametro=IOR.TipoOrden AND AUX.CodTipoParametro = 'OPECAJ') " & _
       "JOIN Moneda MON ON(MON.CodMoneda=IOR.CodMoneda) " & _
       "LEFT JOIN InstitucionPersona IP1 ON (IP1.CodPersona = IOR.CodGirador AND IP1.TipoPersona = '" & Codigo_Tipo_Persona_Emisor & "') " & _
       "LEFT JOIN InstitucionPersona IP2 ON (IP2.CodPersona = IOR.CodObligado AND IP2.TipoPersona = '" & Codigo_Tipo_Persona_Obligado & "') " & _
       "LEFT JOIN InstitucionPersona IP3 ON (IP3.CodPersona = IOR.CodGestor AND IP3.TipoPersona = '" & Codigo_Tipo_Persona_Emisor & "') " & _
       "WHERE IOR.TipoOrden = '" & Codigo_Orden_Compra & "' AND  IOR.CodAdministradora='" & gstrCodAdministradora & "' AND IOR.CodFondo='" & strCodFondo & "' "
        
    If strCodTipoInstrumento <> Valor_Caracter Then
        strSQL = strSQL & "AND IOR.CodFile='" & strCodTipoInstrumento & "' "
    Else
        strSQL = strSQL & "AND IOR.CodFile IN ('" & CodFile_Descuento_Comprobantes_Pago & "','" & CodFile_Descuento_Documentos_Cambiario & "')"
    End If

    If Not IsNull(dtpFechaOrdenDesde.Value) And Not IsNull(dtpFechaOrdenHasta.Value) Then
        strSQL = strSQL & "AND (FechaOrden >='" & strFechaOrdenDesde & "' AND FechaOrden <'" & strFechaOrdenHasta & "') "
    End If
    
    If Not IsNull(dtpFechaLiquidacionDesde.Value) And Not IsNull(dtpFechaLiquidacionHasta.Value) Then
        strSQL = strSQL & "AND (FechaLiquidacion >='" & strFechaLiquidacionDesde & "' AND FechaLiquidacion <'" & strFechaLiquidacionHasta & "') "
    End If
    
    If strCodEstado <> Valor_Caracter Then
        strSQL = strSQL & "AND EstadoOrden='" & strCodEstado & "' "
    End If
    
    If (strLineaClienteLista <> Valor_Caracter) And (cboTipoInstrumento.ListIndex > 0) Then
        strSQL = strSQL & "AND IOR.CodLimiteCli='" & strLineaClienteLista & "' AND IOR.CodEstructura ='" & Codigo_LimiteRE_Cliente & "' "
    End If
    
    strSQL = strSQL & "ORDER BY IOR.NumOrden"
    
    With adoConsulta
        .ConnectionString = gstrConnectConsulta
        .RecordSource = strSQL
        .Refresh
    End With

    tdgConsulta.Refresh

    If adoConsulta.Recordset.RecordCount > 0 Then strEstado = Reg_Consulta

    Me.MousePointer = vbDefault
    
End Sub

Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Ordenes de Inversión"
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Text = "Papeleta de Inversión"
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo3").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo3").Text = "Anexo"
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo4").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo4").Text = "Anexo Cliente"
    
End Sub

Private Sub CargarListas()
    Dim adoRecord   As ADODB.Recordset
    Dim intRegistro As Integer
    
    '*** Fondos ***
    strSQL = "SELECT CodFondo CODIGO,DescripFondo DESCRIP FROM Fondo WHERE CodAdministradora='" & gstrCodAdministradora & "' AND Estado='01' and CodFondo = '" & gstrCodFondoContable & "' ORDER BY DescripFondo"
    CargarControlLista strSQL, cboFondo, arrFondo(), Valor_Caracter
    CargarControlLista strSQL, cboFondoOrden, arrFondoOrden(), Valor_Caracter
        
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
            
    '*** Estados de la Orden ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='ESTORD' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboEstado, arrEstado(), Sel_Todos
    
    intRegistro = ObtenerItemLista(arrEstado(), Estado_Orden_Ingresada)

    If intRegistro >= 0 Then cboEstado.ListIndex = intRegistro
        
    '*** Emisor ***
    strSQL = "SELECT (CodPersona) CODIGO,DescripPersona DESCRIP FROM InstitucionPersona WHERE TipoPersona='" & Codigo_Tipo_Persona_Emisor & "' AND IndVigente='X' ORDER BY DescripPersona"
    CargarControlLista strSQL, cboEmisor, arrEmisor(), Sel_Defecto

    '*** Obligado ***
    strSQL = "SELECT (CodPersona) CODIGO,DescripPersona DESCRIP FROM InstitucionPersona WHERE TipoPersona='" & Codigo_Tipo_Persona_Obligado & "' AND IndVigente='X' ORDER BY DescripPersona"
    CargarControlLista strSQL, cboObligado, arrObligado(), Sel_Defecto

    '*** Gestor ***
    strSQL = "SELECT (CodPersona) CODIGO,DescripPersona DESCRIP FROM InstitucionPersona WHERE TipoPersona='" & Codigo_Tipo_Persona_Emisor & "' AND IndVigente='X' AND IndBanco = 'X' ORDER BY DescripPersona"
    CargarControlLista strSQL, cboGestor, arrGestor(), Sel_Defecto
                
    '*** Moneda Documento ***
    strSQL = "{ call up_ACSelDatos(2) }"
    CargarControlLista strSQL, cboMonedaDocumento, arrMonedaDocumento(), Sel_Defecto
    
    '*** Moneda ***
    strSQL = "{ call up_ACSelDatos(2) }"
    CargarControlLista strSQL, cboMoneda, arrMoneda(), Sel_Defecto
    
    '*** Base de Cálculo ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='BASANU' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboBaseAnual, arrBaseAnual(), Valor_Caracter
    
    '*** Tipo Tasa ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='NATTAS' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboTipoTasa, arrTipoTasa(), ""
    
    '*** Mecanismos de Negociación ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='MECNEG' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboNegociacion, arrNegociacion(), Valor_Caracter
    
    '*** Momento de cobro de los intereses (por defecto es al inicio) ***
    strSQL = "SELECT (CodTipoParametro + CodParametro) CODIGO, DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='MODPAG' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboCobroInteres, arrPagoInteres(), ""

    If cboCobroInteres.ListCount > 0 Then cboCobroInteres.ListIndex = 0
        
    '*** Pagador de la deuda. Inicio Operación ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='RESPAG' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboResponsablePago, arrResponsablePago(), Sel_Defecto

    If cboResponsablePago.ListCount > 0 Then cboResponsablePago.ListIndex = 0
    
    '*** Periodo de Tasa ***
    strSQL = "SELECT CodParametro CODIGO, DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro = 'TIPFRE'"
    CargarControlLista strSQL, cboPeriodoTasa, arrPeriodoTasa(), Sel_Defecto
    cboPeriodoTasa.ListIndex = 1
    
    '*** Periodo de Capitalizacion ***
    strSQL = "SELECT CodParametro CODIGO, DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro = 'TIPFRE'"
    CargarControlLista strSQL, cboPeriodoCapitalizacion, arrPeriodoCapitalizacion(), Sel_Defecto
    cboPeriodoCapitalizacion.ListIndex = 7
    
    adoComm.CommandText = "SELECT isnull(ValorTipoCambioCompra,1) as ValorTipoCambioCompra from " & " TipoCambioFondo " & " where FechaTipoCambio='" & gstrFechaActual & "' and CodTipoCambio = '02' "
                            
    Set adoRecord = adoComm.Execute

    If Not (adoRecord.EOF) Then
        txtTipoCambioDescuento.Text = adoRecord("ValorTipoCambioCompra")
    Else
        txtTipoCambioDescuento.Text = 1
    End If

End Sub

Public Sub Accion(nAccion As ButtonAction)
    
    Select Case nAccion
                
        Case vNew
            Call Adicionar

        Case vDelete
            Call Eliminar

        Case vSearch
            Call Buscar

        Case vReport

            'Call Imprimir
        Case vSave
            Call GrabarNew

        Case vCancel
            blnCancelaPrepago = False
            Call Cancelar

        Case vExit
            Call Salir
        
    End Select
    
End Sub

Private Sub InicializarValores()
    
    Dim adoRegistro As ADODB.Recordset
    
    '*** Valores Iniciales ***
    strEstado = Reg_Defecto
    tabRFCortoPlazo.Tab = 0

    If strCodTipoTasa = Codigo_Tipo_Tasa_Nominal Then
        lblPeriodoCapitalizacion.Enabled = True
        cboPeriodoCapitalizacion.Enabled = True
        strIndCapitalizable = Valor_Indicador
    Else
        lblPeriodoCapitalizacion.Enabled = False
        cboPeriodoCapitalizacion.Enabled = False
        strIndCapitalizable = Valor_Caracter
    End If
    
    txtNumAnexo = Valor_Caracter
    
    If (cbxGenerarLetra.Value = vbChecked) Then
        strIndGeneraLetra = Valor_Indicador
    Else
        strIndGeneraLetra = Valor_Caracter
    End If
    
    dtpFechaOrdenDesde.Value = gdatFechaActual
    dtpFechaOrdenHasta.Value = dtpFechaOrdenDesde.Value
    dtpFechaLiquidacionDesde.Value = Null
    dtpFechaLiquidacionHasta.Value = dtpFechaLiquidacionDesde.Value
    
    txtPorcenIgvInt(0).Text = CStr(gdblTasaIgv * 100)
    txtPorcenIgvInt2(1).Text = CStr(gdblTasaIgv * 100)
    
    txtTasa.Text = "0"
    txtPorcenIgv(0).Text = CStr(gdblTasaIgv * 100)
    
    strCodTitulo = Valor_Caracter
    strResponsablePagoCancel = Valor_Caracter
    strViaCobranza = Valor_Caracter
    
    Set adoRegistro = New ADODB.Recordset

    With adoComm
        .CommandText = "SELECT CodFile FROM InversionFile " & "WHERE TipoValor='" & Codigo_Valor_RentaFija & "' AND TipoPlazo='" & Codigo_Valor_CortoPlazo & "' AND IndInstrumento='X' AND IndVigente='X' " & "ORDER BY DescripFile"
        Set adoRegistro = .Execute
                
        strCodigosFile = Valor_Caracter

        Do While Not adoRegistro.EOF

            If strCodigosFile <> Valor_Caracter Then strCodigosFile = strCodigosFile & ",'"
            
            strCodigosFile = strCodigosFile & Trim$(adoRegistro("CodFile")) & "'"
        
            adoRegistro.MoveNext
        Loop

        adoRegistro.Close: Set adoRegistro = Nothing
                
        strCodigosFile = "('" & strCodigosFile & ",'009')"
    End With
        
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
  
    'Leer si las comisiones van a ser definidas en la operación o ya viene establecida
    Set adoRegistro = New ADODB.Recordset
    adoComm.CommandText = "SELECT ValorParametro AS PersonalizaComi FROM ParametroGeneral WHERE CodParametro = '36'"
    Set adoRegistro = adoComm.Execute

    If Not (adoRegistro.EOF) Then
        strPersonalizaComision = Trim$(adoRegistro("PersonalizaComi"))
    End If
    
    'Leer el porcentaje de descuento
    Set adoRegistro = New ADODB.Recordset
    adoComm.CommandText = "SELECT ValorParametro AS PorcentajeDscto FROM ParametroGeneral WHERE CodParametro = '37'"
    Set adoRegistro = adoComm.Execute

    If Not (adoRegistro.EOF) Then
        dblPorcenDescuento = CDbl(adoRegistro("PorcentajeDscto"))
    End If
        
    adoRegistro.Close: Set adoRegistro = Nothing
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmDesembolsoAcreencias = Nothing
    Call OcultarReportes
    frmMainMdi.stbMdi.Panels(3).Text = "Acción"
    
End Sub

Private Sub lblCantidadResumen_Change()

    Call FormatoMillarEtiqueta(lblCantidadResumen, Decimales_Monto)
    
End Sub

Private Sub lblComisionesResumen_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblComisionesResumen(Index), Decimales_Monto)
    
End Sub

Private Sub lblComisionIgv_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblComisionIgv(Index), Decimales_Monto)
    
End Sub

Private Sub lblInteresesResumen_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblInteresesResumen(Index), Decimales_Monto)
    
End Sub

Private Sub lblMontoTotal_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblMontoTotal(Index), Decimales_Monto)
    
    lblTotalResumen(Index).Caption = lblMontoTotal(Index).Caption
    
End Sub

Private Sub lblMontoVencimiento_Change()

    Call FormatoMillarEtiqueta(lblMontoVencimiento, Decimales_Monto)
    
End Sub

Private Sub lblPorcenBolsa_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblPorcenBolsa(Index), Decimales_Tasa)
    
End Sub

Private Sub lblPrecioResumen_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblPrecioResumen(Index), Decimales_Precio)
    
End Sub

'Private Sub lblSubTotal_Change(Index As Integer)
'
'    Call FormatoMillarEtiqueta(lblSubTotal(Index), Decimales_Monto)
'
'    If Not IsNumeric(txtPorcenAgente(Index).Text) Or Not IsNumeric(lblPorcenBolsa(Index).Caption) Then Exit Sub
'
'    'Calcula comisiones
'    'Mientras se cargue la comisión automáticame no se hace este cálculo
'    txtComisionBolsa(Index).Text = CDbl((lblSubTotal(Index).Caption)) * CDbl(lblPorcenBolsa(Index).Caption) / 100
'
'    If Not IsNumeric(txtTasa.Text) Or Not IsNumeric(txtCantidad.Text) Then Exit Sub
'
'    'Calcula interes corrido
'
'    Call CalculoTotal(Index)
'
'    lblSubTotalResumen(Index).Caption = CStr(CCur(lblSubTotal(Index).Caption))
'
'End Sub

Private Sub lblSubTotalResumen_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblSubTotalResumen(Index), Decimales_Monto)
    
End Sub

Private Sub lblTirBruta_Change()

    Call FormatoMillarEtiqueta(lblTirBruta, Decimales_Tasa)
    
End Sub

Private Sub lblTirBrutaResumen_Change()

    Call FormatoMillarEtiqueta(lblTirBrutaResumen, Decimales_Tasa)
    
End Sub

Private Sub lblTirNeta_Change()

    Call FormatoMillarEtiqueta(lblTirNeta, Decimales_Tasa)
    
End Sub

Private Sub lblTirNetaResumen_Change()

    Call FormatoMillarEtiqueta(lblTirNetaResumen, Decimales_Tasa)
    
End Sub

Private Sub lblTotalResumen_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblTotalResumen(Index), Decimales_Monto)
    
End Sub

Private Sub lblVencimientoResumen_Change()

    Call FormatoMillarEtiqueta(lblVencimientoResumen, Decimales_Monto)
    
End Sub

Private Sub tabRFCortoPlazo_Click(PreviousTab As Integer)
            
    Select Case tabRFCortoPlazo.Tab

        Case 1, 2, 3, 4

            If PreviousTab = 0 And blnCargadoDesdeCartera = False And strEstado = Reg_Consulta Then tabRFCortoPlazo.Tab = 0
            If strEstado = Reg_Defecto Then tabRFCortoPlazo.Tab = 0
            
            If tabRFCortoPlazo.Tab = 1 Then
                cmdGuardar.Enabled = False
            End If
                       
            If tabRFCortoPlazo.Tab = 2 Then
               
                If ValidaRequisitosTab(2, PreviousTab) = True Then
                    txtValorNominal.Text = txtValorNominalDescuento.Text
                    fraDatosNegociacion.Caption = "Negociación" & Space$(1) & "-" & Space$(1) & txtNemonico.Text
                    lblIntAdelantado(0).Caption = CalculoInteresDescuento(CDbl(txtTasa.Text), strCodTipoTasa, strPeriodoTasa, strCodBaseAnual, txtValorNominalDcto, CStr(dtpFechaOrden.Value), CStr(dtpFechaPago.Value))
                    txtCobroMinimoInteres.Text = CalculoInteresDescuento(CDbl(txtTasa.Text), strCodTipoTasa, strPeriodoTasa, strCodBaseAnual, txtValorNominalDcto, CStr(dtpFechaOrden.Value), CStr(DateAdd("d", intDiasInteresMinimo, dtpFechaOrden.Value)))

                    ObtenerParametrosGeneralesEmisor
                    Call CalcularComision
                    Call CalculoTotal(0)
                    cmdGuardar.Enabled = True
                Else
                    tabRFCortoPlazo.Tab = 1
                    cmdGuardar.Enabled = False
                End If
               
            End If
            
    End Select
    
End Sub

Private Sub ActualizarTotalAnexo()
    Dim adoRegistro            As ADODB.Recordset
    
    If Not blnEmisorReady Then Exit Sub

    If txtTotalMNAnexo.Text = "" Then txtTotalMNAnexo.Text = 0
    If txtTotalMEAnexo.Text = "" Then txtTotalMEAnexo.Text = 0
    If txtTotalDctosAnexo.Text = "" Then txtTotalDctosAnexo.Text = 1

    'Ahora el sistema aquí calculará cuanto suma lo registrado + lo que se está registrando actualmente para calcular la comisión.
    txtTotalMNAnexo.Text = 0
    txtTotalMEAnexo.Text = 0
    
    adoComm.CommandText = "select isnull(SUM(ValorNominal),0) as ValorNominalAnexo, isnull(SUM(ValorNominalDscto),0) as ValorNominalDsctoAnexo  from InversionOrden where " & _
                            "CodFondo = '" & strCodFondo & "' and CodEmisor = '" & strCodEmisor & "' and NumAnexo = '" & _
                            strNumAnexo & "' and EstadoOrden <> '01' " & " and TipoOrden = '01' and CodFile in ('" & _
                            CodFile_Descuento_Comprobantes_Pago & "','" & CodFile_Descuento_Documentos_Cambiario & "')"
    Set adoRegistro = adoComm.Execute
        
    dblTotalMNAnexo = adoRegistro("ValorNominalAnexo")
    dblTotalAnexoDscto = adoRegistro("ValorNominalDsctoAnexo")
        
'    If strCodMoneda = Codigo_Moneda_Local Then
        txtTotalMNAnexo.Text = adoRegistro("ValorNominalAnexo") + txtValorNominalDescuento.Value
        txtTotalMEAnexo.Text = adoRegistro("ValorNominalDsctoAnexo") + CDbl(txtValorNominalDcto.Text)
'    Else
'        txtTotalMNAnexo.Text = adoRegistro("ValorNominalAnexo")
'    End If
    
'    adoComm.CommandText = "select isnull(SUM(ValorNominal),0) as ValorNominalAnexoME  from InversionOrden where " & "CodFondo = '" & strCodFondo & "' and CodGirador = '" & strCodEmisor & "' and NumAnexo = '" & strNumAnexo & "' and EstadoOrden <> '01' " & "and CodMoneda = '" & Codigo_Moneda_Dolar_Americano & "'"
'    Set adoRegistro = adoComm.Execute
'
'    If strCodMoneda = Codigo_Moneda_Dolar_Americano Then
'        txtTotalMEAnexo.Text = adoRegistro("ValorNominalAnexoME") + txtValorNominalDescuento.Value
'    Else
'        txtTotalMEAnexo.Text = adoRegistro("ValorNominalAnexoME")
'    End If
End Sub

Private Sub CalcularComision()
    'Calcula la comisión que debe tener cada operación en la moneda de cobro de las comisiones (parámetros globales)

    Dim adoRegistro            As ADODB.Recordset
    Dim TCComision             As Double
    Dim dblTotalMEAnexo        As Double
    Dim dblTotalMEAnexoDesc    As Double
    Dim dblsubTotal1           As Double
    Dim dblsubTotal2           As Double
    Dim intCantidadOperaciones As Integer
    
    ActualizarTotalAnexo

    If (CDec(txtTotalMNAnexo.Text) = 0 And CDec(txtTotalMEAnexo.Text) = 0) Or CInt(txtTotalDctosAnexo.Text) = 0 Then
        txtPorcenAgente(0).Text = 0#
        txtComisionAgente(0).Text = 0#
        Exit Sub
    End If
  
    'La moneda de comision siempre será la moneda de la orden.
    strCodMonedaComision = strCodMoneda
     
    txtPorcenAgente(0).Text = dblPorcentajeComision
     'Pasar los importes totales a la moneda de cobro de la comisión
    If (txtTotalMNAnexo.Text <> "") And (CDec(txtTotalMNAnexo.Text) <> 0) Then
        strCodMonedaParEvaluacion = Trim$(Codigo_Moneda_Local) & Trim$(strCodMoneda)

        If Codigo_Moneda_Local <> strCodMoneda Then
            strCodMonedaParPorDefecto = ObtenerMonedaParPorDefecto(gstrCodClaseTipoCambioOperacionFondo, strCodMonedaParEvaluacion)
        Else
            strCodMonedaParPorDefecto = strCodMonedaParEvaluacion
        End If
                
        If strCodMonedaParPorDefecto = "0000" Then strCodMonedaParPorDefecto = strCodMonedaParEvaluacion
      
        dblsubTotal1 = CDec(txtTotalMEAnexo.Text)
            
    End If
    
'    'Pasar los importes totales a la moneda de cobro de la comisión
'    If (txtTotalMNAnexo.Text <> "") And (CDec(txtTotalMNAnexo.Text) <> 0) Then 'Se indicaron totales en Soles
'        strCodMonedaParEvaluacion = Trim$(Codigo_Moneda_Local) & Trim$(strCodMoneda)
'
'        If Codigo_Moneda_Local <> strCodMonedaComision Then
'            strCodMonedaParPorDefecto = ObtenerMonedaParPorDefecto(gstrCodClaseTipoCambioOperacionFondo, strCodMonedaParEvaluacion)
'        Else
'            strCodMonedaParPorDefecto = strCodMonedaParEvaluacion
'        End If
'
'        If strCodMonedaParPorDefecto = "0000" Then strCodMonedaParPorDefecto = strCodMonedaParEvaluacion
'
'        If strCodMonedaComision <> Codigo_Moneda_Local Then
'
'            TCComision = 0#
'            TCComision = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, dtpFechaOrden.Value, Mid$(strCodMonedaParPorDefecto, 1, 2), Mid$(strCodMonedaParPorDefecto, 3, 2)))
'
'            If TCComision <> 0 Then
'                If strCodMonedaParEvaluacion = strCodMonedaParPorDefecto Then
'                    dblsubTotal1 = txtTotalMNAnexo.Value / TCComision
'                    dblMontoMinComisiones = dblMontoMinComisiones / TCComision
'                Else
'                    dblsubTotal1 = txtTotalMNAnexo.Value * TCComision
'                    dblMontoMinComisiones = dblMontoMinComisiones * TCComision
'                End If
'
'            Else
'                dblsubTotal1 = 0#
'            End If
'
'        Else
'            dblsubTotal1 = CDec(txtTotalMNAnexo.Text)
'
'        End If
'    End If
'
'    If (txtTotalMEAnexo.Text <> "") And CDec(txtTotalMEAnexo.Text) <> 0 Then
'        strCodMonedaParEvaluacion = Trim$(Codigo_Moneda_Dolar_Americano) & Trim$(strCodMonedaComision)
'
'        If Codigo_Moneda_Dolar_Americano <> strCodMonedaComision Then
'            strCodMonedaParPorDefecto = ObtenerMonedaParPorDefecto(gstrCodClaseTipoCambioOperacionFondo, strCodMonedaParEvaluacion)
'        Else
'            strCodMonedaParPorDefecto = strCodMonedaParEvaluacion
'        End If
'
'        If strCodMonedaParPorDefecto = "0000" Then strCodMonedaParPorDefecto = strCodMonedaParEvaluacion
'
'        If strCodMonedaComision <> Codigo_Moneda_Dolar_Americano Then
'
'            TCComision = 0#
'            TCComision = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, dtpFechaOrden.Value, Mid$(strCodMonedaParPorDefecto, 1, 2), Mid$(strCodMonedaParPorDefecto, 3, 2)))
'
'            If TCComision <> 0 Then
'                If strCodMonedaParEvaluacion = strCodMonedaParPorDefecto Then
'                    dblsubTotal2 = txtTotalMEAnexo.Value / TCComision
'                Else
'                    dblsubTotal2 = txtTotalMEAnexo.Value * TCComision
'                End If
'
'            Else
'                dblsubTotal2 = 0#
'            End If
'
'        Else
'            dblsubTotal2 = CDec(txtTotalMEAnexo.Text)
'
'        End If
'    End If
    
    If txtPorcenDctoValorNominal.Text = "" Then
        txtPorcenDctoValorNominal.Text = 0
    End If
    
    dblTotalMEAnexo = (dblsubTotal1 + dblsubTotal2)
    dblTotalMEAnexoDesc = dblTotalMEAnexo '* (CDbl(txtPorcenDctoValorNominal.Text) / 100)
        
    'Se cuentan las operaciones y ordenes del anexo, que no estén anuladas:
    Set adoRegistro = New ADODB.Recordset
    adoComm.CommandText = "select COUNT(NumOrden) as CantidadOperaciones from InversionOrden where " & "CodFondo = '" & strCodFondo & "' and CodGirador = '" & strCodEmisor & "' and NumAnexo = '" & strNumAnexo & "' and EstadoOrden <> '01' and TipoOrden = '01' and CodFile in ('" & CodFile_Descuento_Comprobantes_Pago & "','" & CodFile_Descuento_Documentos_Cambiario & "')"
                            
    Set adoRegistro = adoComm.Execute

    If Not (adoRegistro.EOF) Then
        intCantidadOperaciones = adoRegistro("CantidadOperaciones")
    Else
        intCantidadOperaciones = 0
    End If
 
    If Not (adoRegistro.EOF) Then
        If intCantidadOperaciones = (txtTotalDctosAnexo.Text - 1) Then
            dblComisionOperacion = dblTotalMEAnexoDesc * (dblPorcentajeComision / 100)

            'Comision total
            If dblComisionOperacion < dblMontoMinComisiones Then
                dblComisionOperacion = dblMontoMinComisiones
            End If

        Else
            dblComisionOperacion = 0
        End If
    End If
    
    Call AsignarComisionOperacion
    
End Sub

Private Sub CalcularComisionOld()
    'Calcula la comisión que debe tener cada operación en la moneda de cobro de las comisiones (parámetros globales)

    Dim adoRegistro            As ADODB.Recordset
    Dim TCComision             As Double
    Dim dblTotalMEAnexo        As Double
    Dim dblTotalMEAnexoDesc    As Double
    Dim dblsubTotal1           As Double
    Dim dblsubTotal2           As Double
    Dim intCantidadOperaciones As Integer
    
    ActualizarTotalAnexo

    If (CDec(txtTotalMNAnexo.Text) = 0 And CDec(txtTotalMEAnexo.Text) = 0) Or CInt(txtTotalDctosAnexo.Text) = 0 Then
        txtPorcenAgente(0).Text = 0#
        txtComisionAgente(0).Text = 0#
        Exit Sub
    End If
  
    'La moneda de comision siempre será la moneda de la orden.
    strCodMonedaComision = strCodMoneda
     
    txtPorcenAgente(0).Text = dblPorcentajeComision
    
    'Pasar los importes totales a la moneda de cobro de la comisión
    If (txtTotalMNAnexo.Text <> "") And (CDec(txtTotalMNAnexo.Text) <> 0) Then 'Se indicaron totales en Soles
        strCodMonedaParEvaluacion = Trim$(Codigo_Moneda_Local) & Trim$(strCodMoneda)

        If Codigo_Moneda_Local <> strCodMonedaComision Then
            strCodMonedaParPorDefecto = ObtenerMonedaParPorDefecto(gstrCodClaseTipoCambioOperacionFondo, strCodMonedaParEvaluacion)
        Else
            strCodMonedaParPorDefecto = strCodMonedaParEvaluacion
        End If
                
        If strCodMonedaParPorDefecto = "0000" Then strCodMonedaParPorDefecto = strCodMonedaParEvaluacion
      
        If strCodMonedaComision <> Codigo_Moneda_Local Then
            
            TCComision = 0#
            TCComision = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, dtpFechaOrden.Value, Mid$(strCodMonedaParPorDefecto, 1, 2), Mid$(strCodMonedaParPorDefecto, 3, 2)))

            If TCComision <> 0 Then
                If strCodMonedaParEvaluacion = strCodMonedaParPorDefecto Then
                    dblsubTotal1 = txtTotalMNAnexo.Value / TCComision
                    dblMontoMinComisiones = dblMontoMinComisiones / TCComision
                Else
                    dblsubTotal1 = txtTotalMNAnexo.Value * TCComision
                    dblMontoMinComisiones = dblMontoMinComisiones * TCComision
                End If

            Else
                dblsubTotal1 = 0#
            End If

        Else
            dblsubTotal1 = CDec(txtTotalMNAnexo.Text)
            
        End If
    End If
  
    If (txtTotalMEAnexo.Text <> "") And CDec(txtTotalMEAnexo.Text) <> 0 Then
        strCodMonedaParEvaluacion = Trim$(Codigo_Moneda_Dolar_Americano) & Trim$(strCodMonedaComision)
    
        If Codigo_Moneda_Dolar_Americano <> strCodMonedaComision Then
            strCodMonedaParPorDefecto = ObtenerMonedaParPorDefecto(gstrCodClaseTipoCambioOperacionFondo, strCodMonedaParEvaluacion)
        Else
            strCodMonedaParPorDefecto = strCodMonedaParEvaluacion
        End If
                    
        If strCodMonedaParPorDefecto = "0000" Then strCodMonedaParPorDefecto = strCodMonedaParEvaluacion
          
        If strCodMonedaComision <> Codigo_Moneda_Dolar_Americano Then
                
            TCComision = 0#
            TCComision = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, dtpFechaOrden.Value, Mid$(strCodMonedaParPorDefecto, 1, 2), Mid$(strCodMonedaParPorDefecto, 3, 2)))

            If TCComision <> 0 Then
                If strCodMonedaParEvaluacion = strCodMonedaParPorDefecto Then
                    dblsubTotal2 = txtTotalMEAnexo.Value / TCComision
                Else
                    dblsubTotal2 = txtTotalMEAnexo.Value * TCComision
                End If

            Else
                dblsubTotal2 = 0#
            End If

        Else
            dblsubTotal2 = CDec(txtTotalMEAnexo.Text)
                
        End If
    End If
    
    If txtPorcenDctoValorNominal.Text = "" Then
        txtPorcenDctoValorNominal.Text = 0
    End If
    
    dblTotalMEAnexo = (dblsubTotal1 + dblsubTotal2)
    dblTotalMEAnexoDesc = dblTotalMEAnexo * (CDbl(txtPorcenDctoValorNominal.Text) / 100)
        
    'Se cuentan las operaciones y ordenes del anexo, que no estén anuladas:
    Set adoRegistro = New ADODB.Recordset
    adoComm.CommandText = "select COUNT(NumOrden) as CantidadOperaciones from InversionOrden where " & "CodFondo = '" & strCodFondo & "' and CodGirador = '" & strCodEmisor & "' and NumAnexo = '" & strNumAnexo & "' and EstadoOrden <> '01' and TipoOrden = '01' and CodFile in ('" & CodFile_Descuento_Comprobantes_Pago & "','" & CodFile_Descuento_Documentos_Cambiario & "')"
                            
    Set adoRegistro = adoComm.Execute

    If Not (adoRegistro.EOF) Then
        intCantidadOperaciones = adoRegistro("CantidadOperaciones")
    Else
        intCantidadOperaciones = 0
    End If
 
    If Not (adoRegistro.EOF) Then
        If intCantidadOperaciones = (txtTotalDctosAnexo.Text - 1) Then
            dblComisionOperacion = dblTotalMEAnexoDesc * (dblPorcentajeComision / 100)

            'Comision total
            If dblComisionOperacion < dblMontoMinComisiones Then
                dblComisionOperacion = dblMontoMinComisiones
            End If

        Else
            dblComisionOperacion = 0
        End If
    End If
    
    Call AsignarComisionOperacion
    
End Sub

Private Sub AsignarComisionOperacion()

    Dim dblTipoCambio2 As Double
    
    If Trim$(strCodMonedaDocumento) <> Trim$(strCodMonedaComision) Then
        
        'convertir el monto de comisiones a la moneda de la operación
        strCodMonedaParEvaluacion = Trim$(strCodMonedaComision) & Trim$(strCodMoneda)
        strCodMonedaParPorDefecto = ObtenerMonedaParPorDefecto(gstrCodClaseTipoCambioOperacionFondo, strCodMonedaParEvaluacion)

        If strCodMonedaParPorDefecto = "0000" Then strCodMonedaParPorDefecto = strCodMonedaParEvaluacion
  
        dblTipoCambio2 = 0#
        dblTipoCambio2 = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, dtpFechaOrden.Value, Mid$(strCodMonedaParPorDefecto, 1, 2), Mid$(strCodMonedaParPorDefecto, 3, 2)))
        
        If dblTipoCambio2 <> 0 Then
            If strCodMonedaParEvaluacion = strCodMonedaParPorDefecto Then
                txtComisionAgente(0).Text = dblComisionOperacion / dblTipoCambio2
            Else
                txtComisionAgente(0).Text = dblComisionOperacion * dblTipoCambio2
            End If

        Else
            txtComisionAgente(0).Text = 0#
        End If
    
    Else
        txtComisionAgente(0).Text = dblComisionOperacion
    End If
    
    If CDec(txtComisionAgente(0).Text) <> 0 Then
        Call CalculoTotal(0)
    End If

End Sub

Private Function ValidaRequisitosTab(intIndTab As Integer, intTabOrigen) As Boolean

    ValidaRequisitosTab = False

    Select Case intIndTab

        Case 2

            If cboEmisor.ListIndex <= 0 Then
                MsgBox "Debe seleccionar el Emisor.", vbCritical, Me.Caption

                If cboEmisor.Enabled And tabRFCortoPlazo.TabEnabled(1) = True Then
                    tabRFCortoPlazo.Tab = 1
                    cboEmisor.SetFocus
                End If

                Exit Function
            End If
    
            If cboObligado.ListIndex <= 0 Then
                MsgBox "Debe seleccionar el Obligado.", vbCritical, Me.Caption

                If cboObligado.Enabled And tabRFCortoPlazo.TabEnabled(1) = True Then
                    tabRFCortoPlazo.Tab = 1
                    cboObligado.SetFocus
                End If

                Exit Function
            End If
    
            If cboGestor.ListIndex <= 0 Then
                MsgBox "Debe seleccionar el Gestor.", vbCritical, Me.Caption

                If cboGestor.Enabled And tabRFCortoPlazo.TabEnabled(1) = True Then
                    tabRFCortoPlazo.Tab = 1
                    cboGestor.SetFocus
                End If

                Exit Function
            End If
        
            If Trim$(txtNumContrato.Text) = "" And cboClaseInstrumento.ListIndex > 0 Then
                MsgBox "No se ha podido encontrar el Número de Contrato. Revise los datos de la Línea. ", vbCritical, Me.Caption

                If txtNumContrato.Enabled Then txtNumContrato.SetFocus
                Exit Function
            End If
    
            If cboMoneda.ListIndex <= 0 Then
                MsgBox "Verifique que la moneda esté ingresada.", vbCritical, Me.Caption
                Exit Function
            End If
    
            If CInt(txtTotalDctosAnexo.Text) <= 0 Then
                MsgBox "Debe ingresar la cantidad de documentos que contiene el anexo.", vbCritical, Me.Caption
                Exit Function
            End If
    
            If CInt(txtDiasPlazo.Text) <= 0 Then
                MsgBox "Verifique que el plazo esté ingresado.", vbCritical, Me.Caption
                Exit Function
            End If
    
            If cboResponsablePago.ListIndex <= 0 Then
                MsgBox "Debe especificar quién hará el pago al vencimiento.", vbCritical, Me.Caption

                If cboResponsablePago.Enabled And tabRFCortoPlazo.TabEnabled(1) = True Then
                    tabRFCortoPlazo.Tab = 1
                    cboResponsablePago.SetFocus
                End If

                Exit Function
            End If
    
            'Si no están estos datos: no necesariamente significa error, puede no considerarse comisiones
            'para la ooeración. Sólo hacer un warning
    
    End Select

    ValidaRequisitosTab = True

End Function

Private Sub tdgConsulta_FormatText(ByVal ColIndex As Integer, _
                                   Value As Variant, _
                                   Bookmark As Variant)

    If ColIndex = 6 Then
        Call DarFormatoValor(Value, Decimales_Monto)
    End If
    
    If ColIndex = 7 Then
        Call DarFormatoValor(Value, Decimales_Tasa)
    End If
    
    If ColIndex = 8 Then
        Call DarFormatoValor(Value, Decimales_Monto)
    End If
    
End Sub
'
'Private Sub txtCantidad_Change()
'
'    Call FormatoCajaTexto(txtCantidad, Decimales_Monto)
'
'End Sub

'Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
'
'    Call ValidaCajaTexto(KeyAscii, "M", txtCantidad, Decimales_Monto)
'
'End Sub

Private Sub txtComisionAgente_Change(Index As Integer)

    Call FormatoCajaTexto(txtComisionAgente(Index), Decimales_Monto)
    
    If strPersonalizaComision = "NO" Then
        lblComisionIgv(0).Caption = (CDbl(txtPorcenIgv(0).Text) / 100) * CDbl(txtComisionAgente(0).Text)
    Else
        'Se carga la comisión  y porcentaje automáticamente
        ActualizaPorcentaje txtComisionAgente(Index), txtPorcenAgente(Index)
        lblComisionIgv(0).Caption = (CDbl(txtPorcenIgv(0).Text) / 100) * CDbl(txtComisionAgente(0).Text)
    End If

End Sub

Private Sub txtComisionAgente_KeyPress(Index As Integer, KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtComisionAgente(Index), Decimales_Monto)

    If KeyAscii = vbKeyReturn Then
        If strPersonalizaComision = "NO" Then
            lblComisionIgv(0).Caption = (CDbl(txtPorcenIgv(0).Text) / 100) * CDbl(txtComisionAgente(0).Text)
            Call CalculoTotal(Index)
        Else
            ActualizaPorcentaje txtComisionAgente(Index), txtPorcenAgente(Index)
            lblComisionIgv(0).Caption = (CDbl(txtPorcenIgv(0).Text) / 100) * CDbl(txtComisionAgente(0).Text)
            Call CalculoTotal(Index)
        End If
    End If
    
End Sub

Private Sub txtComisionBolsa_Change(Index As Integer)

    Call FormatoCajaTexto(txtComisionBolsa(Index), Decimales_Monto)

    ActualizaPorcentaje txtComisionBolsa(Index), lblPorcenBolsa(Index)
    
End Sub

Private Sub txtComisionBolsa_KeyPress(Index As Integer, KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtComisionBolsa(Index), Decimales_Monto)
    
    If KeyAscii = vbKeyReturn Then
        ActualizaPorcentaje txtComisionBolsa(Index), lblPorcenBolsa(Index)
        Call CalculoTotal(Index)
    End If
    
End Sub

Private Sub CalculoTotal(Index As Integer)

    Dim curComImp As Currency, curMonTotal As Currency
    Dim curIntImp As Currency

    If Not IsNumeric(txtComisionAgente(Index).Text) Or Not IsNumeric(txtComisionBolsa(Index).Text) Or Not IsNumeric(txtInteresCorrido(Index).Text) Then Exit Sub
    
    txtImptoInteres(Index).Text = Round(CCur(CDbl(lblIntAdelantado(Index).Caption) * (CDbl(txtPorcenIgvInt(Index).Value)) / 100), 2)
    
    txtIGVCobroMinimoInteres.Text = Round(CCur(CDbl(txtCobroMinimoInteres.Text) * (CDbl(txtPorcenIgvInt(Index).Value)) / 100), 2)
    
    txtImptoInteresAdic(Index).Text = Round(CCur(CDbl(txtIntAdicional(Index).Text) * (CDbl(txtPorcenIgvInt(Index).Value)) / 100), 2)
    curIntImp = CCur(txtImptoInteres(Index).Text) + CCur(txtImptoInteresAdic(Index).Text)
        
    curComImp = CCur(CCur(txtComisionAgente(Index).Text) + CCur(txtComisionBolsa(Index).Text))
    
    lblComisionIgv(Index).Caption = CStr(curComImp) * CDbl(txtPorcenIgv(0).Text) / 100
    lblComisionIgvInt(Index).Caption = CStr(curIntImp)
     
    'Calculando todo lo que afecta al monto subtotal
    If strCodCobroInteres = Codigo_Modalidad_Pago_Adelantado Then
        curComImp = CCur(txtComisionAgente(Index).Text) + CCur(txtComisionBolsa(Index).Text) + CCur(lblIntAdelantado(Index).Caption) + CCur(txtIntAdicional(Index).Text) + CCur(lblComisionIgv(Index).Caption) + CCur(lblComisionIgvInt(Index).Caption)
    Else
        curComImp = CCur(txtComisionAgente(Index).Text) + CCur(txtComisionBolsa(Index).Text) + CCur(lblComisionIgv(Index).Caption)
    End If

    lblComisionesResumen(Index).Caption = CStr(curComImp)

    If Index = 0 Then
        curMonTotal = CCur(txtValorNominalDcto.Text) - curComImp
    Else
        curMonTotal = CCur(txtValorNominalDcto.Text) + curComImp
    End If
    
    lblMontoTotal(Index).Caption = CStr(curMonTotal)
        
    If Trim$(txtValorNominalDcto.Text) <> "" Then
        txtMontoVencimiento1.Text = CDbl(txtValorNominalDcto.Text) ' * CCur(txtCantidad.Text)
    End If
    
    If strCodCobroInteres = Codigo_Modalidad_Pago_Vencimiento Then
        txtMontoVencimiento1.Text = (CDbl(txtMontoVencimiento1.Value) + CDbl(lblIntAdelantado(Index).Caption) + CDbl(txtIntAdicional(Index).Text) + curIntImp)
    End If
            
End Sub

Private Sub txtDiasPlazo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then  'ACC 05/04/2010  Agregado
    
        If IsNumeric(txtDiasPlazo.Text) Then
            dtpFechaVencimiento.Value = DateAdd("d", txtDiasPlazo.Text, CVDate(dtpFechaOrden.Value))
        Else
            dtpFechaVencimiento.Value = dtpFechaOrden.Value
        End If
        
        dtpFechaVencimiento_Change
        
        lblDiasPlazo.Caption = CStr(txtDiasPlazo.Text)

        If dtpFechaVencimientoDcto.Value <> dtpFechaVencimiento.Value Then
            dtpFechaVencimientoDcto.Value = dtpFechaVencimiento.Value
        End If
    
    End If
    
End Sub

Private Sub txtDiasPlazo_LostFocus()

    txtDiasPlazo_KeyPress (vbKeyReturn)
    cboEmisor_Click
    
End Sub

Private Sub AsignaComision(strTipoComision As String, _
                           dblValorComision As Double, _
                           ctrlValorComision As Control)
    
    'If Not IsNumeric(lblSubTotal(ctrlValorComision.Index).Caption) Then Exit Sub
    
'    If dblValorComision > 0 Then
'        ctrlValorComision.Text = CStr(CCur(lblSubTotal(ctrlValorComision.Index)) * dblValorComision / 100)
'    End If
            
End Sub

Private Sub ActualizaPorcentaje(ctrlComision As Control, ctrlPorcentaje As Control)

'    If Not IsNumeric(ctrlComision) Then Exit Sub ' Or Not IsNumeric(lblSubTotal(ctrlComision.Index).Caption) Then Exit Sub
'
''    If CCur(lblSubTotal(ctrlComision.Index)) = 0 Then
''     '   ctrlPorcentaje = "0"
''    Else
'
'        If CCur(ctrlComision) > 0 Then
'            ctrlPorcentaje = CStr((CCur(ctrlComision) / CCur(lblSubTotal(ctrlComision.Index).Caption)) * 100)
'        Else
'   '         ctrlPorcentaje = "0"
'        End If
'  '  End If
                
End Sub


Private Sub txtIntAdicional_KeyPress(Index As Integer, KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtIntAdicional(Index), Decimales_Monto)
    
    If KeyAscii = vbKeyReturn Then
        Call CalculoTotal(Index)
    End If

End Sub

Private Sub txtInteresCorrido_Change(Index As Integer)

    Call FormatoCajaTexto(txtInteresCorrido(Index), Decimales_Monto)
    
    If Trim$(txtInteresCorrido(Index).Text) <> Valor_Caracter Then
        lblInteresesResumen(Index).Caption = CStr(CCur(txtInteresCorrido(Index).Text))
        txtImptoInteresCorrido(Index).Text = (txtInteresCorrido(Index).Text * txtPorcenIgvInt(Index).Text / 100)
    End If
    
End Sub

Private Sub txtInteresCorrido_KeyPress(Index As Integer, KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtInteresCorrido(Index), Decimales_Monto)
    
    If KeyAscii = vbKeyReturn Then
        txtImptoInteresCorrido(Index).Text = (txtInteresCorrido(Index).Text * txtPorcenIgvInt(Index).Text / 100)
        Call CalculoTotal(Index)
    End If
    
End Sub

Private Sub ActualizaComision(ctrlPorcentaje As Control, ctrlComision As Control)

'    If Not IsNumeric(lblSubTotal(ctrlComision.Index).Caption) Or Not IsNumeric(ctrlPorcentaje) Then Exit Sub
'
'    If CDbl(ctrlPorcentaje) > 0 Then
'        ctrlComision = CStr(CCur(lblSubTotal(ctrlComision.Index).Caption) * CDbl(ctrlPorcentaje) / 100)
'    Else
'        ctrlComision = "0"
'    End If
'
End Sub

Private Sub txtNemonico_Change()

    txtDescripOrden = Trim$(txtNemonico.Text)
    
End Sub

Private Sub txtNemonico_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    
End Sub

Private Sub txtNumDocDscto_Change()

    'Asignando el nemónico
    txtNemonico.Text = GenerarNemonico(strCodTipoInstrumentoOrden, strCodClaseInstrumento, strCodEmisor, txtNumDocDscto.Text)

End Sub

Private Sub txtNumDocDscto_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then

        'Asignando el nemónico
        txtNemonico.Text = GenerarNemonico(strCodTipoInstrumentoOrden, strCodClaseInstrumento, strCodEmisor, txtNumDocDscto.Text)
        
    End If

End Sub

Private Sub txtPorcenAgente_Change(Index As Integer)

    If strPersonalizaComision <> "NO" Then
    '    Call FormatoCajaTexto(txtPorcenAgente(Index), Decimales_Tasa)
    End If
    
End Sub

Private Sub txtPorcenAgente_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then

        If strPersonalizaComision <> "NO" Then
           ' Call ValidaCajaTexto(KeyAscii, "M", txtPorcenAgente(Index), Decimales_Tasa)
            ActualizaComision txtPorcenAgente(Index), txtComisionAgente(Index)
            Call CalculoTotal(Index)
        End If

    End If
        
End Sub

Private Sub txtPorcenAgente_LostFocus(Index As Integer)
    CalcularComision
End Sub
 
Private Sub txtPorcenDctoValorNominal_Change()

    Call txtValorNominal_Change
    
End Sub

Private Sub txtPorcenIgv_Change(Index As Integer)

    ActualizaComision txtPorcenIgv(Index), lblComisionIgv(Index)

    If Trim$(txtPorcenIgv(0).Text) <> Null Then
        lblComisionIgv(0).Caption = (CDbl(txtPorcenIgv(0).Text) / 100) * CDbl(txtComisionAgente(0).Text)
    End If

    Call CalculoTotal(Index)
    
End Sub

Private Sub txtPorcenIgv_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        ActualizaComision txtPorcenIgv(Index), lblComisionIgv(Index)

        If Trim$(txtPorcenIgv(0).Text) <> Null Then
            lblComisionIgv(0).Caption = (CDbl(txtPorcenIgv(0).Text) / 100) * CDbl(txtComisionAgente(0).Text)
        End If

        Call CalculoTotal(Index)
    End If

End Sub

Private Sub txtPorcenIgvInt_Change(Index As Integer)

    ActualizaComision txtPorcenIgvInt(Index), lblComisionIgvInt(Index)
    Call CalculoTotal(Index)

End Sub

Private Sub txtPrecioUnitario_Change(Index As Integer)

'    Call FormatoCajaTexto(txtPrecioUnitario(Index), Decimales_Precio)
'
'    If Not IsNumeric(txtValorNominal.Text) Or Not IsNumeric(txtPrecioUnitario(Index).Text) Or Not IsNumeric(txtDiasPlazo.Text) Then Exit Sub
'    If Not (CDbl(txtValorNominal.Text) > 0 And CDbl(txtPrecioUnitario(Index).Text) > 0 And CInt(txtDiasPlazo.Text) > 0) Then Exit Sub
'
'   ' lblSubTotal(Index).Caption = CDbl(txtValorNominal.Text) * CDbl(txtPrecioUnitario(Index).Text) / 100
'
'    'Aca calcula la TIR, si no es cambio directo
'    If txtPrecioUnitario(Index).Tag = "0" Then
'        txtTirBruta1.Tag = "1"
'        txtTirBruta1.Text = ((CDbl(txtMontoVencimiento1.Value) / (CDbl(txtPrecioUnitario(0).Text) / 100 * CCur(txtCantidad.Text) * CDbl(txtValorNominal.Text))) ^ (360 / CInt(txtDiasPlazo.Text)) - 1) * 100
'    Else
'        txtPrecioUnitario(Index).Tag = "0"
'    End If
'
'    lblPrecioResumen(Index).Caption = CStr(txtPrecioUnitario(Index).Text)
    
End Sub

Private Sub txtPrecioUnitario_KeyPress(Index As Integer, KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtPrecioUnitario(Index), Decimales_Precio)
    
End Sub

Private Sub txtTasa_Change()

    Call FormatoCajaTexto(txtTasa, Decimales_Tasa)
       
    Call CalculoTotal(0)
    
End Sub

Private Sub txtTasa_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtTasa, Decimales_Tasa)
    
End Sub

Private Sub txtTipoCambio_Change()

    Call FormatoCajaTexto(txtTipoCambio, Decimales_TipoCambio)
    
End Sub

Private Sub txtTipoCambio_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtTipoCambio, Decimales_TipoCambio)
    
End Sub

Private Sub txtTirBruta1_Change()
    
'    If Not (txtTirBruta1.Value <> 0 And CInt(txtDiasPlazo.Text) > 0 And CCur(txtCantidad.Text) > 0 And CDbl(txtValorNominalDcto.Text) > 0) Then Exit Sub
'
'    If txtTirBruta1.Tag = "0" Then 'indica cambio directo en la pantalla
'        txtPrecioUnitario1.Tag = "1"
'
'        txtPrecioUnitario1.Text = (CDbl(txtValorNominalDcto.Text) * (1 - ((1 + 0.01 * txtTirBruta1.Value) ^ (CInt(txtDiasPlazo.Text) / intBaseCalculo) - 1))) / (CDbl(txtValorNominalDcto.Text) * CDbl(txtCantidad.Text)) * 100
'    Else
'        txtTirBruta1.Tag = "0"
'    End If

End Sub

Private Sub txtTirNeta_Change()

    Call FormatoCajaTexto(txtTirNeta, Decimales_Tasa)

End Sub

Private Sub txtTirNeta_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtTirNeta, Decimales_Tasa)

End Sub

Private Sub txtTotalDctosAnexo_KeyPress(KeyAscii As Integer)

    If txtTotalDctosAnexo.Text = "" Then txtTotalDctosAnexo.Text = 1

    If KeyAscii = vbKeyReturn Then
        
        Call CalculoTotal(0)
        
    End If

End Sub

Private Sub txtTotalMEAnexo_KeyPress(KeyAscii As Integer)

    If txtTotalMEAnexo.Text = "" Then txtTotalMEAnexo.Text = 0

    If KeyAscii = vbKeyReturn Then
       
        Call CalculoTotal(0)
        
    End If

End Sub

Private Sub txtTotalMEAnexo_LostFocus()

    txtTotalMEAnexo_KeyPress (vbKeyReturn)
    
End Sub

Private Sub txtTotalMNAnexo_KeyPress(KeyAscii As Integer)
    
    If txtTotalMNAnexo.Text = "" Then txtTotalMNAnexo.Text = 0
    
    If KeyAscii = vbKeyReturn Then
 
        Call CalculoTotal(0)
    End If

End Sub

Private Sub txtTotalMNAnexo_LostFocus()

    txtTotalMNAnexo_KeyPress (vbKeyReturn)
    
End Sub

Private Sub txtValorNominal_Change()
  
 '   If Not IsNumeric(txtCantidad.Text) Then Exit Sub
    
    txtValorNominalDcto.Text = CStr(txtPorcenDctoValorNominal.Value / 100 * txtValorNominal.Value)

    
    Call CalcularComision
    Call CalculoTotal(0)
       
End Sub

Private Sub txtValorNominalDcto_Change()

    Call FormatoCajaTexto(txtValorNominalDcto, Decimales_Monto)
    lblIntAdelantado(0).Caption = CalculoInteresDescuento(CDbl(txtTasa.Text), strCodTipoTasa, strPeriodoTasa, strCodBaseAnual, txtValorNominalDcto, CStr(dtpFechaOrden.Value), CStr(dtpFechaPago.Value))
    txtCobroMinimoInteres.Text = CalculoInteresDescuento(CDbl(txtTasa.Text), strCodTipoTasa, strPeriodoTasa, strCodBaseAnual, txtValorNominalDcto, CStr(dtpFechaOrden.Value), CStr(DateAdd("d", intDiasInteresMinimo, dtpFechaOrden.Value)))

End Sub

Public Function CalculoInteresDescuento(numPorcenTasa As Double, _
                                        strCodTipoTasa As String, _
                                        strCodPeriodoTasa As String, _
                                        strCodBaseCalculo As String, _
                                        numMontoBaseCalculo As Double, _
                                        datFechaInicial As Date, _
                                        datFechaFinal As Date) As Double

    Dim intNumPeriodoAnualTasa          As Integer
    Dim intNumPeriodoTasa               As Integer
    Dim intNumPeriodoCapitalizacion     As Integer
    
    Dim dblCantPeriodosCapEnPeriodoTasa As Double
    Dim dblTasaEfectivaEnPeriodoCap     As Double
    Dim dblCantPeriodosCapEnPeriodoReq  As Double
    
    Dim intDiasProvision                As Integer
    Dim intDiasBaseAnual                As Integer
    Dim numPorcenTasaAnual              As Double
    Dim numMontoCalculoInteres          As Double
    Dim adoConsulta                     As ADODB.Recordset
            
    With adoComm
        Set adoConsulta = New ADODB.Recordset
    
        '*** Obtener el número de días del periodo de tasa ***
        .CommandText = "SELECT ValorParametro FROM AuxiliarParametro WHERE CodTipoParametro='TIPFRE' AND CodParametro='" & strCodPeriodoTasa & "'"
        Set adoConsulta = .Execute
    
        If Not adoConsulta.EOF Then
            intNumPeriodoAnualTasa = CInt(360 / adoConsulta("ValorParametro"))     '*** Numero del periodos por año de la tasa ***
            intNumPeriodoTasa = CInt(adoConsulta("ValorParametro"))
        End If

        adoConsulta.Close: Set adoConsulta = Nothing
    End With
   
    Select Case strCodBaseCalculo

        Case Codigo_Base_30_360:
            intDiasBaseAnual = 360
            intDiasProvision = DateDiff("d", datFechaInicial, datFechaFinal)

        Case Codigo_Base_Actual_365:
            intDiasBaseAnual = 365
            intDiasProvision = DateDiff("d", datFechaInicial, datFechaFinal) + 1

        Case Codigo_Base_Actual_360:
            intDiasBaseAnual = 360
            intDiasProvision = DateDiff("d", datFechaInicial, datFechaFinal)

        Case Codigo_Base_30_365:
            intDiasBaseAnual = 365
            intDiasProvision = Dias360(datFechaInicial, datFechaFinal, True)
    End Select
    
    Select Case strCodTipoTasa
 
        Case Codigo_Tipo_Tasa_Efectiva:
            numPorcenTasaAnual = (1 + (numPorcenTasa / 100)) ^ (intNumPeriodoAnualTasa) - 1
            numMontoCalculoInteres = Round(numMontoBaseCalculo * ((((1 + numPorcenTasaAnual)) ^ (intDiasProvision / intDiasBaseAnual)) - 1), 2)
            
        Case Codigo_Tipo_Tasa_Nominal:
            adoComm.CommandText = "SELECT ValorParametro FROM AuxiliarParametro WHERE CodTipoParametro='TIPFRE' AND CodParametro='" & strPeriodoCapitalizable & "'"
            Set adoConsulta = adoComm.Execute
            
            If Not adoConsulta.EOF Then
                intNumPeriodoCapitalizacion = CInt(adoConsulta("ValorParametro"))
            End If
            
            If intNumPeriodoTasa <> 0 And intNumPeriodoCapitalizacion <> 0 Then
                dblCantPeriodosCapEnPeriodoTasa = intNumPeriodoTasa / intNumPeriodoCapitalizacion
                dblTasaEfectivaEnPeriodoCap = (numPorcenTasa / 100) / dblCantPeriodosCapEnPeriodoTasa
                dblCantPeriodosCapEnPeriodoReq = intDiasProvision / intNumPeriodoCapitalizacion
                
                numMontoCalculoInteres = Round(numMontoBaseCalculo * (((1 + dblTasaEfectivaEnPeriodoCap) ^ dblCantPeriodosCapEnPeriodoReq) - 1), 2)
            End If

        Case Codigo_Tipo_Tasa_Flat:
            numPorcenTasaAnual = numPorcenTasa / 100
            numMontoCalculoInteres = Round(numMontoBaseCalculo * (numPorcenTasaAnual), 2)
    End Select

    CalculoInteresDescuento = numMontoCalculoInteres
End Function

Public Sub CargarCabeceraAnexo(strpCodFondoOrden As String, _
                               strpCodAdministradora As String, _
                               strpstrNumAnexo As String, _
                               strpNumOrden As String)
    
    'Obteniendo los datos de la operación cuya cabecera servirá como cabecera de
    'las operaciones posteriores
    
    Dim adoOperacionOrig As ADODB.Recordset
    
    Set adoOperacionOrig = New ADODB.Recordset

    With adoComm

        .CommandText = "SELECT IOP.CodFile, IOP.CodDetalleFile, IOP.CodSubDetalleFile, IOP.CodGirador, IOP.CodMoneda, " & _
                        " IOP.CodObligado, IOP.CodGestor, IOP.CodNegociacion, IOP.CodTitulo, IOP.CodOrigen, IOP.CodMoneda," & _
                        " IOP.PorcenDsctoValorNominal, IOP.CantDiasPlazo, IOP.TipoTasa, IOP.PeriodoTasa, IOP.PeriodoCapitalizacion, IOP.BaseAnual, " & _
                        "IOP.TasaInteres, IOP.ModoCobroInteres, IOP.PorcenImptoInteres, IOP.PorcenImptoComision, " & _
                        "IOP.ResponsablePago, IOP.DiasInteresAdic, " & _
                        "IOP.MontoTotalAnexo,IOP.CantDocumAnexo, IOP.PorcenComision " & _
                        "FROM InversionOrden IOP " & "WHERE IOP.CodFondo='" & strpCodFondoOrden & _
                        "' AND IOP.CodAdministradora='" & strpCodAdministradora & "' AND " & "NumAnexo= '" & strpstrNumAnexo & _
                        "' /*AND NumOrden='" & strpNumOrden & "'*/ and CodEmisor = '" & strCodEmisor & "'"

        Set adoOperacionOrig = .Execute

        If Not adoOperacionOrig.EOF Then
            blnCargarCabeceraAnexo = True
            '*** Fijación de combos con el valor de la operación base ***
            cboFondoOrden.ListIndex = ObtenerItemLista(arrFondoOrden(), strpCodFondoOrden)
            cboTipoInstrumentoOrden.ListIndex = ObtenerItemLista(arrTipoInstrumentoOrden(), Trim$(adoOperacionOrig("CodFile")))
            cboClaseInstrumento.ListIndex = ObtenerItemLista(arrClaseInstrumento(), Trim$(adoOperacionOrig("CodDetalleFile")))
            cboSubClaseInstrumento.ListIndex = ObtenerItemLista(arrSubClaseInstrumento(), Trim$(adoOperacionOrig("CodSubDetalleFile")))
            cboEmisor.ListIndex = ObtenerItemLista(arrEmisor(), Trim$(adoOperacionOrig("CodGirador")))
            cboGestor.ListIndex = ObtenerItemLista(arrGestor(), Trim$(adoOperacionOrig("CodGestor")))
            cboMoneda.ListIndex = ObtenerItemLista(arrMoneda(), Trim$(adoOperacionOrig("CodMoneda")))
            'puesto que el codigo de operación (directa, etc) no se guarda en operaciones, se trae de órdenes
            'strCodOperacionOrden = traerCampo("InversionOrden", "RTrim(CodOperacion)", "CodFondo", strpCodFondoOrden, "AND  CodAdministradora = '" & strpCodAdministradora & "' AND NumOrden = '" & strpNumOrden & "'")
                    
            '*** Datos del ANexo ***
            txtTotalMNAnexo.Text = CDbl(adoOperacionOrig("MontoTotalAnexo"))
            txtTotalDctosAnexo.Text = CDbl(adoOperacionOrig("CantDocumAnexo"))
            txtPorcenAgente(0).Text = CDbl(adoOperacionOrig("PorcenComision"))
            
            '*** Datos de la Orden ***
            If CDbl(adoOperacionOrig("DiasInteresAdic")) > 0 Then   'Hubo días adicionales
                chkDiasAdicional.Value = Checked
                lblFechaVencimientoAdic.Visible = True
            Else
                chkDiasAdicional.Value = Unchecked
                lblFechaVencimientoAdic.Visible = False
            End If
                    
            '*** Valores de negociación ***
            txtTasa.Text = Trim$(adoOperacionOrig("TasaInteres"))
            cboTipoTasa.ListIndex = ObtenerItemLista(arrTipoTasa(), Trim$(adoOperacionOrig("TipoTasa")))
            cboBaseAnual.ListIndex = ObtenerItemLista(arrBaseAnual(), Trim$(adoOperacionOrig("BaseAnual")))
            cboCobroInteres.ListIndex = ObtenerItemLista(arrPagoInteres(), "MODPAG" & Trim$(adoOperacionOrig("ModoCobroInteres")))
            cboPeriodoCapitalizacion.ListIndex = ObtenerItemLista(arrPeriodoCapitalizacion(), Trim$(adoOperacionOrig("PeriodoCapitalizacion")))
            cboPeriodoTasa.ListIndex = ObtenerItemLista(arrPeriodoTasa(), Trim$(adoOperacionOrig("PeriodoTasa")))
            txtPorcenDctoValorNominal.Text = Trim$(adoOperacionOrig("PorcenDsctoValorNominal"))
            cboNegociacion.ListIndex = ObtenerItemLista(arrNegociacion(), Trim$(adoOperacionOrig("CodNegociacion")))

            Call HabilitaCabeceraAnexo(False)
            
            tabRFCortoPlazo.Tab = 1
            txtNumDocDscto.SetFocus
        Else
            MsgBox "Ha ocurrido un error al traer los datos básicos del Anexo. Deberá indicarlos ud. mismo. ", vbCritical
            Call HabilitaCabeceraAnexo(True)
            adoOperacionOrig.Close: Set adoOperacionOrig = Nothing
            Exit Sub
        End If
    
        adoOperacionOrig.Close: Set adoOperacionOrig = Nothing

    End With
            
End Sub

Public Sub HabilitaCabeceraAnexo(blnAccion As Boolean)

    Call HabilitaCombos(blnAccion)
    chkDiasAdicional.Enabled = blnAccion
    txtTasa.Locked = Not blnAccion
    'txtPorcenDctoValorNominal.Locked = Not blnAccion
    txtTotalMNAnexo.Locked = Not blnAccion
    txtTotalMEAnexo.Locked = Not blnAccion
    txtTotalDctosAnexo.Locked = Not blnAccion
    'txtPorcenDctoValorNominal.Locked = Not blnAccion
End Sub

Public Sub HabilitaCombos(ByVal pBloquea As Boolean)

   ' cboFondoOrden.E
    cboTipoInstrumentoOrden.Enabled = pBloquea
    cboClaseInstrumento.Enabled = pBloquea
    cboSubClaseInstrumento.Enabled = pBloquea
    cboEmisor.Enabled = pBloquea
    cboGestor.Enabled = pBloquea
    cboMoneda.Enabled = pBloquea
    cboTipoTasa.Enabled = pBloquea
    cboPeriodoTasa.Enabled = pBloquea
    cboPeriodoCapitalizacion.Enabled = pBloquea
    cboBaseAnual.Enabled = pBloquea
    cboCobroInteres.Enabled = pBloquea
   ' cboLineaCliente.Enabled = pBloquea

End Sub

Private Sub txtValorNominalDescuento_Change()
    ActualizarTotalAnexo
End Sub

Private Sub txtValorNominalDocumento_Change()
   ActualizarValorNominalTC
End Sub

Private Sub txtTipoCambioDescuento_Change()
   ActualizarValorNominalTC
End Sub

Private Sub ActualizarValorNominalTC()
    Dim dblTempTipoCambio As Double
    
    If Not IsNumeric(txtValorNominalDocumento.Text) Then Exit Sub
    If Not IsNumeric(txtTipoCambioDescuento.Text) Then Exit Sub
    
    If strCodMonedaDocumento = "02" And strCodMoneda = "01" Then
        dblTempTipoCambio = CDec(txtTipoCambioDescuento.Text)
    End If
        
    If strCodMonedaDocumento = "01" And strCodMoneda = "02" Then
        dblTempTipoCambio = 1 / CDec(txtTipoCambioDescuento.Text)
    End If

    If strCodMonedaDocumento = strCodMoneda Then
        dblTempTipoCambio = 1
    End If
    
    txtValorNominalDescuento.Text = CDec(txtValorNominalDocumento.Text) * dblTempTipoCambio
End Sub

