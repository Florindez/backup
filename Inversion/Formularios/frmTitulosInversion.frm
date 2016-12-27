VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{830D5A06-9B70-4F8C-98B6-7A19C4E7760B}#1.0#0"; "TAMControls.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmTitulosInversion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Definición de Instrumentos de Inversión"
   ClientHeight    =   8310
   ClientLeft      =   900
   ClientTop       =   1125
   ClientWidth     =   11325
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
   Icon            =   "frmTitulosInversion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8310
   ScaleWidth      =   11325
   Begin TAMControls2.ucBotonEdicion2 cmdOpcion 
      Height          =   735
      Left            =   600
      TabIndex        =   117
      Top             =   7440
      Width           =   5700
      _ExtentX        =   7408
      _ExtentY        =   1296
      Buttons         =   3
      Caption0        =   "&Nuevo"
      Tag0            =   "0"
      Visible0        =   0   'False
      ToolTipText0    =   "Nuevo"
      Caption1        =   "&Modificar"
      Tag1            =   "3"
      Visible1        =   0   'False
      ToolTipText1    =   "Modificar"
      Caption2        =   "&Buscar"
      Tag2            =   "5"
      Visible2        =   0   'False
      ToolTipText2    =   "Buscar"
      UserControlWidth=   5700
   End
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   9500
      TabIndex        =   116
      Top             =   7440
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
      Caption0        =   "&Salir"
      Tag0            =   "9"
      Visible0        =   0   'False
      ToolTipText0    =   "Salir"
      UserControlWidth=   1200
   End
   Begin TabDlg.SSTab tabTitulos 
      Height          =   7215
      Left            =   0
      TabIndex        =   62
      Top             =   30
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   12726
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Lista"
      TabPicture(0)   =   "frmTitulosInversion.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "tdgConsulta"
      Tab(0).Control(1)=   "tabCriterio"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Datos Generales"
      TabPicture(1)   =   "frmTitulosInversion.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "cmdAccion"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraDatos(0)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdAccion2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Cronograma"
      TabPicture(2)   =   "frmTitulosInversion.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraDatos(1)"
      Tab(2).ControlCount=   1
      Begin TAMControls2.ucBotonEdicion2 cmdAccion2 
         Height          =   735
         Left            =   8040
         TabIndex        =   115
         Top             =   6240
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   1296
         Buttons         =   2
         Caption0        =   "&Grabar"
         Tag0            =   "2"
         Visible0        =   0   'False
         ToolTipText0    =   "Grabar"
         Caption1        =   "&Cancelar"
         Tag1            =   "8"
         Visible1        =   0   'False
         ToolTipText1    =   "Cancelar"
         UserControlWidth=   2700
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmTitulosInversion.frx":0496
         Height          =   2505
         Left            =   -74640
         OleObjectBlob   =   "frmTitulosInversion.frx":04B0
         TabIndex        =   107
         Top             =   3660
         Width           =   10590
      End
      Begin VB.Frame fraDatos 
         Height          =   5445
         Index           =   0
         Left            =   360
         TabIndex        =   75
         Top             =   690
         Width           =   10635
         Begin VB.ComboBox cboMercado 
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
            Left            =   1950
            Style           =   2  'Dropdown List
            TabIndex        =   113
            Top             =   4020
            Width           =   7365
         End
         Begin VB.TextBox txtValorNominal 
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
            Left            =   1950
            TabIndex        =   29
            Top             =   4845
            Width           =   2685
         End
         Begin VB.CheckBox chkQuiebre 
            Caption         =   "Quebrable"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5010
            TabIndex        =   30
            Top             =   4860
            Width           =   1695
         End
         Begin VB.TextBox txtCodigoValor 
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
            Left            =   1950
            MaxLength       =   15
            TabIndex        =   19
            Top             =   1680
            Width           =   2685
         End
         Begin VB.ComboBox cboClaseInstrumento 
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
            Left            =   1950
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   720
            Width           =   2685
         End
         Begin VB.ComboBox cboSubClaseInstrumento 
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
            Left            =   6300
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   720
            Width           =   3975
         End
         Begin VB.TextBox txtNumEmision 
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
            Left            =   1950
            MaxLength       =   12
            TabIndex        =   22
            Top             =   2370
            Width           =   2685
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1950
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   360
            Width           =   2685
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1950
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   2880
            Width           =   7365
         End
         Begin VB.TextBox txtDescripValor 
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
            Left            =   1950
            MaxLength       =   60
            TabIndex        =   21
            Top             =   2025
            Width           =   8355
         End
         Begin VB.TextBox txtNumSerie 
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
            Left            =   6300
            MaxLength       =   12
            TabIndex        =   23
            Top             =   2370
            Width           =   2685
         End
         Begin VB.TextBox txtNemonico 
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
            Left            =   6300
            MaxLength       =   15
            TabIndex        =   20
            Top             =   1680
            Width           =   2685
         End
         Begin VB.ComboBox cboMonedaEmision 
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
            Left            =   1950
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   4440
            Width           =   2685
         End
         Begin VB.ComboBox cboMonedaPago 
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
            Left            =   6300
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   4440
            Width           =   3015
         End
         Begin VB.OptionButton optTipoCodigo 
            Caption         =   "ISIN/BVL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   230
            Index           =   0
            Left            =   1920
            TabIndex        =   17
            Top             =   1320
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optTipoCodigo 
            Caption         =   "Autogenerado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   230
            Index           =   1
            Left            =   3480
            TabIndex        =   18
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Mercado Cotiza"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   38
            Left            =   360
            TabIndex        =   112
            Top             =   4080
            Width           =   1110
         End
         Begin VB.Label lblRiesgo 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6630
            TabIndex        =   31
            Top             =   5490
            Width           =   2685
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Valor Nominal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   23
            Left            =   360
            TabIndex        =   108
            Top             =   4860
            Width           =   975
         End
         Begin VB.Label lblAnalitica 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6300
            TabIndex        =   14
            Top             =   360
            Width           =   2685
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Analítica"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   12
            Left            =   5010
            TabIndex        =   90
            Top             =   375
            Width           =   630
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Mnemotécnico"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   5010
            TabIndex        =   89
            Top             =   1725
            Width           =   1050
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Código Unico"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   88
            Top             =   1725
            Width           =   960
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Nombre del Valor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   87
            Top             =   2070
            Width           =   1215
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Num. Emisión"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   17
            Left            =   360
            TabIndex        =   86
            Top             =   2415
            Width           =   960
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Num. Serie"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   18
            Left            =   5010
            TabIndex        =   85
            Top             =   2415
            Width           =   780
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Instrumento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   7
            Left            =   360
            TabIndex        =   84
            Top             =   380
            Width           =   1185
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Clase"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   8
            Left            =   360
            TabIndex        =   83
            Top             =   740
            Width           =   390
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "SubClase"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   9
            Left            =   5010
            TabIndex        =   82
            Top             =   735
            Width           =   675
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Emisor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   360
            TabIndex        =   81
            Top             =   2910
            Width           =   465
         End
         Begin VB.Label lblGrupoEconomico 
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
            Height          =   315
            Left            =   1950
            TabIndex        =   25
            Top             =   3240
            Width           =   7365
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Grupo Económico"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   360
            TabIndex        =   80
            Top             =   3270
            Width           =   1275
         End
         Begin VB.Label lblCiiu 
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
            Left            =   1950
            TabIndex        =   26
            Top             =   3600
            Width           =   7365
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "CIIU"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   16
            Left            =   360
            TabIndex        =   79
            Top             =   3630
            Width           =   315
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Moneda Emisión"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   10
            Left            =   360
            TabIndex        =   78
            Top             =   4470
            Width           =   1170
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Moneda Pago"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   11
            Left            =   5010
            TabIndex        =   77
            Top             =   4500
            Width           =   1005
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Clasificación de Riesgo Equivalente CONASEV"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   6
            Left            =   1950
            TabIndex        =   76
            Top             =   5535
            Width           =   3585
         End
      End
      Begin TabDlg.SSTab tabCriterio 
         Height          =   3105
         Left            =   -74640
         TabIndex        =   65
         Top             =   420
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   5477
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Tipo / Código Unico"
         TabPicture(0)   =   "frmTitulosInversion.frx":3F20
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fraCriterioTipo"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Rango de Fechas"
         TabPicture(1)   =   "frmTitulosInversion.frx":3F3C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fraCriterioFecha"
         Tab(1).ControlCount=   1
         Begin VB.Frame fraCriterioTipo 
            Height          =   2460
            Left            =   360
            TabIndex        =   69
            Top             =   440
            Width           =   9825
            Begin VB.CheckBox chkSeleccionTipo 
               Caption         =   "Seleccionar"
               Height          =   255
               Left            =   840
               TabIndex        =   0
               Top             =   240
               Width           =   1455
            End
            Begin VB.ComboBox cboClaseCriterio 
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
               Left            =   3270
               Style           =   2  'Dropdown List
               TabIndex        =   2
               Top             =   952
               Width           =   5880
            End
            Begin VB.ComboBox cboGrupoCriterio 
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
               Left            =   3270
               Style           =   2  'Dropdown List
               TabIndex        =   3
               Top             =   1304
               Width           =   5880
            End
            Begin VB.ComboBox cboTipoCriterio 
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
               Left            =   3270
               Style           =   2  'Dropdown List
               TabIndex        =   1
               Top             =   600
               Width           =   5880
            End
            Begin VB.TextBox txtIsinCriterio 
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
               Left            =   3270
               TabIndex        =   4
               Top             =   1680
               Width           =   4620
            End
            Begin VB.TextBox txtNemotecnicoCriterio 
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
               Left            =   3270
               TabIndex        =   5
               Top             =   1980
               Width           =   4620
            End
            Begin VB.Label lblDescrip 
               AutoSize        =   -1  'True
               Caption         =   "Tipo de Instrumento"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   24
               Left            =   840
               TabIndex        =   74
               Top             =   615
               Width           =   1410
            End
            Begin VB.Label lblDescrip 
               AutoSize        =   -1  'True
               Caption         =   "Clase de Instrumento"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   25
               Left            =   840
               TabIndex        =   73
               Top             =   975
               Width           =   1485
            End
            Begin VB.Label lblDescrip 
               AutoSize        =   -1  'True
               Caption         =   "SubClase de Instrumento"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   26
               Left            =   840
               TabIndex        =   72
               Top             =   1320
               Width           =   1770
            End
            Begin VB.Label lblDescrip 
               AutoSize        =   -1  'True
               Caption         =   "Código Isin"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   27
               Left            =   840
               TabIndex        =   71
               Top             =   1680
               Width           =   780
            End
            Begin VB.Label lblDescrip 
               AutoSize        =   -1  'True
               Caption         =   "Código Nemotécnico"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   28
               Left            =   840
               TabIndex        =   70
               Top             =   1995
               Width           =   1485
            End
         End
         Begin VB.Frame fraCriterioFecha 
            Height          =   2460
            Left            =   -74640
            TabIndex        =   66
            Top             =   440
            Width           =   8895
            Begin VB.CheckBox chkSeleccionFecha 
               Caption         =   "Seleccionar"
               Height          =   255
               Left            =   1440
               TabIndex        =   6
               Top             =   360
               Width           =   1455
            End
            Begin VB.CheckBox chkPago 
               Caption         =   "Fecha de Pago"
               Height          =   255
               Left            =   4800
               TabIndex        =   12
               Top             =   1920
               Width           =   2295
            End
            Begin VB.CheckBox chkCorte 
               Caption         =   "Fecha de Corte"
               Height          =   255
               Left            =   4800
               TabIndex        =   10
               Top             =   1440
               Width           =   2295
            End
            Begin VB.CheckBox chkVencimiento 
               Caption         =   "Fecha de Vencimiento"
               Height          =   255
               Left            =   1440
               TabIndex        =   11
               Top             =   1920
               Width           =   2295
            End
            Begin VB.CheckBox chkEmision 
               Caption         =   "Fecha de Emisión"
               Height          =   255
               Left            =   1440
               TabIndex        =   9
               Top             =   1440
               Width           =   2295
            End
            Begin MSComCtl2.DTPicker dtpFechaDesdeCriterio 
               Height          =   285
               Left            =   2400
               TabIndex        =   7
               Top             =   840
               Width           =   1575
               _ExtentX        =   2778
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
               Format          =   175570945
               CurrentDate     =   38765
            End
            Begin MSComCtl2.DTPicker dtpFechaHastaCriterio 
               Height          =   285
               Left            =   5760
               TabIndex        =   8
               Top             =   840
               Width           =   1575
               _ExtentX        =   2778
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
               Format          =   175570945
               CurrentDate     =   38765
            End
            Begin VB.Label lblDescrip 
               AutoSize        =   -1  'True
               Caption         =   "Hasta"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   30
               Left            =   4800
               TabIndex        =   68
               Top             =   855
               Width           =   540
            End
            Begin VB.Label lblDescrip 
               AutoSize        =   -1  'True
               Caption         =   "Desde"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   29
               Left            =   1440
               TabIndex        =   67
               Top             =   855
               Width           =   585
            End
         End
      End
      Begin VB.Frame fraDatos 
         Height          =   5685
         Index           =   1
         Left            =   -74640
         TabIndex        =   63
         Top             =   750
         Width           =   10215
         Begin VB.ComboBox cboFormaCalculo 
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
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   4440
            Width           =   2415
         End
         Begin VB.CommandButton cmdGeneraCuponera 
            Caption         =   "&Tabla"
            Height          =   375
            Index           =   3
            Left            =   8400
            TabIndex        =   61
            ToolTipText     =   "Visualizar la Tabla de Desarrollo"
            Top             =   4840
            Width           =   1000
         End
         Begin VB.ComboBox cboTipoAjuste 
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
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   1800
            Width           =   2415
         End
         Begin VB.TextBox txtTasaAnual 
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
            Left            =   1980
            TabIndex        =   39
            Top             =   2940
            Width           =   2055
         End
         Begin MSComCtl2.DTPicker dtpFechaEmision 
            Height          =   285
            Left            =   7200
            TabIndex        =   45
            Top             =   360
            Width           =   2175
            _ExtentX        =   3836
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
            Format          =   175570945
            CurrentDate     =   38768
         End
         Begin VB.ComboBox cboPeriodoPago 
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
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   4065
            Width           =   2415
         End
         Begin VB.ComboBox cboBaseCalculo 
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
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   3720
            Width           =   2415
         End
         Begin VB.CheckBox chkCapitalizable 
            Caption         =   "Capitalizable"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   360
            TabIndex        =   44
            Top             =   4920
            Width           =   1320
         End
         Begin VB.ComboBox cboTasa 
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
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   1100
            Width           =   2415
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
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   3375
            Width           =   2415
         End
         Begin VB.ComboBox cboCuponCalculo 
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
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   2520
            Width           =   2415
         End
         Begin VB.ComboBox cboTipoVac 
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
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   2160
            Width           =   2415
         End
         Begin VB.CheckBox chkAjuste 
            Caption         =   "Con Ajuste"
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
            Left            =   360
            TabIndex        =   35
            Top             =   1440
            Width           =   1455
         End
         Begin VB.CommandButton cmdGeneraCuponera 
            Caption         =   "&Monto"
            Height          =   375
            Index           =   2
            Left            =   7200
            TabIndex        =   60
            ToolTipText     =   "Generar las Tasas y Montos"
            Top             =   4840
            Width           =   1000
         End
         Begin VB.CheckBox chkCuponCero 
            Caption         =   "Instrumento con Cupón Cero"
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
            Left            =   360
            TabIndex        =   32
            Top             =   330
            Width           =   2535
         End
         Begin VB.CheckBox chkAmortiza 
            Caption         =   "Instrumento con Amortización"
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
            Left            =   360
            TabIndex        =   33
            Top             =   720
            Width           =   2535
         End
         Begin VB.CommandButton cmdGeneraCuponera 
            Caption         =   "&Pago"
            Height          =   375
            Index           =   1
            Left            =   6000
            TabIndex        =   59
            ToolTipText     =   "Generar las Fechas de Pago"
            Top             =   4840
            Width           =   1000
         End
         Begin VB.CommandButton cmdGeneraCuponera 
            Caption         =   "Cor&te"
            Height          =   375
            Index           =   0
            Left            =   4800
            TabIndex        =   58
            ToolTipText     =   "Generar las Fechas de Corte"
            Top             =   4840
            Width           =   1000
         End
         Begin VB.Frame fraGeneraCuponera 
            Caption         =   "Generación de Cuponera"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3240
            Left            =   4875
            TabIndex        =   64
            Top             =   1500
            Width           =   4515
            Begin VB.CommandButton cmdRegenerar 
               Caption         =   "&Generar"
               Height          =   375
               Left            =   3340
               TabIndex        =   109
               ToolTipText     =   "Habilitar botones para generación de cuponera"
               Top             =   2720
               Width           =   1000
            End
            Begin VB.TextBox txtNumDiasPeriodo 
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
               Left            =   2805
               TabIndex        =   51
               Top             =   960
               Width           =   1065
            End
            Begin VB.TextBox txtNumDiasPago 
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
               Left            =   1680
               TabIndex        =   56
               Top             =   2440
               Width           =   1290
            End
            Begin VB.ComboBox cboTipoDia 
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
               TabIndex        =   57
               Top             =   2760
               Width           =   1560
            End
            Begin VB.OptionButton optDiasCupon 
               Caption         =   "Num. Días del Periodo"
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
               Left            =   360
               TabIndex        =   50
               Top             =   960
               Width           =   1935
            End
            Begin VB.CheckBox chk1erCupon 
               Caption         =   "Corte del 1er. Cupón"
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
               Left            =   360
               TabIndex        =   54
               Top             =   1760
               Width           =   2055
            End
            Begin VB.CheckBox chkAPartir 
               Caption         =   "A partir de esta Fecha"
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
               Left            =   360
               TabIndex        =   52
               Top             =   1440
               Width           =   1935
            End
            Begin VB.OptionButton optDiasCupon 
               Caption         =   "Fin de Mes"
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
               Index           =   2
               Left            =   360
               TabIndex        =   49
               Top             =   660
               Width           =   1935
            End
            Begin VB.OptionButton optDiasCupon 
               Caption         =   "Mes Calendario"
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
               Left            =   360
               TabIndex        =   48
               Top             =   360
               Value           =   -1  'True
               Width           =   1455
            End
            Begin MSComCtl2.DTPicker dtpFechaInicioCupon 
               Height          =   285
               Left            =   2805
               TabIndex        =   53
               Top             =   1440
               Width           =   1335
               _ExtentX        =   2355
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
               Format          =   175570945
               CurrentDate     =   38768
            End
            Begin MSComCtl2.DTPicker dtpFechaCorteInicial 
               Height          =   285
               Left            =   2805
               TabIndex        =   55
               Top             =   1755
               Width           =   1335
               _ExtentX        =   2355
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
               Format          =   175570945
               CurrentDate     =   38768
            End
            Begin MSComCtl2.UpDown updDiasPago 
               Height          =   285
               Left            =   2970
               TabIndex        =   104
               Top             =   2445
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   393216
               BuddyControl    =   "txtNumDiasPago"
               BuddyDispid     =   196660
               OrigLeft        =   3510
               OrigTop         =   2440
               OrigRight       =   3765
               OrigBottom      =   2725
               SyncBuddy       =   -1  'True
               BuddyProperty   =   65547
               Enabled         =   -1  'True
            End
            Begin MSComCtl2.UpDown updDiasPeriodo 
               Height          =   285
               Left            =   3870
               TabIndex        =   105
               Top             =   960
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   393216
               BuddyControl    =   "txtNumDiasPeriodo"
               BuddyDispid     =   196659
               OrigLeft        =   3510
               OrigTop         =   2440
               OrigRight       =   3765
               OrigBottom      =   2725
               Max             =   30
               SyncBuddy       =   -1  'True
               BuddyProperty   =   65547
               Enabled         =   -1  'True
            End
            Begin VB.Label lblDescrip 
               AutoSize        =   -1  'True
               Caption         =   "Días para Cálculo de la Fecha de Pago :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   35
               Left            =   360
               TabIndex        =   103
               Top             =   2160
               Width           =   2895
            End
            Begin VB.Label lblDescrip 
               AutoSize        =   -1  'True
               Caption         =   "Tipo Día"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   34
               Left            =   360
               TabIndex        =   102
               Top             =   2880
               Width           =   630
            End
            Begin VB.Label lblDescrip 
               AutoSize        =   -1  'True
               Caption         =   "Num. Días"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   20
               Left            =   360
               TabIndex        =   101
               Top             =   2520
               Width           =   765
            End
         End
         Begin MSComCtl2.DTPicker dtpFechaVencimiento 
            Height          =   285
            Left            =   7200
            TabIndex        =   46
            Top             =   720
            Width           =   2175
            _ExtentX        =   3836
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
            Format          =   175570945
            CurrentDate     =   38768
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Forma de Cálculo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   37
            Left            =   360
            TabIndex        =   111
            Top             =   4440
            Width           =   1230
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Cupón"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   36
            Left            =   360
            TabIndex        =   110
            Top             =   1120
            Width           =   1050
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Ajuste"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   13
            Left            =   360
            TabIndex        =   106
            Top             =   1815
            Width           =   795
         End
         Begin VB.Label lblNumCupones 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7200
            TabIndex        =   47
            Top             =   1100
            Width           =   2175
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Cupones"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   22
            Left            =   5040
            TabIndex        =   100
            Top             =   1125
            Width           =   855
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Frecuencia Pago"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   33
            Left            =   360
            TabIndex        =   99
            Top             =   4080
            Width           =   1215
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Base de Cálculo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   32
            Left            =   360
            TabIndex        =   98
            Top             =   3735
            Width           =   1155
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Tasa"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   31
            Left            =   360
            TabIndex        =   97
            Top             =   3405
            Width           =   945
         End
         Begin VB.Label lblPorc 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   4095
            TabIndex        =   96
            Top             =   2940
            Width           =   255
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tasa Anual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   21
            Left            =   360
            TabIndex        =   95
            Top             =   2955
            Width           =   810
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fin Indice"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   15
            Left            =   360
            TabIndex        =   94
            Top             =   2535
            Width           =   690
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Inicio Indice"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   14
            Left            =   360
            TabIndex        =   93
            Top             =   2175
            Width           =   855
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Vencimiento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   5040
            TabIndex        =   92
            Top             =   720
            Width           =   1365
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Emisión"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   19
            Left            =   5040
            TabIndex        =   91
            Top             =   375
            Width           =   1035
         End
      End
      Begin TAMControls.ucBotonEdicion cmdAccion 
         Height          =   390
         Left            =   6540
         TabIndex        =   114
         Top             =   5700
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   688
         Buttons         =   2
         Caption0        =   "&Guardar"
         Tag0            =   "2"
         Visible0        =   0   'False
         ToolTipText0    =   "Guardar"
         Caption1        =   "&Cancelar"
         Tag1            =   "8"
         Visible1        =   0   'False
         ToolTipText1    =   "Cancelar"
         UserControlHeight=   390
         UserControlWidth=   2700
      End
   End
End
Attribute VB_Name = "frmTitulosInversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Mantenimiento de Títulos de Renta Fija"
Option Explicit

Dim arrTipoCriterio()       As String, arrClaseCriterio()           As String
Dim arrSubClaseCriterio()   As String, arrTipoInstrumento()         As String
Dim arrClaseInstrumento()   As String, arrSubClaseInstrumento()     As String
Dim arrMonedaEmision()      As String, arrMonedaPago()              As String
Dim arrEmisor()             As String, arrBaseCalculo()             As String
Dim arrPeriodoPago()        As String, arrTipoTasa()                As String
Dim arrTasa()               As String, arrTipoDia()                 As String
Dim arrTipoAjuste()         As String, arrTipoVac()                 As String
Dim arrCuponCalculo()       As String, arrFormaCalculo()            As String
Dim arrMercado()            As String

Dim strCodTipoCriterio      As String, strCodClaseCriterio          As String
Dim strCodSubClaseCriterio  As String, strCodTipoInstrumento        As String
Dim strCodClaseInstrumento  As String, strCodSubClaseInstrumento    As String
Dim strCodMonedaEmision     As String, strCodMonedaPago             As String
Dim strCodEmisor            As String, strCodBaseCalculo            As String
Dim strCodPeriodoPago       As String, strCodTipoTasa               As String
Dim strCodTasa              As String, strCodTipoDia                As String
Dim strCodGrupo             As String, strCodCiiu                   As String
Dim strCodRiesgo            As String, strCodSubRiesgo              As String
Dim strCodCategoria         As String, strCodTipoAjuste             As String
Dim strCodTipoVac           As String, strCodCuponCalculo           As String
Dim strIndGenerado          As String, strCodSector                 As String
Dim strIndTasaAjustada      As String, strIndAmortizacion           As String
Dim strIndQuiebre           As String, strIndTasaCapitalizable      As String
Dim strIndCuponCero         As String, strCodValor                  As String
Dim strCodTipoPlazo         As String, strCodFormaCalculo           As String
Dim strEstado               As String, strSQL                       As String
Dim strCodMercado           As String

'*** Para el procedimiento de guardar en tabla temporal ***
Dim strNumCupon             As String, strFechaInicio               As String
Dim strFechaFin             As String, strCodAnalitica              As String
Dim strFechaInicioIndice    As String, strFechaFinIndice            As String
Dim strFechaPago            As String
Dim strIndVencido           As String, strIndVigente                As String
Dim intCantDias             As Integer
Dim curValorCupon           As Double
Dim dblSaldoAmortizacion    As Double, curValorAmortizacion         As Double
Dim dblValorInteres         As Double, dblAcumuladoAmortizacion     As Double
Dim dblFactorDiario         As Double, dblTasaAnualCupon            As Double
Dim dblFactorAnualCupon     As Double, dblTasaCuponNormal           As Double
Dim dblFactorDiarioNormal   As Double, dblPorcenAmortizacion        As Double

Dim adoConsulta             As ADODB.Recordset
Dim indSortAsc              As Boolean, indSortDesc                 As Boolean

Private Sub CargarTemporal()

    With adoComm
        .CommandText = "DELETE InstrumentoInversionCalendarioTmp " & _
            "WHERE CodTitulo='" & Trim(txtCodigoValor.Text) & "'"
        adoConn.Execute .CommandText
        
        .CommandText = "INSERT INTO InstrumentoInversionCalendarioTmp SELECT * FROM InstrumentoInversionCalendario " & _
            "WHERE CodTitulo='" & Trim(txtCodigoValor.Text) & "'"
        adoConn.Execute .CommandText
    End With
    
End Sub

Private Sub Deshabilita()

    optTipoCodigo(0).Enabled = False
    optTipoCodigo(1).Enabled = False
    txtCodigoValor.Enabled = False
    txtNemonico.Enabled = False
    cboTipoInstrumento.Enabled = False

End Sub

Private Sub GrabarFechaCorteTmp()

    Dim intRegistro As Integer
    
    With adoComm
        .CommandText = "UPDATE InstrumentoInversionCalendarioTmp SET FechaInicio='" & strFechaInicio & "'," & _
            "FechaVencimiento='" & strFechaFin & "',CantDiasPeriodo=" & intCantDias & " " & _
            "WHERE CodTitulo='" & Trim(txtCodigoValor.Text) & "' AND NumCupon='" & strNumCupon & "'"
            
        adoConn.Execute .CommandText, intRegistro
        
        If intRegistro = 0 Then
            .CommandText = "INSERT INTO InstrumentoInversionCalendarioTmp " & _
                "(CodTitulo,NumCupon,CodFile,NumSecuencial,FechaInicio,FechaVencimiento,CantDiasPeriodo) VALUES('" & _
                Trim(txtCodigoValor.Text) & "','" & strNumCupon & "','" & strCodTipoInstrumento & "'," & _
                CInt(strNumCupon) & ",'" & strFechaInicio & "','" & strFechaFin & "'," & intCantDias & ")"

            adoConn.Execute .CommandText
        End If
        
    End With
    
End Sub

Private Sub GrabarFechaIndiceTmp()

    Dim intRegistro As Integer
    
    With adoComm
        .CommandText = "UPDATE InstrumentoInversionCalendarioTmp SET " & _
            "FechaInicioIndice='" & strFechaInicioIndice & "'," & _
            "FechaFinIndice='" & strFechaFinIndice & "'," & _
            "IndVencido='" & strIndVencido & "',IndVigente='" & strIndVigente & "' " & _
            "WHERE CodTitulo='" & Trim(txtCodigoValor.Text) & "' AND NumCupon='" & strNumCupon & "'"
        adoConn.Execute .CommandText, intRegistro
        
        If intRegistro = 0 Then
            .CommandText = "INSERT INTO InstrumentoInversionCalendarioTmp " & _
                "(CodTitulo,NumCupon,CodFile,NumSecuencial,FechaInicioIndice,FechaFinIndice,IndVencido,IndVigente) VALUES('" & _
                Trim(txtCodigoValor.Text) & "','" & strNumCupon & "','" & strCodTipoInstrumento & "'," & _
                CInt(strNumCupon) & ",'" & strFechaInicioIndice & "','" & strFechaFinIndice & "','" & strIndVencido & "','" & strIndVigente & "')"
            adoConn.Execute .CommandText
        End If
        
    End With
    
End Sub

Private Sub GrabarFechaPagoTmp()

    Dim intRegistro As Integer
    
    With adoComm
        .CommandText = "UPDATE InstrumentoInversionCalendarioTmp SET FechaPago='" & strFechaPago & "'," & _
            "IndVencido='" & strIndVencido & "',PorcenAmortizacion=" & dblPorcenAmortizacion & ",IndVigente='" & strIndVigente & "' " & _
            "WHERE CodTitulo='" & Trim(txtCodigoValor.Text) & "' AND NumCupon='" & strNumCupon & "'"
        adoConn.Execute .CommandText, intRegistro
        
        If intRegistro = 0 Then
            .CommandText = "INSERT INTO InstrumentoInversionCalendarioTmp " & _
                "(CodTitulo,NumCupon,CodFile,NumSecuencial,FechaPago,IndVencido,IndVigente) VALUES('" & _
                Trim(txtCodigoValor.Text) & "','" & strNumCupon & "','" & strCodTipoInstrumento & "'," & _
                CInt(strNumCupon) & ",'" & strFechaPago & "','" & strIndVencido & "','" & strIndVigente & "')"
            adoConn.Execute .CommandText
        End If
        
    End With
    
End Sub

Private Sub GrabarFactoresMontosTmp()

    Dim intRegistro As Integer
    
    With adoComm
        .CommandText = "UPDATE InstrumentoInversionCalendarioTmp SET TasaInteres=" & CDec(dblTasaAnualCupon) & "," & _
            "FactorInteres=" & CDec(dblFactorAnualCupon) & ",FactorDiario=" & CDec(dblFactorDiario) & ",ValorInteres=" & CDec(dblValorInteres) & "," & _
            "ValorAmortizacion=" & CDec(curValorAmortizacion) & ",SaldoAmortizacion=" & CDec(dblSaldoAmortizacion) & ",AcumuladoAmortizacion=" & CDec(dblAcumuladoAmortizacion) & "," & _
            "ValorCupon=" & CDec(curValorCupon) & ",FactorInteres1=" & CDec(dblTasaCuponNormal) & ",FactorDiario1=" & CDec(dblFactorDiarioNormal) & " " & _
            "WHERE CodTitulo='" & Trim(txtCodigoValor.Text) & "' AND NumCupon='" & strNumCupon & "'"
        adoConn.Execute .CommandText, intRegistro
        
        If intRegistro = 0 Then
            .CommandText = "INSERT INTO InstrumentoInversionCalendarioTmp " & _
                "(CodTitulo,NumCupon,CodFile,NumSecuencial,TasaInteres,FactorInteres,FactorDiario,ValorInteres,ValorAmortizacion,SaldoAmortizacion,AcumuladoAmortizacion,ValorCupon,FactorInteres1) VALUES('" & _
                Trim(txtCodigoValor.Text) & "','" & strNumCupon & "','" & strCodTipoInstrumento & "'," & CInt(strNumCupon) & "," & CDec(dblTasaAnualCupon) & "," & _
                CDec(dblFactorAnualCupon) & "," & CDec(dblFactorDiario) & "," & CDec(dblValorInteres) & "," & CDec(curValorAmortizacion) & "," & CDec(dblSaldoAmortizacion) & "," & _
                CDec(dblAcumuladoAmortizacion) & "," & CDec(curValorCupon) & "," & CDec(dblTasaCuponNormal) & "," & CDec(dblFactorDiarioNormal) & ")"
            adoConn.Execute .CommandText
        End If
        
    End With
    
End Sub
Private Sub Habilita()

    optTipoCodigo(0).Enabled = True
    optTipoCodigo(1).Enabled = True
    txtCodigoValor.Enabled = True
    txtNemonico.Enabled = True
    cboTipoInstrumento.Enabled = True
    
End Sub

Private Sub LlenarFormulario(strModo As String)

    Dim adoRegistro As ADODB.Recordset
    Dim intRegistro As Integer
    
    Select Case strModo
        Case Reg_Adicion
            optTipoCodigo(0).Value = True
            
            strCodValor = Valor_Caracter: lblAnalitica = "???-????????"
            txtCodigoValor.Text = Valor_Caracter: txtNemonico.Text = Valor_Caracter
            txtDescripValor.Text = Valor_Caracter
            txtNumEmision.Text = Valor_Caracter: txtNumSerie.Text = Valor_Caracter
            
            cboTipoInstrumento.ListIndex = -1
            If cboTipoInstrumento.ListCount > 0 Then cboTipoInstrumento.ListIndex = 0
            
            cboEmisor.ListIndex = -1
            If cboEmisor.ListCount > 0 Then cboEmisor.ListIndex = 0
            
            lblGrupoEconomico.Caption = Valor_Caracter
            lblCiiu.Caption = Valor_Caracter
            
            cboMercado.ListIndex = -1
            If cboMercado.ListCount > 0 Then cboMercado.ListIndex = 0
            
            cboMonedaEmision.ListIndex = -1
            If cboMonedaEmision.ListCount > 0 Then cboMonedaEmision.ListIndex = 0
            
            cboMonedaPago.ListIndex = -1
            If cboMonedaPago.ListCount > 0 Then cboMonedaPago.ListIndex = 0
            
            lblRiesgo.Caption = Valor_Caracter
            strCodRiesgo = Valor_Caracter: strCodSubRiesgo = Valor_Caracter
            
            chkQuiebre.Value = vbUnchecked: chkQuiebre.Enabled = False
            
            dtpFechaEmision.Value = gdatFechaActual
            dtpFechaVencimiento.Value = gdatFechaActual

            txtValorNominal.Text = "0"
            
            chkCuponCero.Value = vbUnchecked
            chkAmortiza.Value = vbUnchecked
            chkAjuste.Value = vbUnchecked
            
            intRegistro = ObtenerItemLista(arrTipoAjuste(), Codigo_Tipo_Ajuste_Vac)
            If intRegistro >= 0 Then cboTipoAjuste.ListIndex = intRegistro
            
            cboTipoAjuste.Enabled = False
            cboTipoVac.Enabled = False
            cboCuponCalculo.Enabled = False
            
            txtTasaAnual.Text = "0"
            
            cboTipoTasa.ListIndex = -1
            If cboTipoTasa.ListCount > 0 Then cboTipoTasa.ListIndex = 0
            
            cboTasa.ListIndex = -1
            If cboTasa.ListCount > 0 Then cboTasa.ListIndex = 0
            
            chkCapitalizable.Value = vbUnchecked: chkCapitalizable.Enabled = False
            
            cboBaseCalculo.ListIndex = -1
            If cboBaseCalculo.ListCount > 0 Then cboBaseCalculo.ListIndex = 0
            
            cboPeriodoPago.ListIndex = -1
            If cboPeriodoPago.ListCount > 0 Then cboPeriodoPago.ListIndex = 0
            
            cboFormaCalculo.ListIndex = -1
            If cboFormaCalculo.ListCount > 0 Then cboFormaCalculo.ListIndex = 0
            
            lblNumCupones.Caption = "0"
            
            optDiasCupon(0).Value = True
            txtNumDiasPeriodo.Text = "0": txtNumDiasPeriodo.Visible = False
            updDiasPeriodo.Visible = False
            
            chkAPartir.Value = vbUnchecked
            dtpFechaInicioCupon.Value = dtpFechaEmision.Value
            dtpFechaInicioCupon.Visible = False
            
            chk1erCupon.Value = vbUnchecked
            dtpFechaCorteInicial.Value = dtpFechaInicioCupon.Value
            dtpFechaCorteInicial.Visible = False
            
            txtNumDiasPago.Text = "0"
            cboTipoDia.ListIndex = -1
            If cboTipoDia.ListCount > 0 Then cboTipoDia.ListIndex = 0
            
            cmdGeneraCuponera(1).Enabled = False: cmdGeneraCuponera(2).Enabled = False
            cmdGeneraCuponera(3).Enabled = False: cmdGeneraCuponera(0).Enabled = True
            
            chkCuponCero.Value = vbChecked
            txtCodigoValor.Enabled = True
            cboTipoInstrumento.Enabled = True
            cmdRegenerar.Visible = False
            cboTipoInstrumento.SetFocus
                        
        Case Reg_Edicion
            Set adoRegistro = New ADODB.Recordset
            
            strCodValor = Valor_Caracter
            strCodValor = Trim(tdgConsulta.Columns("CodTitulo").Value)

            adoComm.CommandText = "SELECT * FROM InstrumentoInversion WHERE CodTitulo='" & strCodValor & "'"
            Set adoRegistro = adoComm.Execute

            If Not adoRegistro.EOF Then
                optTipoCodigo(0).Value = True
                If Trim(adoRegistro("IndGenerado")) = Valor_Indicador Then optTipoCodigo(1).Value = True
                
                txtCodigoValor.Text = Trim(adoRegistro("CodTitulo"))
                txtNemonico.Text = Trim(adoRegistro("Nemotecnico"))
                txtDescripValor.Text = Trim(adoRegistro("DescripTitulo"))
                txtNumEmision.Text = Trim(adoRegistro("NumEmision"))
                txtNumSerie.Text = Trim(adoRegistro("NumSerie"))
                strCodAnalitica = Trim(adoRegistro("CodAnalitica"))

                intRegistro = ObtenerItemLista(arrTipoInstrumento(), adoRegistro("CodFile"))
                If intRegistro >= 0 Then cboTipoInstrumento.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrClaseInstrumento(), adoRegistro("CodDetalleFile"))
                If intRegistro >= 0 Then cboClaseInstrumento.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrSubClaseInstrumento(), adoRegistro("CodSubDetalleFile"))
                If intRegistro >= 0 Then cboSubClaseInstrumento.ListIndex = intRegistro

                'intRegistro = ObtenerItemLista(arrEmisor(), adoRegistro("CodEmisor") + adoRegistro("CodGrupo") + adoRegistro("CodCiiu"))
                intRegistro = ObtenerItemLista(arrEmisor(), adoRegistro("CodEmisor"))
                If intRegistro >= 0 Then cboEmisor.ListIndex = intRegistro
                
               
                intRegistro = ObtenerItemLista(arrMercado(), adoRegistro("CodMercado"))
                If intRegistro >= 0 Then cboMercado.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrMonedaEmision(), adoRegistro("CodMoneda"))
                If intRegistro >= 0 Then cboMonedaEmision.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrMonedaPago(), adoRegistro("CodMoneda1"))
                If intRegistro >= 0 Then cboMonedaPago.ListIndex = intRegistro

                strCodRiesgo = adoRegistro("CodRiesgo")
                strCodSubRiesgo = adoRegistro("CodSubRiesgo")
                
                dtpFechaEmision.Value = adoRegistro("FechaEmision")
                dtpFechaVencimiento.Value = adoRegistro("FechaVencimiento")
                txtValorNominal.Text = adoRegistro("ValorNominal")
                
                intRegistro = ObtenerItemLista(arrTipoAjuste(), Codigo_Tipo_Ajuste_Vac)
                If intRegistro >= 0 Then cboTipoAjuste.ListIndex = intRegistro
                cboTipoAjuste.Enabled = False
                cboTipoVac.Enabled = False
                cboCuponCalculo.Enabled = False
            
                If Trim(adoRegistro("IndCuponCero")) = Valor_Indicador Then
                    chkCuponCero.Value = vbChecked
                    txtTasaAnual.Text = CStr(adoRegistro("TasaInteres"))
                Else
                    chkCuponCero.Value = vbUnchecked
                    
                    intRegistro = ObtenerItemLista(arrBaseCalculo(), adoRegistro("BaseAnual"))
                    If intRegistro >= 0 Then cboBaseCalculo.ListIndex = intRegistro
                    
                    intRegistro = ObtenerItemLista(arrFormaCalculo(), adoRegistro("FormaCalculo"))
                    If intRegistro >= 0 Then cboFormaCalculo.ListIndex = intRegistro
                    
                    intRegistro = ObtenerItemLista(arrPeriodoPago(), adoRegistro("PeriodoPago"))
                    If intRegistro >= 0 Then cboPeriodoPago.ListIndex = intRegistro
                    
                    lblNumCupones.Caption = CStr(adoRegistro("CantCupones"))
                    chkAmortiza.Value = vbUnchecked
                    If Trim(adoRegistro("IndAmortizacion")) = Valor_Indicador Then chkAmortiza.Value = vbChecked
                    chkAjuste.Value = vbUnchecked
                    If Trim(adoRegistro("IndTasaAjustada")) = Valor_Indicador Then chkAjuste.Value = vbChecked
                    
                    intRegistro = ObtenerItemLista(arrTipoVac(), adoRegistro("CodTipoVac"))
                    If intRegistro >= 0 Then cboTipoVac.ListIndex = intRegistro
                    
                    If cboTipoVac.ListIndex > 0 Then
                        intRegistro = ObtenerItemLista(arrCuponCalculo(), adoRegistro("CuponCalculo"))
                        If intRegistro >= 0 Then cboCuponCalculo.ListIndex = intRegistro
                    End If
                    
                    txtTasaAnual.Text = CStr(adoRegistro("TasaInteres"))
                    
                    intRegistro = ObtenerItemLista(arrTipoTasa(), adoRegistro("CodTipoTasa"))
                    If intRegistro >= 0 Then cboTipoTasa.ListIndex = intRegistro
                    
                    intRegistro = ObtenerItemLista(arrTasa(), adoRegistro("CodTasa"))
                    If intRegistro >= 0 Then cboTasa.ListIndex = intRegistro
                    
                    chkCapitalizable.Value = vbUnchecked
                    If Trim(adoRegistro("IndTasaCapitalizable")) = Valor_Indicador Then chkCapitalizable.Value = vbChecked
                                                            
                End If
                
                optDiasCupon(0).Value = True
                txtNumDiasPeriodo.Text = "0": txtNumDiasPeriodo.Visible = False
                updDiasPeriodo.Visible = False
            
                chkAPartir.Value = vbUnchecked
                dtpFechaInicioCupon.Value = dtpFechaEmision.Value
                dtpFechaInicioCupon.Visible = False
                
                chk1erCupon.Value = vbUnchecked
                dtpFechaCorteInicial.Value = dtpFechaInicioCupon.Value
                dtpFechaCorteInicial.Visible = False
                
                txtNumDiasPago.Text = "0"
                cboTipoDia.ListIndex = -1
                If cboTipoDia.ListCount > 0 Then cboTipoDia.ListIndex = 0
                
                lblAnalitica.Caption = Trim(adoRegistro("CodFile")) & "-" & Trim(adoRegistro("CodAnalitica"))
                
                cmdGeneraCuponera(0).Enabled = False: cmdGeneraCuponera(1).Enabled = False
                cmdGeneraCuponera(2).Enabled = False: cmdGeneraCuponera(3).Enabled = True
            
                Call CargarTemporal
            End If
            adoRegistro.Close: Set adoRegistro = Nothing
            cmdRegenerar.Visible = True
    End Select
    
End Sub
Public Sub Adicionar()

    frmMainMdi.stbMdi.Panels(3).Text = "Adicionar Título Valor..."
                
    strEstado = Reg_Adicion
    LlenarFormulario strEstado
    cmdOpcion.Visible = False
    With tabTitulos
        .TabEnabled(0) = False
        .TabEnabled(1) = True
        .TabEnabled(2) = True
        .Tab = 1
    End With
    Call Habilita
    
End Sub

Public Sub Cancelar()

    cmdOpcion.Visible = True
    With tabTitulos
        .TabEnabled(0) = True
        .Tab = 0
        .TabEnabled(1) = False
        .TabEnabled(2) = False
    End With
    Call Buscar
    
End Sub

Public Sub Eliminar()

End Sub

Public Sub GenerarFactoresMontos(strpNumCupon As String, dblpTasaModificar As Double)

    Dim adoConsulta             As ADODB.Recordset
    Dim adoRegistro             As ADODB.Recordset
    Dim datFechaPago            As Date, datFechaInicio         As Date
    Dim datFechaCorte           As Date, datFechaFinAnterior    As Date
    Dim datFechaInicioAnterior  As Date
    Dim intContador             As Integer, intNumDiasTmp       As Integer
    Dim intNumDiasPago          As Integer, intCuponVigente     As Integer
    Dim curValorNominal         As Currency, intNumDiasPeriodo  As Integer
    Dim intNumMesesPeriodo      As Integer, intNumPeriodosAnual As Integer
    Dim intDiasPeriodo          As Integer, intCantDiasPeriodo  As Integer
    Dim intIndCorrecto          As Integer, intCantDiasCupon    As Integer
    Dim intNumDiasAnterior      As Integer, intBaseVac          As Integer
'    Dim lngDiasAnual            As Long, lngDiasMensual         As Long
'    Dim lngDiasDiario           As Long+
    Dim intFrecuenciaPago       As Integer
    Dim dblTasaCupon            As Double, dblTasaInteres       As Double
    Dim dblFactor1              As Double, dblFactor2           As Double
    Dim dblFactorInicioVac      As Double, dblFactorFinVac      As Double
    Dim dblTasaAnual            As Double, dblNumCupones        As Double
    
    '*** Cronograma a partir de la personalización de los días del mes ***
    If optDiasCupon(1).Value Then
        intNumDiasPeriodo = CInt(txtNumDiasPeriodo.Text) * intNumMesesPeriodo '*** Dias del periodo ***
    End If
        
    '*** Base de Cálculo ***
    intDiasPeriodo = 365
    Select Case strCodBaseCalculo
        Case Codigo_Base_30_360: intDiasPeriodo = 360
        Case Codigo_Base_Actual_365: intDiasPeriodo = 365
        Case Codigo_Base_Actual_360: intDiasPeriodo = 360
        Case Codigo_Base_30_365: intDiasPeriodo = 365
    End Select
    
    With adoComm
        Set adoConsulta = New ADODB.Recordset

        '*** Obtener el número de días del periodo de pago ***
        .CommandText = "SELECT ValorParametro FROM AuxiliarParametro WHERE CodTipoParametro='TIPFRE' AND CodParametro='" & strCodPeriodoPago & "'"
        Set adoConsulta = .Execute

        If Not adoConsulta.EOF Then
            intNumDiasPeriodo = CInt(adoConsulta("Valorparametro")) '*** Días del periodo  ***
            intNumMesesPeriodo = CInt(intNumDiasPeriodo / 30)       '*** Meses del periodo ***
            If intNumMesesPeriodo = 0 Then intNumMesesPeriodo = 1
            intNumPeriodosAnual = CInt(12 / intNumMesesPeriodo)     '*** Periodos al año   ***
            intCantDiasCupon = intNumDiasPeriodo
        End If
        adoConsulta.Close: Set adoConsulta = Nothing
    End With
    
     '*** Valor Nominal del Título ***
    curValorNominal = CCur(txtValorNominal.Text)
    dblAcumuladoAmortizacion = 0

    If chkAmortiza.Value = vbUnchecked Then
       dblTasaAnual = CDbl(txtTasaAnual.Text) '*** Tasa de interés ***
       If dblpTasaModificar > 0 Then dblTasaAnual = dblpTasaModificar
       curValorCupon = curValorNominal * (dblTasaAnual * 0.01) * (intNumDiasPeriodo / intDiasPeriodo) '*** Cálculo del valor del cupón ***
       dblSaldoAmortizacion = curValorNominal '*** Saldo por Amortizar por defecto ***
    End If
    
    dblTasaCupon = CDbl(txtTasaAnual.Text)
    If dblpTasaModificar > 0 Then dblTasaCupon = dblpTasaModificar

    '*** Amortización ?? ***
    If chkAmortiza.Value Then
        If strCodTipoTasa = Codigo_Tipo_Tasa_Efectiva Then '*** Efectiva ***
            dblTasaAnual = ((1 + (CDbl(txtTasaAnual.Text) * 0.01)) ^ (intNumDiasPeriodo / intDiasPeriodo)) - 1
            If dblpTasaModificar > 0 Then
                dblTasaAnual = ((1 + (dblpTasaModificar * 0.01)) ^ (intNumDiasPeriodo / intDiasPeriodo)) - 1
            End If
            dblFactorDiario = ((1 + dblTasaAnual) ^ (1 / intNumDiasPeriodo)) - 1
        ElseIf chkCapitalizable.Value Then '*** Nominal Capitalizable ***
            If strCodFormaCalculo = Codigo_Calculo_Prorrateo Then
                dblTasaAnual = (CDbl(txtTasaAnual.Text) * 0.01) * intNumDiasPeriodo / intDiasPeriodo
                dblTasaCuponNormal = (CDbl(txtTasaAnual.Text) * 0.01) * intNumDiasPeriodo / intDiasPeriodo
            Else
                dblFactorAnualCupon = (CDbl(txtTasaAnual.Text) * 0.01) / (intDiasPeriodo / intNumDiasPeriodo)
                dblTasaCuponNormal = (CDbl(txtTasaAnual.Text) * 0.01) / (intDiasPeriodo / intNumDiasPeriodo)
            End If
'            dblTasaAnual = (CDbl(txtTasaAnual.Text) * 0.01) * (intNumDiasPeriodo / intDiasPeriodo)
'            dblTasaCuponNormal = (CDbl(txtTasaAnual.Text) * 0.01) * intNumDiasPeriodo / intDiasPeriodo
            If dblpTasaModificar > 0 Then
                dblTasaAnual = (dblpTasaModificar * 0.01) * (intNumDiasPeriodo / intDiasPeriodo)
                dblTasaCuponNormal = (dblpTasaModificar * 0.01) * intNumDiasPeriodo / intDiasPeriodo
            End If
            dblFactorDiario = dblTasaAnual / intNumDiasPeriodo
            dblFactorDiarioNormal = dblTasaCuponNormal / intNumDiasPeriodo
        ElseIf Not chkCapitalizable.Value Then '*** Nominal No Capitalizable ***
            dblTasaAnual = (CDbl(txtTasaAnual.Text) * 0.01) / intNumPeriodosAnual
            If dblpTasaModificar > 0 Then
                dblTasaAnual = (dblpTasaModificar * 0.01) / intNumPeriodosAnual
            End If
            dblFactorDiario = dblTasaAnual / intNumDiasPeriodo
        End If

        '*** Cálculo del valor del cupón ***
        If (1 - (1 / (1 + dblTasaAnual))) <> 0 Then
        
            If CInt(lblNumCupones) = 0 Then
                dblNumCupones = CInt(lblNumCupones)
            Else
                dblNumCupones = 1
            End If
           curValorCupon = (curValorNominal * dblTasaAnual) / (1 - (1 / (1 + dblTasaAnual)) ^ dblNumCupones)
        Else
           curValorCupon = 0
        End If

        dblSaldoAmortizacion = curValorNominal '*** Saldo a amortizar inicial = Valor nominal del titulo ***
    End If
        
    Set adoConsulta = New ADODB.Recordset
    intCuponVigente = 0
    
    With adoComm
        If dblpTasaModificar > 0 Then
            .CommandText = "SELECT CodTitulo,NumCupon,FechaInicio,FechaVencimiento,CantDiasPeriodo,PorcenAmortizacion " & _
                "FROM InstrumentoInversionCalendarioTmp " & _
                "WHERE CodTitulo='" & Trim(txtCodigoValor.Text) & "' AND NumCupon='" & strpNumCupon & "' ORDER BY NumCupon"
        Else
            .CommandText = "SELECT CodTitulo,NumCupon,FechaInicio,FechaVencimiento,CantDiasPeriodo,PorcenAmortizacion " & _
                "FROM InstrumentoInversionCalendarioTmp " & _
                "WHERE CodTitulo='" & Trim(txtCodigoValor.Text) & "' ORDER BY NumCupon"
        End If
        Set adoConsulta = .Execute
        
        Do While Not adoConsulta.EOF
        
            If chkAmortiza.Value Then
               dblTasaInteres = (((dblTasaAnual + 1) ^ (intDiasPeriodo / CInt(adoConsulta("CantDiasPeriodo")))) - 1) * 100
               dblTasaInteres = dblTasaCupon
            Else
               dblTasaInteres = dblTasaCupon
            End If

            '*** Cálculos principales ***
            datFechaInicio = CVDate(adoConsulta("FechaInicio"))     'fecha de inicio de cupon
            datFechaCorte = CVDate(adoConsulta("FechaVencimiento")) 'fecha de corte de cupon
            
            If strCodFormaCalculo = Codigo_Calculo_Prorrateo Then
'                lngDiasAnual = (CLng(Year(datFechaCorte)) - CLng(Year(datFechaInicio))) * 360
'                lngDiasMensual = (CLng(Month(datFechaCorte)) - CLng(Month(datFechaInicio))) * 30
'                lngDiasDiario = (CLng(Day(datFechaCorte)) - CLng(Day(datFechaInicio))) + 1
'                intCantDiasPeriodo = CInt(lngDiasAnual + lngDiasMensual + lngDiasDiario)
                If DateDiff("d", dtpFechaEmision.Value, datFechaInicio) > 0 And DateDiff("d", dtpFechaEmision.Value, datFechaInicio) <= 2 Then
                    intCantDiasPeriodo = Dias360(dtpFechaEmision.Value, datFechaCorte, True) - 1
                Else
                    intCantDiasPeriodo = Dias360(datFechaInicio, datFechaCorte, True) - 1
                End If
            Else
                intCantDiasPeriodo = CInt(adoConsulta("CantDiasPeriodo"))
            
'                Select Case strCodBaseCalculo
'                    Case Codigo_Base_30_360: intCantDiasPeriodo = intNumDiasPeriodo
'                    Case Codigo_Base_Actual_365: intCantDiasPeriodo = CInt(adoConsulta("CantDiasPeriodo"))
'                    Case Codigo_Base_Actual_360: intCantDiasPeriodo = CInt(adoConsulta("CantDiasPeriodo"))
'                    Case Codigo_Base_30_365: intCantDiasPeriodo = intNumDiasPeriodo
'                End Select
                
            End If
            
            dblFactor1 = dblTasaInteres
            dblFactor2 = dblTasaInteres

'            If strCodTipoVac = Codigo_Vac_Emision Then  '*** VAC periodico ***
'                If dblFactor1 = 0 Then
'                    dblFactor2 = 0
'                Else
'                    dblFactor2 = CDbl(txtTasaAnual.Text)
'                    If dblpTasaModificar > 0 Then dblFactor2 = dblpTasaModificar
'                End If
'            End If

            intIndCorrecto = 1

            '*** Número de días a considerar según la periodicidad de pago ***
            Select Case strCodPeriodoPago
            
                Case "01" '*** Anual ***
                    If (intCantDiasPeriodo >= 360) And (intCantDiasPeriodo <= 366) Then intIndCorrecto = 0
                    
                Case "02" '*** Semestral ***
                    If (intCantDiasPeriodo >= 180) And (intCantDiasPeriodo <= 183) Then intIndCorrecto = 0
                    
                Case "03"  '*** Trimestral ***
                    If (intCantDiasPeriodo >= 89) And (intCantDiasPeriodo <= 92) Then intIndCorrecto = 0
                    
                Case "04" '*** Bimestral ***
                    If (intCantDiasPeriodo >= 60) And (intCantDiasPeriodo <= 61) Then intIndCorrecto = 0
                            
                Case "05" '*** Mensual ***
                    If (intCantDiasPeriodo >= 28) And (intCantDiasPeriodo <= 31) Then intIndCorrecto = 0
                    
                Case "06" '*** Quincenal ***
                    If intCantDiasPeriodo = 15 Then intIndCorrecto = 0
                    
                Case "07" '*** Diaria ***
                    If intCantDiasPeriodo = 1 Then intIndCorrecto = 0
                    
            End Select

            '*** Para VAC ***
            intBaseVac = 365

            '*** Si es VAC ***
'            If chkAjuste.Value Then
'                If strCodTipoVac = Codigo_Vac_Emision Then '*** VAC periodico ***
'                    If strCodCuponCalculo = Codigo_Vac_Liquidacion Then '*** A partir del cupón anterior ***
'                        datFechaFinAnterior = DateAdd("d", -1, datFechaInicio)
'                        datFechaInicioAnterior = DateAdd("m", Int(intCantDiasCupon / 30) * -1, datFechaFinAnterior) + 1
'                    Else '*** A partir del cupón vigente ***
'                        datFechaFinAnterior = datFechaCorte
'                        datFechaInicioAnterior = datFechaInicio
'                    End If
'
'                    Set adoRegistro = New ADODB.Recordset
'
'                    '*** Obtener tasa VAC al inicio del cupón ***
'                    .CommandText = "SELECT ValorTasa FROM InversionTasa WHERE CodTasa='" & strCodTipoAjuste & "' AND " & _
'                        "(FechaRegistro>='" & Convertyyyymmdd(datFechaInicioAnterior) & "' AND " & _
'                        "FechaRegistro<'" & Convertyyyymmdd(DateAdd("d", 1, datFechaInicioAnterior)) & "')"
'                    Set adoRegistro = .Execute
'
'                    If adoRegistro.EOF Then
'                       dblFactorInicioVac = 0
'                       If gdatFechaActual > datFechaInicioAnterior Then
'                            MsgBox "No existe valor de ajuste para el día " & CStr(datFechaInicioAnterior), vbCritical, Me.Caption
'                       End If
'                    Else
'                       dblFactorInicioVac = adoRegistro("ValorTasa")
'                    End If
'                    adoRegistro.Close
'
'                    '*** Obtener tasa VAC al corte del cupón ***
'                    .CommandText = "SELECT ValorTasa FROM InversionTasa WHERE CodTasa='" & strCodTipoAjuste & "' AND " & _
'                        "(FechaRegistro>='" & Convertyyyymmdd(datFechaFinAnterior) & "' AND " & _
'                        "FechaRegistro<'" & Convertyyyymmdd(DateAdd("d", 1, datFechaFinAnterior)) & "')"
'                    Set adoRegistro = .Execute
'
'                    If adoRegistro.EOF Then
'                       dblFactorInicioVac = 0
'                       If gdatFechaActual > datFechaFinAnterior Then
'                            MsgBox "No existe valor de ajuste para el día " & CStr(datFechaFinAnterior), vbCritical, Me.Caption
'                       End If
'                    Else
'                       dblFactorFinVac = adoRegistro("ValorTasa")
'                    End If
'                    adoRegistro.Close: Set adoRegistro = Nothing
'
'
'                    If (dblFactorInicioVac > 0) And (dblFactorFinVac > 0) Then
'                       If strCodCuponCalculo = Codigo_Vac_Liquidacion Then '*** A partir del cupón anterior ***
'                          intNumDiasAnterior = DateDiff("d", datFechaInicioAnterior, datFechaFinAnterior) + 1
'                          dblTasaInteres = (((dblFactorFinVac / dblFactorInicioVac) ^ (intBaseVac / intNumDiasAnterior)) * (dblFactor1 / 100 + 1) - 1) * 100
'                       Else '*** A partir del cupón vigente ***
'                          dblTasaInteres = (((dblFactorFinVac / dblFactorInicioVac) ^ (intBaseVac / intCantDiasPeriodo)) * (dblFactor1 / 100 + 1) - 1) * 100
'                       End If
'                       dblFactor1 = dblTasaInteres
'                    Else '*** No hay tasas VAC al inicio y/o corte del cupón ***
'                       If CInt(adoConsulta("NumCupon")) > 1 Then
'                          If dblFactor1 = 0 Then
'                             dblTasaInteres = 0
'                          Else
'                             dblTasaInteres = 0 'aRCupBon(n_Sec - 1).TAE_INTE PENDIENTE
'                          End If
'                          dblFactor1 = dblTasaInteres
'                       End If
'                    End If
'                Else
'
'                    datFechaInicioAnterior = dtpFechaEmision
'
'                End If
'            End If

            dblTasaAnualCupon = dblFactor1
            
            '*** Calculando factores ***
            dblFactorAnualCupon = FactorAnual(dblFactor1, intNumDiasPeriodo, intDiasPeriodo, strCodTipoTasa, strIndTasaCapitalizable, strCodFormaCalculo, intIndCorrecto, intCantDiasCupon, intNumPeriodosAnual)
            dblTasaCuponNormal = FactorAnualNormal(dblFactor2, intNumDiasPeriodo, intDiasPeriodo, strCodTipoTasa, strIndTasaCapitalizable, strCodFormaCalculo, intIndCorrecto, intCantDiasCupon, intNumPeriodosAnual)
            dblFactorDiario = FactorDiario(dblFactorAnualCupon, intCantDiasPeriodo, strCodTipoTasa, strIndTasaCapitalizable, intCantDiasCupon)
            dblFactorDiarioNormal = FactorDiarioNormal(dblTasaCuponNormal, intCantDiasPeriodo, strCodTipoTasa, strIndTasaCapitalizable, CInt(adoConsulta("CantDiasPeriodo")))
            
'            If strCodTipoTasa = Codigo_Tipo_Tasa_Efectiva Then '*** Efectiva ***
'               dblFactorAnualCupon = ((1 + (0.01 * dblFactor1)) ^ (intCantDiasPeriodo / intDiasPeriodo)) - 1
'               dblTasaCuponNormal = ((1 + (0.01 * dblFactor2)) ^ (intCantDiasPeriodo / intDiasPeriodo)) - 1
'               dblFactorDiario = ((1 + dblFactorAnualCupon) ^ (1 / intCantDiasPeriodo)) - 1
'               dblFactorDiarioNormal = ((1 + dblTasaCuponNormal) ^ (1 / intCantDiasPeriodo)) - 1
'
'            ElseIf chkCapitalizable.Value Then '*** Nominal Capitalizable ***
'                If strCodFormaCalculo = Codigo_Calculo_Prorrateo Then
'                    dblFactorAnualCupon = (0.01 * dblFactor1) * intCantDiasPeriodo / intDiasPeriodo
'                    dblTasaCuponNormal = (0.01 * dblFactor2) * intCantDiasPeriodo / intDiasPeriodo
'                Else
'                    dblFactorAnualCupon = (0.01 * dblFactor1) / (intDiasPeriodo / intCantDiasPeriodo)
'                    dblTasaCuponNormal = (0.01 * dblFactor2) / (intDiasPeriodo / intCantDiasPeriodo)
'                End If
'
'               If intIndCorrecto = 1 Then
'                    If strCodFormaCalculo = Codigo_Calculo_Normal Then
'                        dblFactorAnualCupon = dblFactorAnualCupon / (intCantDiasCupon / intCantDiasPeriodo)
'                        dblTasaCuponNormal = dblTasaCuponNormal / (intCantDiasCupon / intCantDiasPeriodo)
'                    End If
'               End If
'               dblFactorDiario = dblFactorAnualCupon / intCantDiasPeriodo
'               dblFactorDiarioNormal = dblTasaCuponNormal / CInt(adoConsulta("CantDiasPeriodo"))
'
'            ElseIf Not chkCapitalizable.Value Then '*** Nominal No Capitalizable ***
'               dblFactorAnualCupon = (0.01 * dblFactor1) / intNumPeriodosAnual
'               dblTasaCuponNormal = (0.01 * dblFactor2) / intNumPeriodosAnual
'               If intIndCorrecto = 1 Then
'                    dblFactorAnualCupon = dblFactorAnualCupon / (intCantDiasCupon / intCantDiasPeriodo)
'                    dblTasaCuponNormal = dblTasaCuponNormal / (intCantDiasCupon / intCantDiasPeriodo)
'               End If
'               dblFactorDiario = dblFactorAnualCupon / intCantDiasPeriodo
'
'            End If
            '****************************
                    
            If chkAmortiza.Value = vbUnchecked Then
                If CInt(lblNumCupones.Caption) = CInt(adoConsulta("NumCupon")) Then
                    curValorAmortizacion = curValorNominal
                Else
                    curValorAmortizacion = 0
                End If
                dblValorInteres = Round(dblSaldoAmortizacion * dblFactorAnualCupon, 12)
                curValorCupon = Round(curValorAmortizacion + dblValorInteres, 12)
                dblSaldoAmortizacion = Round(dblSaldoAmortizacion - curValorAmortizacion, 12)
                dblAcumuladoAmortizacion = Round(dblAcumuladoAmortizacion + curValorAmortizacion, 12)
            Else
                dblValorInteres = dblSaldoAmortizacion * dblFactorAnualCupon 'dblTasaAnual
'                curValorAmortizacion = curValorCupon - dblValorInteres
                curValorAmortizacion = curValorNominal * CDbl(adoConsulta("PorcenAmortizacion")) * 0.01
                curValorCupon = curValorAmortizacion + dblValorInteres
                dblSaldoAmortizacion = dblSaldoAmortizacion - curValorAmortizacion
                dblAcumuladoAmortizacion = dblAcumuladoAmortizacion + curValorAmortizacion
    
                If CInt(lblNumCupones.Caption) = CInt(adoConsulta("NumCupon")) And dblSaldoAmortizacion <> 0 Then
                    dblAcumuladoAmortizacion = dblAcumuladoAmortizacion + dblSaldoAmortizacion
                    curValorAmortizacion = curValorAmortizacion + dblSaldoAmortizacion
                    dblValorInteres = dblValorInteres - dblSaldoAmortizacion
                    dblSaldoAmortizacion = 0
                End If
            End If
            
            strNumCupon = Trim(adoConsulta("NumCupon"))
            
            Call GrabarFactoresMontosTmp
            
            adoConsulta.MoveNext
        Loop
        adoConsulta.Close: Set adoConsulta = Nothing
    End With
        
End Sub

Private Sub GenerarFechasPago()

    Dim adoConsulta     As ADODB.Recordset
    Dim datFechaPago    As Date
    Dim intContador     As Integer, intNumDiasTmp       As Integer
    Dim intNumDiasPago  As Integer, intCuponVigente     As Integer
    
    Set adoConsulta = New ADODB.Recordset
    intCuponVigente = 0
    
    If chkAmortiza = Checked Then
        dblPorcenAmortizacion = 100 / CInt(lblNumCupones)
    Else
        dblPorcenAmortizacion = 0#
    End If
    
    With adoComm
        .CommandText = "SELECT CodTitulo,NumCupon,FechaVencimiento FROM InstrumentoInversionCalendarioTmp WHERE CodTitulo='" & Trim(txtCodigoValor.Text) & "' ORDER BY NumCupon"
        
        Set adoConsulta = .Execute
        
        Do While Not adoConsulta.EOF
            '*** Días calendario ***
            If strCodTipoDia = Codigo_Tipo_Dia_Calendario Then
                datFechaPago = DateAdd("d", CInt(txtNumDiasPago.Text), CVDate(adoConsulta("FechaVencimiento")))
                
            ElseIf CInt(txtNumDiasPago.Text) = 0 Then '*** Días útiles ***
                datFechaPago = CVDate(adoConsulta("FechaVencimiento"))

                If Not EsDiaUtil(datFechaPago) Then
                    datFechaPago = ProximoDiaUtil(datFechaPago)
                End If
            
            ElseIf CInt(txtNumDiasPago.Text) > 0 Then '*** Días útiles ***
                intContador = 1: intNumDiasTmp = 1: intNumDiasPago = CInt(txtNumDiasPago.Text)

                Do While intContador < 3
                    datFechaPago = DateAdd("d", intNumDiasTmp, CVDate(adoConsulta("FechaVencimiento")))
                    If Not EsDiaUtil(datFechaPago) Then
                        intNumDiasPago = intNumDiasPago + 1: intNumDiasTmp = intNumDiasTmp + 1
                    Else
                        intContador = intContador + 1
                    End If
                Loop
                
                datFechaPago = DateAdd("d", intNumDiasPago, CVDate(adoConsulta("FechaVencimiento")))
                
                If Not EsDiaUtil(datFechaPago) Then
                    datFechaPago = ProximoDiaUtil(datFechaPago)
                End If
            
            End If
            
            strFechaPago = Convertyyyymmdd(datFechaPago)
                        
            If DateDiff("d", CVDate(adoConsulta("FechaVencimiento")), gdatFechaActual) >= 0 Then
                strIndVencido = "X"
                strIndVigente = ""
            Else
                If intCuponVigente = 0 Then
                    strIndVigente = "X": intCuponVigente = 1
                Else
                    strIndVigente = " "
                End If
                strIndVencido = ""
            End If
            strNumCupon = Trim(adoConsulta("NumCupon"))
            
            If chkAmortiza = Unchecked And CInt(lblNumCupones.Caption) = CInt(adoConsulta("NumCupon")) Then
                dblPorcenAmortizacion = 100
            End If
            
            Call GrabarFechaPagoTmp
            
            adoConsulta.MoveNext
        Loop
        adoConsulta.Close: Set adoConsulta = Nothing
    End With
            
End Sub

Private Sub GenerarFechasCorte()

    Dim adoConsulta             As ADODB.Recordset
    Dim intNumSecuencial        As Integer, intNumMesesPeriodo  As Integer
    Dim datFechaIniValor        As Date, datFechaIniCupon       As Date
    Dim datFechaFinCupon        As Date, datFechaCorte1erCupon  As Date
    Dim datFechaInicioIndice    As Date, datFechaFinIndice      As Date
    Dim blnCupon                As Boolean, blnIndProceso       As Boolean
    Dim intNumDiasPeriodo       As Integer, intNumDias          As Integer
    Dim intNumPeriodosAnual     As Integer
        
    '*** Cronograma a partir de una fecha determinada ***
    If chkAPartir.Value Then
       datFechaIniValor = dtpFechaInicioCupon.Value
'       If strCodTipoVac = Codigo_Vac_InicioPrimerCupon Then datFechaInicioIndice = datFechaIniValor
    Else
       datFechaIniValor = dtpFechaEmision.Value  '*** Por defecto a partir de la fecha de emisión ***
'       If strCodTipoVac = Codigo_Vac_Emision Then datFechaInicioIndice = datFechaIniValor
    End If

    If strCodTipoVac = Codigo_Vac_InicioPrimerCupon Then datFechaInicioIndice = dtpFechaInicioCupon.Value
    If strCodTipoVac = Codigo_Vac_Emision Then datFechaInicioIndice = dtpFechaEmision.Value
    
    datFechaIniCupon = datFechaIniValor '*** Fecha de inicio del cupón ***

    With adoComm
        Set adoConsulta = New ADODB.Recordset

        '*** Obtener el número de días del peridodo de pago ***
        .CommandText = "SELECT ValorParametro FROM AuxiliarParametro WHERE CodTipoParametro='TIPFRE' AND CodParametro='" & strCodPeriodoPago & "'"
        Set adoConsulta = .Execute

        If Not adoConsulta.EOF Then
            intNumDiasPeriodo = CInt(adoConsulta("Valorparametro")) '*** Días del periodo  ***
            intNumMesesPeriodo = CInt(intNumDiasPeriodo / 30)       '*** Meses del periodo ***
            If intNumMesesPeriodo = 0 Then intNumMesesPeriodo = 1
            intNumPeriodosAnual = CInt(12 / intNumMesesPeriodo)     '*** Periodos al año   ***
        End If
        adoConsulta.Close: Set adoConsulta = Nothing
    End With
    
    '*** Cronograma a partir de la personalización de los días del mes ***
    If optDiasCupon(1).Value Then
        intNumDiasPeriodo = CInt(txtNumDiasPeriodo.Text) * intNumMesesPeriodo '*** Dias del periodo ***
    End If
    
    If strCodTipoVac = Codigo_Vac_InicioCuponAnterior Then
        datFechaInicioIndice = DateAdd("d", intNumDiasPeriodo * (intNumSecuencial - 1), dtpFechaCorteInicial)
    End If
    
    If strCodCuponCalculo = Codigo_Vac_FinCuponAnterior Then
        datFechaFinIndice = DateAdd("d", -1, datFechaIniCupon)
    End If

    intNumSecuencial = 0: blnCupon = True: blnIndProceso = False

    If chk1erCupon.Value Then
       datFechaCorte1erCupon = dtpFechaCorteInicial
    End If
    If strCodTipoVac = Codigo_Vac_InicioCuponVigente Then datFechaInicioIndice = datFechaIniCupon
    
    Do While blnCupon
       intNumSecuencial = intNumSecuencial + 1
       
       '*** Fecha de corte del primer cupón ***
       If chk1erCupon.Value = vbChecked And intNumSecuencial = 1 Then
          datFechaFinCupon = dtpFechaCorteInicial
          If strCodCuponCalculo = Codigo_Vac_FinPrimerCupon Then datFechaFinIndice = datFechaFinCupon
       Else
          '*** Mes calendario ***
          If optDiasCupon(0).Value = True Then
             datFechaFinCupon = DateAdd("m", intNumMesesPeriodo * (intNumSecuencial - 1), datFechaCorte1erCupon)

          '*** Días del periodo ***
          ElseIf optDiasCupon(1).Value = True Then
             datFechaFinCupon = DateAdd("d", CInt(txtNumDiasPeriodo.Text) * intNumMesesPeriodo - 1, datFechaIniCupon)

          '*** Fines de mes ***
          ElseIf optDiasCupon(2).Value = True Then
             datFechaFinCupon = DateAdd("d", -1, DateAdd("m", intNumMesesPeriodo, datFechaIniCupon))
             datFechaFinCupon = UltimaFechaMes(Month(datFechaFinCupon), Year(datFechaFinCupon))
          End If
       End If
       
       If strCodCuponCalculo = Codigo_Vac_FinCuponVigente Then datFechaFinIndice = datFechaFinCupon

       '*** SI EL NRO DE DIAS QUE FALTA PARA EL VCTO. YA NO CUBRE PARA GENERAR OTRO ***
       '*** CUPON, CONTROLAR QUE EL ULTIMO TENGA MAS DE UN DIA                      ***
       'intNumDias = DateDiff("d", datFechaFinCupon, dtpFechaVencimiento.Value) - 1

       intNumDias = CalculaDias(datFechaFinCupon, dtpFechaVencimiento.Value, strCodBaseCalculo)

       If intNumDias <= 0 Then
          blnIndProceso = True
       Else
          blnIndProceso = False
          If dtpFechaVencimiento = datFechaFinCupon Then
             blnIndProceso = True
          Else
             blnIndProceso = False
          End If
       End If
       
       '*** Inicio para el último cupón
       If blnIndProceso = True Then
            datFechaFinCupon = dtpFechaVencimiento
            strNumCupon = Format(intNumSecuencial, "000")
            strFechaInicio = Convertyyyymmdd(datFechaIniCupon)
            strFechaFin = Convertyyyymmdd(datFechaFinCupon)
            intCantDias = DateDiff("d", datFechaIniCupon, datFechaFinCupon) '- 1
            'intCantDias = CalculaDias(datFechaIniCupon, datFechaFinCupon, strCodBaseCalculo)
            
            If strCodTipoAjuste <> Codigo_Tipo_Ajuste_Vac Then
                strFechaInicioIndice = Convertyyyymmdd(CVDate(Valor_Fecha))
                strFechaFinIndice = Convertyyyymmdd(CVDate(Valor_Fecha))
            Else
                strFechaInicioIndice = Convertyyyymmdd(datFechaInicioIndice)
                strFechaFinIndice = Convertyyyymmdd(datFechaFinIndice)
            End If

            '*** Grabar en temporal ***
            Call GrabarFechaCorteTmp
            Call GrabarFechaIndiceTmp
            
            blnCupon = False: Exit Do
       End If

       strNumCupon = Format(intNumSecuencial, "000")
       strFechaInicio = Convertyyyymmdd(datFechaIniCupon)
       strFechaFin = Convertyyyymmdd(datFechaFinCupon)
       intCantDias = DateDiff("d", datFechaIniCupon, datFechaFinCupon) '- 1

       'intCantDias = CalculaDias(datFechaIniCupon, datFechaFinCupon, strCodBaseCalculo)
       
       If strCodCuponCalculo = Codigo_Vac_Liquidacion Then datFechaFinIndice = CVDate(Valor_Fecha)
       
       If strCodTipoAjuste <> Codigo_Tipo_Ajuste_Vac Then
            strFechaInicioIndice = Convertyyyymmdd(CVDate(Valor_Fecha))
            strFechaFinIndice = Convertyyyymmdd(CVDate(Valor_Fecha))
       Else
            strFechaInicioIndice = Convertyyyymmdd(datFechaInicioIndice)
            strFechaFinIndice = Convertyyyymmdd(datFechaFinIndice)
       End If
       
       '*** Grabar en temporal ***
       Call GrabarFechaCorteTmp
       Call GrabarFechaIndiceTmp
       
       '*** Para siguiente entrada ***
       datFechaIniCupon = datFechaFinCupon 'DateAdd("d", 1, datFechaFinCupon)
       If strCodTipoVac = Codigo_Vac_InicioCuponVigente Then datFechaInicioIndice = datFechaIniCupon

    Loop

    Set adoConsulta = New ADODB.Recordset
    
    With adoComm
        .CommandText = "SELECT COUNT(NumCupon) NumRegistros FROM InstrumentoInversionCalendarioTmp " & _
            "WHERE CodTitulo='" & Trim(txtCodigoValor.Text) & "'"
            
        Set adoConsulta = .Execute
        
        If Not adoConsulta.EOF Then
            If CInt(adoConsulta("NumRegistros")) > intNumSecuencial Then
                .CommandText = "DELETE InstrumentoInversionCalendarioTmp WHERE CodTitulo='" & Trim(txtCodigoValor.Text) & "' AND " & _
                    "NumSecuencial >" & intNumSecuencial
                    
                adoConn.Execute .CommandText
            End If
        End If
        adoConsulta.Close: Set adoConsulta = Nothing
    End With
    
    lblNumCupones.Caption = intNumSecuencial
    
End Sub

Public Sub Grabar()

    If strEstado = Reg_Consulta Then Exit Sub
    
    Dim intCantRegistros    As Integer, intRegistro         As Integer
    Dim adoRegistro         As ADODB.Recordset
    Dim strNumAsiento       As String
    Dim strFechaEmision     As String, strFechaVencimiento  As String
            
    If strEstado = Reg_Adicion Then
        If TodoOK() Then
            Dim dblTasaAnual    As Double
            Dim dblFactor1      As Double, dblFactor2       As Double
            Dim intDiasPeriodo  As Integer, intDiasPlazo    As Integer
            Dim strSQLCupon     As String
            
            strFechaEmision = Convertyyyymmdd(dtpFechaEmision.Value)
            strFechaVencimiento = Convertyyyymmdd(dtpFechaVencimiento.Value)
            
            Me.MousePointer = vbHourglass
                                                
            With adoComm
                '*** Obtener el número de la analítica ***
                .CommandText = "{call up_ACSelDatosParametro(21,'" & strCodTipoInstrumento & "') }"
                Set adoRegistro = .Execute
                
                If Not adoRegistro.EOF Then
                    strCodAnalitica = Format(CInt(adoRegistro("NumUltimo")) + 1, "00000000")
                End If
                adoRegistro.Close: Set adoRegistro = Nothing
                
                strIndGenerado = "X"
                If optTipoCodigo(0).Value Then strIndGenerado = Valor_Caracter
                
                If strIndTasaAjustada <> Valor_Indicador Then strCodTipoAjuste = Valor_Caracter
                
                If strIndCuponCero = Valor_Indicador And CDbl(txtTasaAnual.Text) > 0 Then
                    dblTasaAnual = CDbl(txtTasaAnual.Text)
                    dblFactor1 = CDbl(txtTasaAnual.Text)
                    dblFactor2 = CDbl(txtTasaAnual.Text)
                    
                    curValorCupon = CDbl(txtValorNominal.Text) * dblTasaAnual * 0.01
                    dblSaldoAmortizacion = CDbl(txtValorNominal.Text)
                    '*** Base de Cálculo ***
                    intDiasPeriodo = 365
                    Select Case strCodBaseCalculo
                        Case Codigo_Base_30_360: intDiasPeriodo = 360
                        Case Codigo_Base_Actual_365: intDiasPeriodo = 365
                        Case Codigo_Base_Actual_360: intDiasPeriodo = 360
                        Case Codigo_Base_30_365: intDiasPeriodo = 365
                    End Select
                    
                    intDiasPlazo = DateDiff("d", dtpFechaEmision.Value, dtpFechaVencimiento.Value)
                    
                    '*** Calculando factores ***
                    dblFactorAnualCupon = FactorAnual(dblFactor1, intDiasPlazo, intDiasPeriodo, strCodTipoTasa, Valor_Indicador, Codigo_Calculo_Normal, 0, intDiasPlazo, 1)
                    dblTasaCuponNormal = FactorAnualNormal(dblFactor2, intDiasPlazo, intDiasPeriodo, strCodTipoTasa, Valor_Indicador, Codigo_Calculo_Normal, 0, intDiasPlazo, 1)
                    dblFactorDiario = FactorDiario(dblFactorAnualCupon, intDiasPlazo, strCodTipoTasa, Valor_Indicador, intDiasPlazo)
                    dblFactorDiarioNormal = FactorDiarioNormal(dblTasaCuponNormal, intDiasPlazo, strCodTipoTasa, Valor_Indicador, intDiasPlazo)
                
                    If strCodTipoTasa = Codigo_Tipo_Tasa_Efectiva Then '*** Efectiva ***
                       dblFactorAnualCupon = ((1 + (0.01 * dblFactor1)) ^ (intDiasPlazo / intDiasPeriodo)) - 1
                       dblTasaCuponNormal = ((1 + (0.01 * dblFactor2)) ^ (intDiasPlazo / intDiasPeriodo)) - 1
                       dblFactorDiario = ((1 + dblFactorAnualCupon) ^ (1 / intDiasPlazo)) - 1
                       
                    Else '*** Nominal Capitalizable ***
                       dblFactorAnualCupon = (0.01 * dblFactor1) / (intDiasPeriodo / intDiasPlazo)
                       dblTasaCuponNormal = (0.01 * dblFactor2) / (intDiasPeriodo / intDiasPlazo)
                       dblFactorDiario = dblFactorAnualCupon / intDiasPlazo
                       
    '                ElseIf Not chkCapitalizable.Value Then '*** Nominal No Capitalizable ***
    '                   dblFactorAnualCupon = (0.01 * dblFactor1) / intNumPeriodosAnual
    '                   dblTasaCuponNormal = (0.01 * dblFactor2) / intNumPeriodosAnual
    '                   dblFactorDiario = dblFactorAnualCupon / intDiasPlazo
                       
                    End If
                    
                    dblValorInteres = dblSaldoAmortizacion * dblTasaAnual
                    curValorAmortizacion = curValorCupon - dblValorInteres
                    dblSaldoAmortizacion = dblSaldoAmortizacion - curValorAmortizacion
                    dblAcumuladoAmortizacion = dblAcumuladoAmortizacion + curValorAmortizacion
                
                    strSQLCupon = "{ call up_IVManInstrumentoInversionCalendario('" & Trim(txtCodigoValor.Text) & "','001','" & strCodTipoInstrumento & "','" & _
                    strCodAnalitica & "',1,'" & strFechaEmision & "','" & strFechaVencimiento & "','" & strFechaVencimiento & "'," & _
                    CDbl(txtTasaAnual.Text) & "," & CDbl(dblFactorAnualCupon) & "," & CDbl(dblFactorDiario) & "," & CDbl(dblValorInteres) & "," & _
                    CDbl(curValorAmortizacion) & "," & CDbl(dblSaldoAmortizacion) & "," & dblAcumuladoAmortizacion & "," & _
                    CDbl(curValorCupon) & ",'',0," & intDiasPlazo & ",'','X'," & dblTasaCuponNormal & ",'I') }"
                End If
                
                If strCodTipoInstrumento = "004" Then
                    strFechaEmision = Convertyyyymmdd(CVDate(Valor_Fecha))
                    strFechaVencimiento = Convertyyyymmdd(CVDate(Valor_Fecha))
                End If
                                            
                .CommandText = "BEGIN TRAN ProcInstrumento"
                adoConn.Execute .CommandText
                
                On Error GoTo Ctrl_Error
                
                'strCodMercado
                
                '*** Datos del título ***
                .CommandText = "{ call up_IVManInstrumentoInversion('" & Trim(txtCodigoValor.Text) & "','" & _
                    strCodTipoInstrumento & "','" & strCodAnalitica & "','" & strCodClaseInstrumento & "','" & strCodSubClaseInstrumento & "','" & _
                    strCodAnalitica & "','','','','" & Trim(txtNemonico.Text) & "','" & Trim(txtDescripValor.Text) & "','" & Trim(txtNumEmision.Text) & "','" & _
                    Trim(txtNumSerie.Text) & "','" & strFechaEmision & "','" & strFechaVencimiento & "'," & DateDiff("d", dtpFechaEmision, dtpFechaVencimiento) & ",'" & _
                    strCodEmisor & "','','" & strCodCiiu & "','" & strCodGrupo & "','" & strCodSector & "','" & strCodMercado & "'," & CDec(txtValorNominal.Text) & ",'" & _
                    strCodMonedaEmision & "','" & strCodMonedaPago & "',0,0,0,'" & strCodTipoTasa & "','" & strCodTasa & "','" & strCodBaseCalculo & "','" & strCodFormaCalculo & "','" & _
                    strCodPeriodoPago & "','','" & strIndTasaAjustada & "'," & CDec(txtTasaAnual.Text) & ",0,0," & CInt(lblNumCupones.Caption) & ",'" & _
                    strCodTipoAjuste & "','" & strCodTipoVac & "','" & strCodCuponCalculo & "','" & strIndAmortizacion & "','" & strIndGenerado & "','X','" & strIndQuiebre & "','" & _
                    strIndTasaCapitalizable & "','" & strIndCuponCero & "','01','" & strCodRiesgo & "','" & strCodSubRiesgo & "','" & _
                    gstrLogin & "','" & Convertyyyymmdd(gdatFechaActual) & Space(1) & Format(Time, "hh:mm") & "','I') }"
                adoConn.Execute .CommandText
                                
                '*** Cronograma ***
                If strIndCuponCero = Valor_Indicador And CDbl(txtTasaAnual.Text) > 0 Then
                    adoConn.Execute strSQLCupon
                Else
                    .CommandText = "INSERT INTO InstrumentoInversionCalendario " & _
                        "SELECT CodTitulo,NumCupon,CodFile,'" & strCodAnalitica & "',NumSecuencial,FechaInicio,FechaVencimiento," & _
                        "FechaPago,TasaInteres,FactorInteres,FactorDiario,ValorInteres,ValorAmortizacion,SaldoAmortizacion," & _
                        "AcumuladoAmortizacion,ValorCupon,IndVencido,PorcenAmortizacion,CantDiasPeriodo,IndConfirma," & _
                        "IndVigente,FactorInteres1,FactorDiario1,FechaInicioIndice,FechaFinIndice FROM InstrumentoInversionCalendarioTmp " & _
                        "WHERE CodTitulo='" & Trim(txtCodigoValor.Text) & "'"
                    adoConn.Execute .CommandText
                End If
                                                                                
                '*** Actualizar el número de analítica **
                .CommandText = "UPDATE InversionFile SET NumUltimo = NumUltimo + 1 " & _
                    "WHERE CodFile='" & strCodTipoInstrumento & "'"
                adoConn.Execute .CommandText
                
                .CommandText = "DELETE InstrumentoInversionCalendarioTmp " & _
                    "WHERE CodTitulo='" & Trim(txtCodigoValor) & "'"
                adoConn.Execute .CommandText
                                
                .CommandText = "COMMIT TRAN ProcInstrumento"
                adoConn.Execute .CommandText
                                                                                
            End With
                
            Me.MousePointer = vbDefault
        
            MsgBox Mensaje_Adicion_Exitosa, vbExclamation
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
            cmdOpcion.Visible = True
            With tabTitulos
                .TabEnabled(0) = True
                .Tab = 0
            End With
            
            Call Buscar
        End If
    End If
    
    If strEstado = Reg_Edicion Then
        If TodoOK() Then
            strFechaEmision = Convertyyyymmdd(dtpFechaEmision.Value)
            strFechaVencimiento = Convertyyyymmdd(dtpFechaVencimiento.Value)
            
            Me.MousePointer = vbHourglass
                                                
            With adoComm
                strIndGenerado = "X"
                If optTipoCodigo(0).Value Then strIndGenerado = Valor_Caracter
                
                If strIndTasaAjustada <> Valor_Indicador Then strCodTipoAjuste = Valor_Caracter
                
                .CommandText = "BEGIN TRAN ProcInstrumento"
                adoConn.Execute .CommandText
                
                On Error GoTo Ctrl_Error
                
                '*** Datos del título ***
                .CommandText = "{ call up_IVManInstrumentoInversion('" & Trim(txtCodigoValor.Text) & "','" & _
                    strCodTipoInstrumento & "','" & strCodAnalitica & "','" & strCodClaseInstrumento & "','" & strCodSubClaseInstrumento & "','" & _
                    strCodAnalitica & "','','','','" & Trim(txtNemonico.Text) & "','" & Trim(txtDescripValor.Text) & "','" & Trim(txtNumEmision.Text) & "','" & _
                    Trim(txtNumSerie.Text) & "','" & strFechaEmision & "','" & strFechaVencimiento & "'," & DateDiff("d", dtpFechaEmision, dtpFechaVencimiento) & ",'" & _
                    strCodEmisor & "','','" & strCodCiiu & "','" & strCodGrupo & "','" & strCodSector & "','" & strCodMercado & "'," & CDec(txtValorNominal.Text) & ",'" & _
                    strCodMonedaEmision & "','" & strCodMonedaPago & "',0,0,0,'" & strCodTipoTasa & "','" & strCodTasa & "','" & strCodBaseCalculo & "','" & strCodFormaCalculo & "','" & _
                    strCodPeriodoPago & "','','" & strIndTasaAjustada & "'," & CDec(txtTasaAnual.Text) & ",0,0," & CInt(lblNumCupones.Caption) & ",'" & _
                    strCodTipoAjuste & "','" & strCodTipoVac & "','" & strCodCuponCalculo & "','" & strIndAmortizacion & "','" & strIndGenerado & "','X','" & strIndQuiebre & "','" & _
                    strIndTasaCapitalizable & "','" & strIndCuponCero & "','01','" & strCodRiesgo & "','" & strCodSubRiesgo & "','" & _
                    gstrLogin & "','" & Convertyyyymmdd(gdatFechaActual) & Space(1) & Format(Time, "hh:mm") & "','U') }"
                adoConn.Execute .CommandText
                                
                '*** Eliminar en Tabla Definitiva ***
                .CommandText = "DELETE InstrumentoInversionCalendario " & _
                    "WHERE CodTitulo='" & Trim(txtCodigoValor) & "'"
                adoConn.Execute .CommandText
                            
                '*** Cronograma ***
                .CommandText = "INSERT INTO InstrumentoInversionCalendario " & _
                    "SELECT CodTitulo,NumCupon,CodFile,'" & strCodAnalitica & "',NumSecuencial,FechaInicio,FechaVencimiento," & _
                    "FechaPago,TasaInteres,FactorInteres,FactorDiario,ValorInteres,ValorAmortizacion,SaldoAmortizacion," & _
                    "AcumuladoAmortizacion,ValorCupon,IndVencido,PorcenAmortizacion,CantDiasPeriodo,IndConfirma," & _
                    "IndVigente,FactorInteres1,FactorDiario1,FechaInicioIndice,FechaFinIndice FROM InstrumentoInversionCalendarioTmp " & _
                    "WHERE CodTitulo='" & Trim(txtCodigoValor.Text) & "'"
                adoConn.Execute .CommandText, intRegistro
                
                '*** Eliminar en Tabla Temporal ***
                .CommandText = "DELETE InstrumentoInversionCalendarioTmp " & _
                    "WHERE CodTitulo='" & Trim(txtCodigoValor) & "'"
                adoConn.Execute .CommandText
                                                                                                
                .CommandText = "COMMIT TRAN ProcInstrumento"
                adoConn.Execute .CommandText
                                                                                
            End With
                
            Me.MousePointer = vbDefault
        
            MsgBox Mensaje_Edicion_Exitosa, vbExclamation
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
            cmdOpcion.Visible = True
            With tabTitulos
                .TabEnabled(0) = True
                .Tab = 0
            End With
            
            Call Buscar
        End If
    End If
    Exit Sub
    
Ctrl_Error:
    adoComm.CommandText = "ROLLBACK TRAN ProcInstrumento"
    adoConn.Execute adoComm.CommandText
    
    MsgBox Mensaje_Proceso_NoExitoso, vbCritical
    Me.MousePointer = vbDefault
            
End Sub

Private Function TodoOK() As Boolean
        
    TodoOK = False
          
    If Trim(txtCodigoValor.Text) = Valor_Caracter Then
        MsgBox "Debe indicar el Código ISIN.", vbCritical, Me.Caption
        If txtCodigoValor.Enabled Then txtCodigoValor.SetFocus
        Exit Function
    End If
    
    If Trim(txtNemonico.Text) = Valor_Caracter Then
        MsgBox "Debe indicar el Código Nemotécnico.", vbCritical, Me.Caption
        If txtNemonico.Enabled Then txtNemonico.SetFocus
        Exit Function
    End If
    
    If Trim(txtDescripValor.Text) = Valor_Caracter Then
        MsgBox "Debe indicar la Descripción del Instrumento.", vbCritical, Me.Caption
        If txtDescripValor.Enabled Then txtDescripValor.SetFocus
        Exit Function
    End If
    
    If cboTipoInstrumento.ListIndex <= 0 Then
        MsgBox "Debe seleccionar el Tipo de Instrumento.", vbCritical, Me.Caption
        If cboTipoInstrumento.Enabled Then cboTipoInstrumento.SetFocus
        Exit Function
    End If
    
    If cboClaseInstrumento.ListCount > 1 Then
        If cboClaseInstrumento.ListIndex <= 0 Then
            MsgBox "Debe seleccionar la Clase de Instrumento.", vbCritical, Me.Caption
            If cboClaseInstrumento.Enabled Then cboClaseInstrumento.SetFocus
            Exit Function
        End If
    End If
    
    If cboSubClaseInstrumento.ListCount > 1 Then
        If cboSubClaseInstrumento.ListIndex <= 0 Then
            MsgBox "Debe seleccionar la SubClase de Instrumento.", vbCritical, Me.Caption
            If cboSubClaseInstrumento.Enabled Then cboSubClaseInstrumento.SetFocus
            Exit Function
        End If
    End If
        
    If cboEmisor.ListIndex <= 0 Then
        MsgBox "Debe seleccionar el Emisor.", vbCritical, Me.Caption
        If cboEmisor.Enabled Then cboEmisor.SetFocus
        Exit Function
    End If
                              
    If cboMonedaEmision.ListIndex <= 0 Then
        MsgBox "Debe seleccionar la Moneda de Emisión.", vbCritical, Me.Caption
        If cboMonedaEmision.Enabled Then cboMonedaEmision.SetFocus
        Exit Function
    End If
        
    If cboMonedaPago.ListIndex <= 0 Then
        MsgBox "Debe seleccionar la Moneda de Pago.", vbCritical, Me.Caption
        If cboMonedaPago.Enabled Then cboMonedaPago.SetFocus
        Exit Function
    End If
    
    If cboMercado.ListIndex < 0 Then
        MsgBox "Debe seleccionar el mercado donde cotiza el valor", vbCritical, Me.Caption
        If cboMercado.Enabled Then cboMercado.SetFocus
        Exit Function
    End If
                
    If CCur(txtValorNominal.Text) = 0 Then
        MsgBox "Debe indicar el Valor Nominal.", vbCritical, Me.Caption
        If txtValorNominal.Enabled Then txtValorNominal.SetFocus
        Exit Function
    End If
        
    '*** Si todo paso OK ***
    TodoOK = True
  
End Function

Private Function ValidaDatos() As Boolean
        
    ValidaDatos = False
          
    If Trim(txtCodigoValor.Text) = Valor_Caracter Then
        MsgBox "Debe indicar el Código ISIN.", vbCritical, Me.Caption
        If txtCodigoValor.Enabled Then txtCodigoValor.SetFocus
        Exit Function
    End If
        
    If cboTipoInstrumento.ListIndex <= 0 Then
        MsgBox "Debe seleccionar el Tipo de Instrumento.", vbCritical, Me.Caption
        If cboTipoInstrumento.Enabled Then cboTipoInstrumento.SetFocus
        Exit Function
    End If
    
    If cboClaseInstrumento.ListCount > 1 Then
        If cboClaseInstrumento.ListIndex <= 0 Then
            MsgBox "Debe seleccionar la Clase de Instrumento.", vbCritical, Me.Caption
            If cboClaseInstrumento.Enabled Then cboClaseInstrumento.SetFocus
            Exit Function
        End If
    End If
    
    If cboSubClaseInstrumento.ListCount > 1 Then
        If cboSubClaseInstrumento.ListIndex <= 0 Then
            MsgBox "Debe seleccionar la SubClase de Instrumento.", vbCritical, Me.Caption
            If cboSubClaseInstrumento.Enabled Then cboSubClaseInstrumento.SetFocus
            Exit Function
        End If
    End If
                    
    If CCur(txtValorNominal.Text) = 0 Then
        MsgBox "Debe indicar el Valor Nominal.", vbCritical, Me.Caption
        If txtValorNominal.Enabled Then txtValorNominal.SetFocus
        Exit Function
    End If
        
    '*** Si todo paso OK ***
    ValidaDatos = True
  
End Function

Public Sub Imprimir()

End Sub


Public Sub SubImprimir(Index As Integer)

    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    Dim strFechaDesde           As String, strFechaHasta        As String

    If tabTitulos.Tab = 1 Then Exit Sub
    
    Select Case Index
        Case 1, 2
        
            If Index = 1 Then
                gstrNameRepo = "InstrumentoInversionCaracteristicas"
            Else
                gstrNameRepo = "InstrumentoInversionCalendario"
            End If
            
            Set frmReporte = New frmVisorReporte

            ReDim aReportParamS(1)
            ReDim aReportParamFn(2)
            ReDim aReportParamF(2)
                        
            aReportParamFn(0) = "Usuario"
            aReportParamFn(1) = "Hora"
            aReportParamFn(2) = "NombreEmpresa"
                        
            aReportParamF(0) = gstrLogin
            aReportParamF(1) = Format(Time(), "hh:mm:ss")
            aReportParamF(2) = gstrNombreEmpresa & Space(1)
                        
            aReportParamS(0) = Trim(tdgConsulta.Columns(0).Value)
            aReportParamS(1) = gstrCodAdministradora
            
    End Select

    gstrSelFrml = ""
    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"

    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())

    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
    frmReporte.Show vbModal

    Set frmReporte = Nothing

    Screen.MousePointer = vbNormal
    
End Sub

Public Sub Modificar()

    If strEstado = Reg_Consulta Then
        strEstado = Reg_Edicion
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabTitulos
            .TabEnabled(0) = False
            .TabEnabled(1) = True
            .TabEnabled(2) = True
            .Tab = 1
        End With
        Call Deshabilita
    End If
    
End Sub

Public Sub Salir()

    Unload Me
    
End Sub

Private Function ValidaFechasyTasas() As Boolean
    
    ValidaFechasyTasas = False

    If dtpFechaVencimiento <= dtpFechaEmision Then
        MsgBox "La Fecha de Vencimiento debe ser posterior a la Fecha de Emisión", vbOKOnly + vbCritical, Me.Caption
        If dtpFechaVencimiento.Visible Then dtpFechaVencimiento.SetFocus
        Exit Function
    End If

    If dtpFechaVencimiento <= gdatFechaActual Then
        MsgBox "La Fecha de Vencimiento debe ser posterior a la Fecha Vigente del Fondo.", vbOKOnly + vbCritical, Me.Caption
        If dtpFechaVencimiento.Visible Then dtpFechaVencimiento.SetFocus
        Exit Function
    End If
            
    If CCur(txtValorNominal.Text) <= 0 Then
        MsgBox "Debe indicar el Valor Nominal del TITULO.", vbOKOnly + vbCritical, Me.Caption
        If txtValorNominal.Enabled Then txtValorNominal.SetFocus
        Exit Function
    End If
    
    If Not chkCuponCero.Value And CDbl(txtTasaAnual.Text) <= 0 Then
        MsgBox "Debe indicar la Tasa de Interés Anual (Tasa Cupón) del TITULO.", vbOKOnly + vbCritical, Me.Caption
        txtTasaAnual.SetFocus
        Exit Function
    End If
    
    '*** Si todo paso ok ***
    ValidaFechasyTasas = True
    
End Function

Private Sub cboBaseCalculo_Click()

    strCodBaseCalculo = Valor_Caracter
    If cboBaseCalculo.ListIndex < 0 Then Exit Sub
    
    strCodBaseCalculo = Trim(arrBaseCalculo(cboBaseCalculo.ListIndex))
    
End Sub


Private Sub cboClaseInstrumento_Click()

    strCodClaseInstrumento = ""
    If cboClaseInstrumento.ListIndex < 0 Then Exit Sub
    
    strCodClaseInstrumento = Trim(arrClaseInstrumento(cboClaseInstrumento.ListIndex))
    
    '*** Grupo de Instrumento ***
    strSQL = "SELECT CodSubDetalleFile CODIGO,DescripSubDetalleFile DESCRIP FROM InversionSubDetalleFile WHERE " & _
        "CodDetalleFile='" & strCodClaseInstrumento & "' AND CodFile='" & strCodTipoInstrumento & "' AND IndVigente='X' ORDER BY DescripSubDetalleFile"
        
    CargarControlLista strSQL, cboSubClaseInstrumento, arrSubClaseInstrumento(), Sel_Defecto
    
    If cboSubClaseInstrumento.ListCount > 0 Then cboSubClaseInstrumento.ListIndex = 0
    
End Sub

Private Sub cboCuponCalculo_Click()

    strCodCuponCalculo = ""
    If cboCuponCalculo.ListIndex < 0 Then Exit Sub
    
    strCodCuponCalculo = Trim(arrCuponCalculo(cboCuponCalculo.ListIndex))
    
End Sub


Private Sub cboFormaCalculo_Click()

    strCodFormaCalculo = Valor_Caracter
    If cboFormaCalculo.ListIndex < 0 Then Exit Sub
    
    strCodFormaCalculo = Trim(arrFormaCalculo(cboFormaCalculo.ListIndex))
    
End Sub



Private Sub cboMercado_Click()

    strCodMercado = ""
    If cboMercado.ListIndex < 0 Then Exit Sub
    
    strCodMercado = Trim(arrMercado(cboMercado.ListIndex))

End Sub

Private Sub cboMonedaEmision_Click()

    Dim intRegistro As Integer
    
    strCodMonedaEmision = ""
    If cboMonedaEmision.ListIndex < 0 Then Exit Sub
    
    strCodMonedaEmision = Trim(arrMonedaEmision(cboMonedaEmision.ListIndex))
        
    intRegistro = ObtenerItemLista(arrMonedaPago(), strCodMonedaEmision)
    If intRegistro >= 0 Then cboMonedaPago.ListIndex = intRegistro
    
End Sub


Private Sub cboMonedaPago_Click()

    strCodMonedaPago = ""
    If cboMonedaPago.ListIndex < 0 Then Exit Sub
    
    strCodMonedaPago = Trim(arrMonedaPago(cboMonedaPago.ListIndex))
    
End Sub


Private Sub cboClaseCriterio_Click()

    strCodClaseCriterio = Valor_Caracter
    If cboClaseCriterio.ListIndex < 0 Then Exit Sub
    
    strCodClaseCriterio = Trim(arrClaseCriterio(cboClaseCriterio.ListIndex))
    
    '*** Grupo de Instrumento ***
    strSQL = "SELECT CodSubDetalleFile CODIGO,DescripSubDetalleFile DESCRIP FROM InversionSubDetalleFile WHERE " & _
        "CodDetalleFile='" & strCodClaseCriterio & "' AND CodFile='" & strCodTipoCriterio & "' AND IndVigente='X' ORDER BY DescripSubDetalleFile"
        
    CargarControlLista strSQL, cboGrupoCriterio, arrSubClaseCriterio(), Sel_Defecto
    
    If cboGrupoCriterio.ListCount > 0 Then cboGrupoCriterio.ListIndex = 0
    
End Sub


Private Sub cboSubClaseInstrumento_Click()

    strCodSubClaseInstrumento = ""
    If cboSubClaseInstrumento.ListIndex < 0 Then Exit Sub
    
    strCodSubClaseInstrumento = Trim(arrSubClaseInstrumento(cboSubClaseInstrumento.ListIndex))
   
End Sub

Private Sub cboGrupoCriterio_Click()

    strCodSubClaseCriterio = ""
    If cboGrupoCriterio.ListIndex < 0 Then Exit Sub
    
    strCodSubClaseCriterio = Trim(arrSubClaseCriterio(cboGrupoCriterio.ListIndex))
    
End Sub




Private Sub cboTasa_Click()

    strCodTasa = Valor_Caracter
    If cboTasa.ListIndex < 0 Then Exit Sub
    
    strCodTasa = Trim(arrTasa(cboTasa.ListIndex))
    
End Sub


Private Sub cboTipoAjuste_Click()

    strCodTipoAjuste = Valor_Caracter
    If cboTipoAjuste.ListIndex < 0 Then Exit Sub
    
    strCodTipoAjuste = Trim(arrTipoAjuste(cboTipoAjuste.ListIndex))
    
    If strCodTipoAjuste = Codigo_Tipo_Ajuste_Vac Then
        lblDescrip(14).Caption = "Inicio Indice"
        strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPVAC' ORDER BY DescripParametro"
        CargarControlLista strSQL, cboTipoVac, arrTipoVac(), Sel_Defecto
        
        If cboTipoVac.ListCount > 0 Then cboTipoVac.ListIndex = 0
        cboTipoVac.Enabled = True
        
        lblDescrip(15).Caption = "Fin Indice"
        strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='CUPCAL' ORDER BY DescripParametro"
        CargarControlLista strSQL, cboCuponCalculo, arrCuponCalculo(), Sel_Defecto
        
        If cboCuponCalculo.ListCount > 0 Then cboCuponCalculo.ListIndex = 0
        cboCuponCalculo.Enabled = True
    Else
        lblDescrip(14).Caption = "Clase Ajuste"
        strSQL = "SELECT CodClaseTasa CODIGO,DescripClaseTasa DESCRIP FROM ClaseTasa WHERE CodTasa='" & strCodTipoAjuste & "' ORDER BY DescripClaseTasa"
        CargarControlLista strSQL, cboTipoVac, arrTipoVac(), Sel_Defecto
        
        If cboTipoVac.ListCount > 0 Then cboTipoVac.ListIndex = 0
        cboTipoVac.Enabled = True
        
        lblDescrip(15).Caption = Valor_Caracter
        If cboCuponCalculo.ListCount > 0 Then cboCuponCalculo.ListIndex = 0
        cboCuponCalculo.Enabled = False
    End If
    
End Sub


Private Sub cboTipoDia_Click()

    strCodTipoDia = ""
    If cboTipoDia.ListIndex < 0 Then Exit Sub
    
    strCodTipoDia = Trim(arrTipoDia(cboTipoDia.ListIndex))
    
End Sub


Private Sub cboTipoInstrumento_Click()

    Dim adoTemporal     As ADODB.Recordset
    
    strCodTipoInstrumento = Valor_Caracter
    If cboTipoInstrumento.ListIndex < 0 Then Exit Sub
    
    strCodTipoInstrumento = Trim(arrTipoInstrumento(cboTipoInstrumento.ListIndex))
    
    lblAnalitica = strCodTipoInstrumento & "-????????"
    
    '*** Clase de Instrumento ***
    strSQL = "SELECT CodDetalleFile CODIGO,DescripDetalleFile DESCRIP FROM InversionDetalleFile WHERE CodFile='" & strCodTipoInstrumento & "' AND IndVigente='X' ORDER BY DescripDetalleFile"
    CargarControlLista strSQL, cboClaseInstrumento, arrClaseInstrumento(), Sel_Defecto
    
    If cboClaseInstrumento.ListCount > 0 Then cboClaseInstrumento.ListIndex = 0
        
    Set adoTemporal = New ADODB.Recordset
    If strEstado = Reg_Adicion Then
    If optTipoCodigo(1).Value Then
        
        txtCodigoValor.Text = Trim(gstrInicialTitulo) & Format(CLng(gstrPeriodoActual) & CLng(gstrCodAdministradora) & CLng(strCodTipoInstrumento), String(12, "0"))
                                
        adoComm.CommandText = "SELECT DescripFile,DescripInicial FROM InversionFile WHERE CodFile='" & strCodTipoInstrumento & "'"
        Set adoTemporal = adoComm.Execute
        
        If Not adoTemporal.EOF Then
            txtNemonico.Text = Trim(adoTemporal("DescripInicial")) & CStr(gdatFechaActual)
        End If
        adoTemporal.Close
    End If
    End If
    
    adoComm.CommandText = "SELECT TipoPlazo FROM InversionFile WHERE CodFile='" & strCodTipoInstrumento & "'"
    Set adoTemporal = adoComm.Execute
    
    If Not adoTemporal.EOF Then
        strCodTipoPlazo = Trim(adoTemporal("TipoPlazo"))
    End If
    adoTemporal.Close: Set adoTemporal = Nothing
    
End Sub

Private Sub cboTipoCriterio_Click()

    strCodTipoCriterio = ""
    If cboTipoCriterio.ListIndex < 0 Then Exit Sub
    
    strCodTipoCriterio = Trim(arrTipoCriterio(cboTipoCriterio.ListIndex))
    
    '*** Clase de Instrumento ***
    strSQL = "SELECT CodDetalleFile CODIGO,DescripDetalleFile DESCRIP FROM InversionDetalleFile WHERE CodFile='" & strCodTipoCriterio & "' AND IndVigente='X' and SUBSTRING (LTRIM(RTRIM(DescripDetalleFile)),1,1)<>'6' ORDER BY DescripDetalleFile"
    CargarControlLista strSQL, cboClaseCriterio, arrClaseCriterio(), Sel_Defecto
    
    If cboClaseCriterio.ListCount > 0 Then cboClaseCriterio.ListIndex = 0
    
End Sub


Private Sub cboTipoTasa_Click()

    strCodTipoTasa = Valor_Caracter
    If cboTipoTasa.ListIndex < 0 Then Exit Sub
    
    strCodTipoTasa = Trim(arrTipoTasa(cboTipoTasa.ListIndex))
    
    chkCapitalizable = vbChecked
    chkCapitalizable.Enabled = True
    If strCodTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
        chkCapitalizable = vbUnchecked
        chkCapitalizable.Enabled = False
    End If
    
End Sub


Private Sub cboTipoVac_Click()

    strCodTipoVac = Valor_Caracter
    If cboTipoVac.ListIndex < 0 Then Exit Sub
    
    strCodTipoVac = Trim(arrTipoVac(cboTipoVac.ListIndex))

'    If strCodTipoVac = Codigo_Vac_Periodico Then
'        strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='CUPCAL' ORDER BY DescripParametro"
'        CargarControlLista strSQL, cboCuponCalculo, arrCuponCalculo(), ""
'
'        If cboCuponCalculo.ListCount > 0 Then cboCuponCalculo.ListIndex = 0
'        cboCuponCalculo.Enabled = True
'    Else
'        If cboCuponCalculo.ListCount > 0 Then cboCuponCalculo.Clear
'        cboCuponCalculo.Enabled = False
'    End If
    
End Sub


Private Sub chkAmortiza_Click()

    strIndAmortizacion = ""
    If chkAmortiza.Value Then strIndAmortizacion = "X"
    
End Sub

Private Sub chkApartir_Click()

    If chkAPartir.Value = vbChecked Then
       dtpFechaInicioCupon.Visible = True
       dtpFechaInicioCupon.Value = dtpFechaEmision.Value
    Else
       dtpFechaInicioCupon.Visible = False
    End If
    
End Sub

Private Sub chk1erCupon_Click()

'    Dim dFchIni As Variant, dFechita As Variant
'    Dim n_NroDiaPer As Integer, n_MesxPrd As Integer
'
    If chk1erCupon.Value = vbChecked Then
'        With gadoComando
'            Set adoRecCONS = New ADODB.Recordset
'            .CommandType = adCmdText
'
'            .CommandText = "SELECT CNT_DIAS FROM tblFrecuencias WHERE PRD_FREC='" & aMapFrePag(cboPeriodoPago.ListIndex) & "'"
'            Set adoRecCONS = .Execute
'            n_NroDiaPer = adoRecCONS!CNT_DIAS
'            n_MesxPrd = Int(adoRecCONS!CNT_DIAS / 30)
'
'            If optDiasCupon(1).Value = True Then  'PERSONALIZAR DIAS DEL MES
'               n_NroDiaPer = CInt(txtNumDiasPeriodo.Text) * n_MesxPrd   'numero de dias por periodo
'            End If
'            adoRecCONS.Close: Set adoRecCONS = Nothing
'        End With
'
'        If chkAPartir.Value = vbChecked Then
'           dFchIni = dtpFechaInicioCupon.Text
'        Else
'           dFchIni = dtpFechaEmision.Text
'        End If
'
'        If optDiasCupon(0).Value Then                     'MESES DEL Calendario
'            dFechita = DateAdd("m", n_MesxPrd, dFchIni)
'            dtpFechaCorteInicial = DateAdd("d", -1, dFechita)
'        ElseIf optDiasCupon(1).Value Then                 'Personalizar dias del mes
'            dtpFechaCorteInicial = DateAdd("d", CInt(txtNumDiasPeriodo.Text) * n_MesxPrd, dFchIni)
'        ElseIf optDiasCupon(2).Value Then                 'Fines de mes ( Mayor al Período )
'            dtpFechaCorteInicial = DateAdd("d", -1, DateAdd("m", n_MesxPrd, dFchIni))
'            dtpFechaCorteInicial = LUltDiaMes(dtpFechaCorteInicial)
'        End If
        dtpFechaCorteInicial.Visible = True
    Else
        dtpFechaCorteInicial.Visible = False
    End If

End Sub

Private Sub chkCapitalizable_Click()

    strIndTasaCapitalizable = Valor_Caracter
    If chkCapitalizable.Value Then strIndTasaCapitalizable = Valor_Indicador
    
End Sub

Private Sub chkCuponCero_Click()

    If chkCuponCero.Value Then
    
        strIndCuponCero = "X"
        chkAmortiza.Value = vbUnchecked
        chkAmortiza.Enabled = False
        chkAjuste.Value = vbUnchecked
        chkAjuste.Enabled = False
        
        txtTasaAnual.Text = "0"
        
        cboTipoTasa.Enabled = False
        cboTasa.Enabled = False
        
        chkCapitalizable.Value = vbUnchecked
        chkCapitalizable.Enabled = False
        
        cboPeriodoPago.ListIndex = -1
        If cboPeriodoPago.ListCount > 0 Then cboPeriodoPago.ListIndex = 0
        cboPeriodoPago.Enabled = False
        
        lblNumCupones.Caption = "0"
        fraGeneraCuponera.Enabled = False
                                                                                        
        cmdGeneraCuponera(0).Enabled = False
        cmdGeneraCuponera(1).Enabled = False
        cmdGeneraCuponera(2).Enabled = False
        cmdGeneraCuponera(3).Enabled = False
    Else
        strIndCuponCero = Valor_Caracter
        chkAmortiza.Enabled = True
        chkAjuste.Enabled = True
        
        txtTasaAnual.Text = "0"
        
        cboTipoTasa.Enabled = True
        cboTasa.Enabled = True
        cboPeriodoPago.Enabled = True
        
        lblNumCupones.Caption = "0"
        fraGeneraCuponera.Enabled = True
                                                                                        
        cmdGeneraCuponera(0).Enabled = True
    End If
     
End Sub

Private Sub chkQuiebre_Click()

    strIndQuiebre = ""
    If chkQuiebre.Value Then strIndQuiebre = "X"
    
End Sub

Private Sub chkSeleccionFecha_Click()

    If chkSeleccionFecha.Value Then
        dtpFechaDesdeCriterio.Enabled = True
        dtpFechaHastaCriterio.Enabled = True
        chkEmision.Enabled = True
        chkCorte.Enabled = True
        chkVencimiento.Enabled = True
        chkPago.Enabled = True
        
        dtpFechaDesdeCriterio.Value = gdatFechaActual
        dtpFechaHastaCriterio.Value = gdatFechaActual
        chkEmision.Value = vbUnchecked
        chkCorte.Value = vbUnchecked
        chkVencimiento.Value = vbUnchecked
        chkPago.Value = vbUnchecked
    Else
        dtpFechaDesdeCriterio.Enabled = False
        dtpFechaHastaCriterio.Enabled = False
        chkEmision.Enabled = False
        chkCorte.Enabled = False
        chkVencimiento.Enabled = False
        chkPago.Enabled = False
    End If
    
End Sub

Private Sub chkSeleccionTipo_Click()

    If chkSeleccionTipo.Value Then
        cboTipoCriterio.Enabled = True
        cboClaseCriterio.Enabled = True
        cboGrupoCriterio.Enabled = True
        txtIsinCriterio.Enabled = True
        txtNemotecnicoCriterio.Enabled = True
        
        If cboTipoCriterio.ListCount > 0 Then cboTipoCriterio.ListIndex = 0
        If cboClaseCriterio.ListCount > 0 Then cboClaseCriterio.ListIndex = 0
        If cboGrupoCriterio.ListCount > 0 Then cboGrupoCriterio.ListIndex = 0
        txtIsinCriterio.Text = Valor_Caracter
        txtNemotecnicoCriterio = Valor_Caracter
    Else
        cboTipoCriterio.Enabled = False
        cboClaseCriterio.Enabled = False
        cboGrupoCriterio.Enabled = False
        txtIsinCriterio.Enabled = False
        txtNemotecnicoCriterio.Enabled = False
    End If
    
End Sub


Private Sub chkAjuste_Click()

    If chkAjuste.Value Then
        strIndTasaAjustada = Valor_Indicador
        cboTipoAjuste.Enabled = True
        cboTipoVac.Enabled = True
        cboCuponCalculo.Enabled = True
    Else
        strIndTasaAjustada = Valor_Caracter
        cboTipoAjuste.ListIndex = 0: cboTipoAjuste.Enabled = False
        cboTipoVac.ListIndex = 0: cboTipoVac.Enabled = False
        cboCuponCalculo.ListIndex = 0: cboCuponCalculo.Enabled = False
    End If
    
End Sub



Private Sub cboEmisor_Click()
      
    Dim adoRegistro     As ADODB.Recordset
    Dim intRegistro     As Integer
    
    strCodEmisor = Valor_Caracter: strCodGrupo = Valor_Caracter: strCodCiiu = Valor_Caracter
    If cboEmisor.ListIndex < 0 Then Exit Sub
    
    strCodEmisor = Left(Trim(arrEmisor(cboEmisor.ListIndex)), 8)
    strCodGrupo = Mid(Trim(arrEmisor(cboEmisor.ListIndex)), 9, 3)
    strCodCiiu = Right(Trim(arrEmisor(cboEmisor.ListIndex)), 4)
    
    With adoComm
        Set adoRegistro = New ADODB.Recordset
        
        '*** Descripción del grupo económico ***
        .CommandText = "SELECT DescripGrupo FROM GrupoEconomico WHERE CodGrupo='" & strCodGrupo & "'"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            lblGrupoEconomico.Caption = Trim(adoRegistro("DescripGrupo"))
        End If
        adoRegistro.Close
        
        '*** Descripción del ciiu ***
        .CommandText = "SELECT (CodCiiu + Space(1) + DescripCiiu) DescripCiiu FROM Ciiu WHERE CodCiiu='" & strCodCiiu & "'"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            lblCiiu.Caption = Trim(adoRegistro("DescripCiiu"))
        End If
        adoRegistro.Close
        
        '*** Categoría del instrumento emitido por el emisor ***
        .CommandText = "SELECT CodCategoriaRiesgo,CodRiesgoFinal,CodSubRiesgoFinal FROM EmisionInstitucionPersona " & _
            "WHERE CodEmisor='" & strCodEmisor & "' AND CodFile='" & strCodTipoInstrumento & "' AND " & _
            "CodDetalleFile='" & strCodClaseInstrumento & "'"
            
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            strCodCategoria = Trim(adoRegistro("CodCategoriaRiesgo"))
            strCodRiesgo = Trim(adoRegistro("CodRiesgoFinal"))
            strCodSubRiesgo = Trim(adoRegistro("CodSubRiesgoFinal"))
        End If
        adoRegistro.Close
        
        .CommandText = "SELECT CodParametro,ValorParametro FROM AuxiliarParametro WHERE CodTipoParametro='TIPRIE' AND CodParametro='" & strCodRiesgo & "'"
        Set adoRegistro = adoComm.Execute

        If Not adoRegistro.EOF Then
            lblRiesgo.Caption = Trim(adoRegistro("ValorParametro")) & Space(5) & strCodSubRiesgo
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
            
    End With

End Sub

Private Sub cboPeriodoPago_Click()

    strCodPeriodoPago = Valor_Caracter
    If cboPeriodoPago.ListIndex < 0 Then Exit Sub
    
    strCodPeriodoPago = Trim(arrPeriodoPago(cboPeriodoPago.ListIndex))
    
End Sub

Private Sub cmdGeneraCuponera_Click(Index As Integer)

    Dim strFechaIniCupon        As String  'Fecha de Inicio del 1er. Cupón
    Dim strFechaIniValor        As String  'Fecha de Emisión del Instrumento
    Dim strFechaFinValor        As String  'Fecha de Vencimiento del Instrumento
    Dim strFechaIniCorteCupon   As String  'Fecha de Vencimiento del 1er. Cupón

    If Not ValidaDatos() Then Exit Sub
    
    '** Generar Fechas de Corte ***
    If Index = 0 Then
        If Not ValidaFechasyTasas Then Exit Sub

        Me.MousePointer = vbHourglass
        
        If optDiasCupon(0).Value = True And chk1erCupon.Value = vbUnchecked Then
            MsgBox "Indique la Fecha de Vencimiento del Primer Cupón", vbCritical, Me.Caption
            chk1erCupon.SetFocus
            Me.MousePointer = vbDefault
            Exit Sub
        End If

        strFechaIniCupon = Convertyyyymmdd(dtpFechaInicioCupon.Value)
        strFechaIniValor = Convertyyyymmdd(dtpFechaEmision.Value)
        strFechaFinValor = Convertyyyymmdd(dtpFechaVencimiento.Value)
        strFechaIniCorteCupon = Convertyyyymmdd(dtpFechaCorteInicial.Value)

        '*** Generación del cronograma a partir del día... ***
        If chkAPartir.Value = vbChecked Then
        
            If dtpFechaInicioCupon < dtpFechaEmision Or dtpFechaInicioCupon >= dtpFechaVencimiento Then
                MsgBox "La Fecha de Inicio del Primer Cupón no es VALIDA.", vbCritical + vbOKOnly, Me.Caption
                If dtpFechaInicioCupon.Enabled = True Then dtpFechaInicioCupon.SetFocus
                Me.MousePointer = vbDefault
                Exit Sub
            End If

            '*** Fecha emisión <= Fecha inicio primer cupón <= Fecha emisión + 1 ***
            If dtpFechaInicioCupon > DateAdd("d", 1, dtpFechaEmision.Value) Then
                MsgBox "La Fecha de Inicio del Primer Cupón no es VALIDA.", vbCritical + vbOKOnly, Me.Caption
                If dtpFechaInicioCupon.Enabled = True Then dtpFechaInicioCupon.SetFocus
                Me.MousePointer = vbDefault
                Exit Sub
            End If

            '*** Fecha de vencimiento del primer cupoón ***
            If chk1erCupon.Value = vbChecked Then
                If dtpFechaCorteInicial <= dtpFechaInicioCupon Then
                    MsgBox "La Fecha de Vencimiento del Primer Cupón no es VALIDA.", vbCritical + vbOKOnly, Me.Caption
                    dtpFechaCorteInicial.SetFocus
                    Me.MousePointer = vbDefault
                    Exit Sub
                End If
            End If
        End If
        
        '*** FECHA DE VCTO. DEL PRIMER CUPON ***
        If chk1erCupon.Value = vbChecked Then
            If dtpFechaCorteInicial <= dtpFechaEmision Or dtpFechaCorteInicial >= dtpFechaVencimiento Then
                MsgBox "La Fecha de Vencimiento del Primer Cupón no es VALIDA.", vbCritical + vbOKOnly, Me.Caption
                dtpFechaCorteInicial.SetFocus
                Me.MousePointer = vbDefault
                Exit Sub
            End If
        End If

        '*** Generación previa de las fechas de corte ***
        Call GenerarFechasCorte
        cmdGeneraCuponera(1).Enabled = True
        cmdGeneraCuponera(3).Enabled = True
        
        Me.MousePointer = vbDefault
        
    '*** Generar fechas de pago ***
    ElseIf Index = 1 Then
    
        Me.MousePointer = vbHourglass

        '*** Generación previa de las fechas de pago ***
        Call GenerarFechasPago
        cmdGeneraCuponera(2).Enabled = True
        
        Me.MousePointer = vbDefault
        
    '*** Generar los factores y montos del cronograma ***
    ElseIf Index = 2 Then
    
        Me.MousePointer = vbHourglass

        Call GenerarFactoresMontos("000", 0)

        Me.MousePointer = vbDefault

    ElseIf Index = 3 Then
            
        If strEstado = Reg_Edicion Then
            If strCodValor <> Trim(txtCodigoValor.Text) Then
                MsgBox "Se ha modificado el Código ISIN, por favor guarde la modificación.", vbExclamation, Me.Caption
                Exit Sub
            End If
        End If
        
        frmTablaDesarrollo.strCodTitulop = Trim(txtCodigoValor.Text)
        frmTablaDesarrollo.strEstadop = strEstado
        frmTablaDesarrollo.Caption = "Tabla de Desarrollo" & Space(1) & "-" & Space(1) & Trim(txtNemonico.Text)
        frmTablaDesarrollo.Show vbModal
        
    End If
        
End Sub





Private Sub cmdRegenerar_Click()

    cmdGeneraCuponera(1).Enabled = True: cmdGeneraCuponera(2).Enabled = True
    cmdGeneraCuponera(3).Enabled = True: cmdGeneraCuponera(0).Enabled = True
    
End Sub

Private Sub dtpFechaVencimiento_LostFocus()

'    Dim s_BASE As String
'
'    s_BASE = IIf(optBase(0).Value, "C", "B")
'    lblNumCupones = CALCULA_NROCUP(dtpFechaEmision, dtpFechaVencimiento, s_BASE)
'    If DateValue(dtpFechaVencimiento) < DateValue(dtpFechaEmision) Then
'        MsgBox "Fecha de Vencimiento debe ser posterior a Fecha de Emisión", vbCritical + vbOKOnly, Me.Caption
'        dtpFechaVencimiento.SetFocus
'    End If
    
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
        
    Call ValidaExisteTipoCambio(Codigo_TipoCambio_SBS, gstrFechaActual)
    
End Sub

Private Sub DarFormato()

    Dim intCont As Integer
    Dim elemento As Object
    
    For intCont = 0 To (lblDescrip.Count - 1)
        Call FormatoEtiqueta(lblDescrip(intCont), vbLeftJustify)
    Next
    
    For Each elemento In Me.Controls
    
        If TypeOf elemento Is TDBGrid Then
            Call FormatoGrilla(elemento)
        End If
    
    Next
        
End Sub
Public Sub Buscar()
        
    Dim strSQL As String, strAND As String
    
    Set adoConsulta = New ADODB.Recordset

    strAND = Valor_Caracter

    strSQL = "SELECT CodTitulo,Nemotecnico,DescripTitulo," & _
        "FechaEmision = CASE Convert(char(8),FechaEmision,112) WHEN '19000101' THEN NULL ELSE FechaEmision END," & _
        "FechaVencimiento = CASE Convert(char(8),FechaVencimiento,112) WHEN '19000101' THEN NULL ELSE FechaVencimiento END " & _
        "FROM InstrumentoInversion "
        
    If tabCriterio.Tab = 0 Then
        If chkSeleccionTipo.Value Then
            strSQL = strSQL & " WHERE "
            If cboTipoCriterio.ListIndex > 0 Then
                strSQL = strSQL & "CodFile='" & strCodTipoCriterio & "'"
                strAND = " AND "
            End If
            
            If cboClaseCriterio.ListIndex > 0 Then
                strSQL = strSQL & strAND & "CodDetalleFile='" & strCodClaseCriterio & "'"
                strAND = " AND "
            End If
            
            If cboGrupoCriterio.ListIndex > 0 Then
                strSQL = strSQL & strAND & "CodSubDetalleFile='" & strCodSubClaseCriterio & "'"
                strAND = " AND "
            End If
            
            If Trim(txtIsinCriterio.Text) <> "" Then
                strSQL = strSQL & strAND & "CodTitulo LIKE '%" & UCase(Trim(txtIsinCriterio.Text)) & "%'"
                strAND = " AND "
            End If
            
            If Trim(txtNemotecnicoCriterio.Text) <> "" Then
                strSQL = strSQL & strAND & "Nemotecnico LIKE '%" & UCase(Trim(txtNemotecnicoCriterio.Text)) & "%'"
            End If
            
            If cboTipoCriterio.ListIndex = 0 And cboClaseCriterio.ListIndex = 0 And cboGrupoCriterio.ListIndex = 0 And Trim(txtIsinCriterio.Text) = "" And Trim(txtNemotecnicoCriterio.Text) = "" Then
                MsgBox "No ha seleccionado ningún criterio...", vbCritical
                Exit Sub
            End If
        End If
    End If
    strSQL = strSQL & " ORDER BY Nemotecnico"

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

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Características del Instrumento"
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Text = "Cuponera del Instrumento"
    
    
End Sub
Private Sub CargarListas()
                      
    '*** Tipo de Instrumento ***
    strSQL = "SELECT CodFile CODIGO,DescripFile DESCRIP FROM InversionFile WHERE IndInstrumento='X' AND IndVigente='X' ORDER BY DescripFile"
    CargarControlLista strSQL, cboTipoCriterio, arrTipoCriterio(), Sel_Defecto
    CargarControlLista strSQL, cboTipoInstrumento, arrTipoInstrumento(), Sel_Defecto
    
    If cboTipoCriterio.ListCount > 0 Then cboTipoCriterio.ListIndex = 0
    
    '*** Emisor ***
    'strSql = "SELECT (CodPersona + ISNULL(CodGrupo,'') + ISNULL(CodCiiu,'')) CODIGO,DescripPersona DESCRIP FROM InstitucionPersona WHERE TipoPersona='" & Codigo_Tipo_Persona_Emisor & "' AND IndVigente='X' ORDER BY DescripPersona"
    strSQL = "SELECT CodPersona CODIGO,DescripPersona DESCRIP FROM InstitucionPersona WHERE TipoPersona='" & Codigo_Tipo_Persona_Emisor & "' AND IndVigente='X' ORDER BY DescripPersona"
    CargarControlLista strSQL, cboEmisor, arrEmisor(), Sel_Defecto
    
    
    '*** Moneda ***
    strSQL = "{ call up_ACSelDatos(2) }"
    CargarControlLista strSQL, cboMonedaEmision, arrMonedaEmision(), Sel_Defecto
    CargarControlLista strSQL, cboMonedaPago, arrMonedaPago(), Sel_Defecto
        
    '*** Bolsas ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='BOLVAL' ORDER BY CodParametro"
    CargarControlLista strSQL, cboMercado, arrMercado(), Valor_Caracter
        
    '*** Base de Cálculo ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='BASANU' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboBaseCalculo, arrBaseCalculo(), Valor_Caracter
    
    '*** Forma de Cálculo ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='FORCAL' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboFormaCalculo, arrFormaCalculo(), Valor_Caracter
    
    '*** Periodo de Pago - Capitalización ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPFRE' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboPeriodoPago, arrPeriodoPago(), Valor_Caracter
    
    '*** Tipo Ajuste ***
    strSQL = "SELECT CodTasa CODIGO,DescripTasa DESCRIP FROM TipoTasa ORDER BY DescripTasa"
    CargarControlLista strSQL, cboTipoAjuste, arrTipoAjuste(), Valor_Caracter
    
    If cboTipoAjuste.ListCount > 0 Then cboTipoAjuste.ListIndex = 0
    
    '*** Tipo Tasa ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='NATTAS' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboTipoTasa, arrTipoTasa(), Valor_Caracter

    '*** Tipo Cupón ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPTAS' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboTasa, arrTasa(), Valor_Caracter
    
    '*** Tipo Día ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPDIA' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboTipoDia, arrTipoDia(), Valor_Caracter
    
End Sub
Private Sub InicializarValores()

    '*** Valores Iniciales ***
    strEstado = Reg_Defecto
    tabTitulos.Tab = 0
    tabTitulos.TabEnabled(1) = False
    tabTitulos.TabEnabled(2) = False
    tabCriterio.Tab = 0
    chkSeleccionTipo.Value = vbChecked
    chkSeleccionTipo.Value = vbUnchecked
    
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgConsulta.Columns(0).Width = tdgConsulta.Width * 0.01 * 15
    tdgConsulta.Columns(1).Width = tdgConsulta.Width * 0.01 * 15
    tdgConsulta.Columns(2).Width = tdgConsulta.Width * 0.01 * 38
    tdgConsulta.Columns(3).Width = tdgConsulta.Width * 0.01 * 12
    
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion2.FormularioActivo = Me
    Set cmdOpcion.FormularioActivo = Me
    
End Sub
Private Sub Form_Unload(Cancel As Integer)

     Set frmTitulosInversion = Nothing
     Call OcultarReportes
     frmMainMdi.stbMdi.Panels(3).Text = "Acción..."
     
End Sub



Private Sub optDiasCupon_Click(Index As Integer)

    txtNumDiasPeriodo.Visible = False
    Select Case Index
        Case 0
            txtNumDiasPeriodo.Visible = False
            updDiasPeriodo.Visible = False
        Case 1
            txtNumDiasPeriodo.Text = 30
            txtNumDiasPeriodo.Visible = True
            updDiasPeriodo.Visible = True
        Case 2
            txtNumDiasPeriodo.Visible = False
            updDiasPeriodo.Visible = False
    End Select
    
    If dtpFechaCorteInicial.Visible Then
        Call chk1erCupon_Click
    End If
    
End Sub

Private Sub optTipoCodigo_Click(Index As Integer)

    If optTipoCodigo(1).Value Then
        txtCodigoValor.Enabled = False
    Else
        txtCodigoValor.Enabled = True
    End If
    
End Sub

Private Sub tabCriterio_Click(PreviousTab As Integer)

    Select Case tabCriterio.Tab
        Case 0
            chkSeleccionTipo.Value = vbChecked
            chkSeleccionTipo.Value = vbUnchecked
            
        Case 1
            chkSeleccionFecha.Value = vbChecked
            chkSeleccionFecha.Value = vbUnchecked
            
    End Select
    
End Sub

Private Sub tabTitulos_Click(PreviousTab As Integer)

    Select Case tabTitulos.Tab
        Case 1, 2
            If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vModify)
            If strEstado = Reg_Defecto Then tabTitulos.Tab = 0
    End Select
    
End Sub

Private Sub txtCodigoValor_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
End Sub

Private Sub txtDescripValor_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
End Sub

Private Sub txtNemonico_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
End Sub

Private Sub txtNumDiasPago_Change()

    Call FormatoCajaTexto(txtNumDiasPago, 0)
    
End Sub

Private Sub txtNumDiasPago_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "N")
    
End Sub

Private Sub txtNumDiasPeriodo_Change()

    Call FormatoCajaTexto(txtNumDiasPeriodo, 0)
    
End Sub

Private Sub txtNumDiasPeriodo_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "N")
    
End Sub


Private Sub txtTasaAnual_Change()

    Call FormatoCajaTexto(txtTasaAnual, Decimales_Tasa)
    
End Sub


Private Sub txtTasaAnual_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtTasaAnual, Decimales_Tasa)
    
End Sub


Private Sub txtTasaAnual_LostFocus()

'    n_CurTas = CDbl(txtTasaAnual)
    
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

Private Sub txtValorNominal_Change()

    Call FormatoCajaTexto(txtValorNominal, Decimales_Monto)
    
End Sub

Private Sub txtValorNominal_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtValorNominal, Decimales_Monto)
    
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
