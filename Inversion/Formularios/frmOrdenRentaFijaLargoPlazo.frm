VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{830D5A06-9B70-4F8C-98B6-7A19C4E7760B}#1.0#0"; "TAMControls.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmOrdenRentaFijaLargoPlazo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ordenes - Renta Fija Largo Plazo"
   ClientHeight    =   9105
   ClientLeft      =   1065
   ClientTop       =   1725
   ClientWidth     =   14715
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
   Icon            =   "frmOrdenRentaFijaLargoPlazo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9105
   ScaleWidth      =   14715
   Begin TAMControls2.ucBotonEdicion2 cmdOpcion 
      Height          =   735
      Left            =   600
      TabIndex        =   193
      Top             =   8340
      Width           =   4200
      _ExtentX        =   7408
      _ExtentY        =   1296
      Buttons         =   3
      Caption0        =   "&Nuevo"
      Tag0            =   "0"
      ToolTipText0    =   "Nuevo"
      Caption1        =   "&Eliminar"
      Tag1            =   "4"
      ToolTipText1    =   "Eliminar"
      Caption2        =   "&Buscar"
      Tag2            =   "5"
      ToolTipText2    =   "Buscar"
      UserControlWidth=   4200
   End
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   12780
      TabIndex        =   191
      Top             =   8340
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
      Caption0        =   "&Salir"
      Tag0            =   "9"
      ToolTipText0    =   "Salir"
      UserControlWidth=   1200
   End
   Begin TabDlg.SSTab tabReporte 
      Height          =   8325
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14685
      _ExtentX        =   25903
      _ExtentY        =   14684
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Lista"
      TabPicture(0)   =   "frmOrdenRentaFijaLargoPlazo.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tdgConsulta"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraCriterio"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Datos Orden Inversión"
      TabPicture(1)   =   "frmOrdenRentaFijaLargoPlazo.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblDescrip(50)"
      Tab(1).Control(1)=   "txtObservacion"
      Tab(1).Control(2)=   "fraResumen"
      Tab(1).Control(3)=   "fraDatosOrden"
      Tab(1).Control(4)=   "fraDatosBasicos"
      Tab(1).Control(5)=   "cmdAccion"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Negociación"
      TabPicture(2)   =   "frmOrdenRentaFijaLargoPlazo.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraPosicion"
      Tab(2).Control(1)=   "fraDatosNegociacion"
      Tab(2).Control(2)=   "fraComisionMontoFL2"
      Tab(2).Control(3)=   "fraComisionMontoFL1"
      Tab(2).ControlCount=   4
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   -63720
         TabIndex        =   192
         Top             =   7320
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
         Height          =   2055
         Left            =   360
         TabIndex        =   162
         Top             =   600
         Width           =   13935
         Begin VB.CommandButton cmdExportarExcel 
            Caption         =   "Excel"
            Height          =   735
            Left            =   10560
            Picture         =   "frmOrdenRentaFijaLargoPlazo.frx":035E
            Style           =   1  'Graphical
            TabIndex        =   190
            Top             =   1200
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
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   166
            Top             =   1200
            Width           =   5145
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
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   165
            Top             =   780
            Width           =   5145
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
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   164
            Top             =   360
            Width           =   5145
         End
         Begin VB.CommandButton cmdEnviar 
            Caption         =   "En&viar"
            Height          =   735
            Left            =   12200
            Picture         =   "frmOrdenRentaFijaLargoPlazo.frx":0966
            Style           =   1  'Graphical
            TabIndex        =   163
            ToolTipText     =   "Enviar a BackOffice"
            Top             =   1200
            Width           =   1200
         End
         Begin MSComCtl2.DTPicker dtpFechaOrdenDesde 
            Height          =   285
            Left            =   9600
            TabIndex        =   167
            Top             =   360
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
            Format          =   293404673
            CurrentDate     =   38785
         End
         Begin MSComCtl2.DTPicker dtpFechaOrdenHasta 
            Height          =   285
            Left            =   11955
            TabIndex        =   168
            Top             =   360
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
            Format          =   293404673
            CurrentDate     =   38785
         End
         Begin MSComCtl2.DTPicker dtpFechaLiquidacionDesde 
            Height          =   285
            Left            =   9600
            TabIndex        =   169
            Top             =   780
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
            Format          =   293404673
            CurrentDate     =   38785
         End
         Begin MSComCtl2.DTPicker dtpFechaLiquidacionHasta 
            Height          =   285
            Left            =   11955
            TabIndex        =   170
            Top             =   780
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
            Format          =   293404673
            CurrentDate     =   38785
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Estado"
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
            Left            =   480
            TabIndex        =   179
            Top             =   1220
            Width           =   495
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Instrumento"
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
            Left            =   480
            TabIndex        =   178
            Top             =   800
            Width           =   825
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
            Index           =   21
            Left            =   11280
            TabIndex        =   177
            Top             =   380
            Width           =   420
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
            Index           =   20
            Left            =   8880
            TabIndex        =   176
            Top             =   380
            Width           =   465
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fondo"
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
            Left            =   480
            TabIndex        =   175
            Top             =   380
            Width           =   450
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Orden"
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
            Index           =   43
            Left            =   7200
            TabIndex        =   174
            Top             =   380
            Width           =   930
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Liquidación"
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
            Index           =   44
            Left            =   7200
            TabIndex        =   173
            Top             =   800
            Width           =   1305
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
            Index           =   45
            Left            =   8880
            TabIndex        =   172
            Top             =   800
            Width           =   465
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
            Index           =   46
            Left            =   11280
            TabIndex        =   171
            Top             =   800
            Width           =   420
         End
      End
      Begin VB.Frame fraDatosBasicos 
         Caption         =   "Datos Básicos"
         Height          =   2460
         Left            =   -74640
         TabIndex        =   141
         Top             =   465
         Width           =   13935
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
            Left            =   1755
            Style           =   2  'Dropdown List
            TabIndex        =   151
            Top             =   360
            Width           =   4185
         End
         Begin VB.ComboBox cboAgente 
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
            Left            =   1755
            Style           =   2  'Dropdown List
            TabIndex        =   150
            Top             =   737
            Width           =   4185
         End
         Begin VB.ComboBox cboTitulo 
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
            Left            =   9315
            Style           =   2  'Dropdown List
            TabIndex        =   149
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
            Left            =   1755
            Style           =   2  'Dropdown List
            TabIndex        =   148
            Top             =   1114
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
            Left            =   1755
            Style           =   2  'Dropdown List
            TabIndex        =   147
            Top             =   1491
            Width           =   4185
         End
         Begin VB.ComboBox cboTipoOrden 
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
            Left            =   1755
            Style           =   2  'Dropdown List
            TabIndex        =   146
            Top             =   1870
            Width           =   4185
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
            Left            =   9315
            Style           =   2  'Dropdown List
            TabIndex        =   145
            Top             =   1080
            Width           =   4185
         End
         Begin VB.ComboBox cboOperacion 
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
            Left            =   9315
            Style           =   2  'Dropdown List
            TabIndex        =   144
            Top             =   737
            Width           =   4185
         End
         Begin VB.ComboBox cboOrigen 
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
            Left            =   9315
            Style           =   2  'Dropdown List
            TabIndex        =   143
            Top             =   1870
            Width           =   4185
         End
         Begin VB.ComboBox cboConceptoCosto 
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
            Left            =   9315
            Style           =   2  'Dropdown List
            TabIndex        =   142
            Top             =   1491
            Width           =   4185
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fondo"
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
            TabIndex        =   161
            Top             =   380
            Width           =   450
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Agente"
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
            Left            =   360
            TabIndex        =   160
            Top             =   757
            Width           =   510
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Orden de"
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
            TabIndex        =   159
            Top             =   1890
            Width           =   660
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
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   29
            Left            =   360
            TabIndex        =   158
            Top             =   1511
            Width           =   390
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Instrumento"
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
            Height          =   195
            Index           =   13
            Left            =   360
            TabIndex        =   157
            Top             =   1134
            Width           =   825
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Título"
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
            Height          =   195
            Index           =   15
            Left            =   6840
            TabIndex        =   156
            Top             =   375
            Width           =   420
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Mecanismo Negociación"
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
            Height          =   195
            Index           =   31
            Left            =   6840
            TabIndex        =   155
            Top             =   1140
            Width           =   1755
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Liquidación Operación"
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
            Height          =   195
            Index           =   28
            Left            =   6840
            TabIndex        =   154
            Top             =   750
            Width           =   1590
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Mercado Negociación"
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
            Height          =   195
            Index           =   76
            Left            =   6840
            TabIndex        =   153
            Top             =   1890
            Width           =   1575
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Concepto Costo Neg."
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
            Height          =   195
            Index           =   11
            Left            =   6840
            TabIndex        =   152
            Top             =   1511
            Width           =   1530
         End
      End
      Begin VB.Frame fraDatosOrden 
         Caption         =   "Datos de la Orden"
         Height          =   885
         Left            =   -74640
         TabIndex        =   134
         Top             =   2940
         Width           =   13935
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
            Left            =   8835
            TabIndex        =   135
            Top             =   360
            Width           =   4605
         End
         Begin MSComCtl2.DTPicker dtpFechaOrden 
            Height          =   285
            Left            =   1635
            TabIndex        =   136
            Top             =   360
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   503
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
            Format          =   293404673
            CurrentDate     =   38776
         End
         Begin MSComCtl2.DTPicker dtpFechaLiquidacion 
            Height          =   285
            Left            =   5040
            TabIndex        =   137
            Top             =   360
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
            Format          =   293404673
            CurrentDate     =   38776
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Orden"
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
            Left            =   270
            TabIndex        =   140
            Top             =   380
            Width           =   930
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Liquidacion"
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
            Left            =   3510
            TabIndex        =   139
            Top             =   380
            Width           =   1305
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Descripción"
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
            Left            =   6960
            TabIndex        =   138
            Top             =   380
            Width           =   840
         End
      End
      Begin VB.Frame fraComisionMontoFL1 
         Caption         =   "Comisiones y Montos - Contado (FL1)"
         Height          =   4095
         Left            =   -74640
         TabIndex        =   104
         Top             =   3450
         Width           =   7575
         Begin VB.TextBox txtPorcenAgente 
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
            Left            =   2625
            MaxLength       =   45
            TabIndex        =   114
            Top             =   980
            Width           =   1905
         End
         Begin VB.TextBox txtComisionConasev 
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
            Left            =   4590
            MaxLength       =   45
            TabIndex        =   113
            Top             =   2292
            Width           =   2025
         End
         Begin VB.TextBox txtComisionFondo 
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
            Left            =   4590
            MaxLength       =   45
            TabIndex        =   112
            Top             =   1964
            Width           =   2025
         End
         Begin VB.TextBox txtComisionCavali 
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
            Left            =   4590
            MaxLength       =   45
            TabIndex        =   111
            Top             =   1636
            Width           =   2025
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
            Left            =   4590
            MaxLength       =   45
            TabIndex        =   110
            Top             =   1308
            Width           =   2025
         End
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
            Left            =   4590
            MaxLength       =   45
            TabIndex        =   109
            Top             =   980
            Width           =   2025
         End
         Begin VB.CheckBox chkAplicar 
            Caption         =   "Aplicar"
            Height          =   255
            Index           =   0
            Left            =   390
            TabIndex        =   108
            ToolTipText     =   "Aplicar Costos Negociación"
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox txtTasaMensual 
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
            Left            =   4590
            MaxLength       =   45
            TabIndex        =   107
            Top             =   4240
            Visible         =   0   'False
            Width           =   2025
         End
         Begin VB.TextBox txtVacCorrido 
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
            Left            =   4590
            MaxLength       =   45
            TabIndex        =   106
            Top             =   3420
            Width           =   2025
         End
         Begin VB.CheckBox chkAjustePrecio 
            Caption         =   "Reajuste x Precio"
            Height          =   255
            Index           =   0
            Left            =   390
            TabIndex        =   105
            ToolTipText     =   "Aplicar Costos Negociación"
            Top             =   3440
            Width           =   1935
         End
         Begin TAMControls.TAMTextBox txtInteresCorrido 
            Height          =   315
            Index           =   0
            Left            =   4590
            TabIndex        =   186
            Top             =   3060
            Width           =   2025
            _ExtentX        =   3572
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
            Container       =   "frmOrdenRentaFijaLargoPlazo.frx":0EC1
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   16776960
            AceptaNegativos =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   12
            MaximoValor     =   999999999
         End
         Begin TAMControls.TAMTextBox txtSubTotal 
            Height          =   315
            Index           =   0
            Left            =   4590
            TabIndex        =   188
            Top             =   510
            Width           =   2025
            _ExtentX        =   3572
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
            Container       =   "frmOrdenRentaFijaLargoPlazo.frx":0EDD
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   16776960
            AceptaNegativos =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   12
            MaximoValor     =   999999999
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisión SAB"
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
            Height          =   195
            Index           =   24
            Left            =   390
            TabIndex        =   133
            Top             =   1000
            Width           =   990
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisión BVL"
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
            Height          =   195
            Index           =   26
            Left            =   390
            TabIndex        =   132
            Top             =   1328
            Width           =   975
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisión Cavali"
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
            Height          =   195
            Index           =   27
            Left            =   390
            TabIndex        =   131
            Top             =   1656
            Width           =   1110
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisión Fondo Garantía"
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
            Height          =   195
            Index           =   32
            Left            =   390
            TabIndex        =   130
            Top             =   1984
            Width           =   1800
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisión Conasev"
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
            Height          =   195
            Index           =   33
            Left            =   390
            TabIndex        =   129
            Top             =   2302
            Width           =   1305
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "IGV"
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
            Height          =   195
            Index           =   34
            Left            =   390
            TabIndex        =   128
            Top             =   2640
            Width           =   270
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "SubTotal"
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
            Height          =   195
            Index           =   25
            Left            =   2640
            TabIndex        =   127
            Top             =   580
            Width           =   645
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Interés Corrido"
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
            Height          =   195
            Index           =   35
            Left            =   2640
            TabIndex        =   126
            Top             =   3120
            Width           =   1020
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto Total"
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
            Height          =   195
            Index           =   36
            Left            =   2610
            TabIndex        =   125
            Top             =   3765
            Width           =   855
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   4665
            TabIndex        =   124
            Top             =   240
            Width           =   1845
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   0
            X1              =   2580
            X2              =   6810
            Y1              =   885
            Y2              =   885
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
            Left            =   4590
            TabIndex        =   123
            Tag             =   "0.00"
            ToolTipText     =   "Monto de Comisión IGV"
            Top             =   2625
            Width           =   2025
         End
         Begin VB.Label lblPorcenIgv 
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
            Left            =   2625
            TabIndex        =   122
            Tag             =   "0.00"
            ToolTipText     =   "Porcentaje de Comisión IGV"
            Top             =   2625
            Width           =   1905
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   1
            X1              =   360
            X2              =   6780
            Y1              =   3000
            Y2              =   3000
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
            Left            =   4590
            TabIndex        =   121
            Tag             =   "0.00"
            ToolTipText     =   "Monto Total de la Orden"
            Top             =   3735
            Width           =   2025
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
            Left            =   2625
            TabIndex        =   120
            Tag             =   "0"
            ToolTipText     =   "Porcentaje de Comisión BVL"
            Top             =   1305
            Width           =   1905
         End
         Begin VB.Label lblPorcenCavali 
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
            Left            =   2625
            TabIndex        =   119
            Tag             =   "0"
            ToolTipText     =   "Porcentaje de Comisión Cavali"
            Top             =   1635
            Width           =   1905
         End
         Begin VB.Label lblPorcenFondo 
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
            Left            =   2625
            TabIndex        =   118
            Tag             =   "0"
            ToolTipText     =   "Porcentaje de Comisión Fondo Liquidación"
            Top             =   1965
            Width           =   1905
         End
         Begin VB.Label lblPorcenConasev 
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
            Left            =   2625
            TabIndex        =   117
            Tag             =   "0"
            ToolTipText     =   "Porcentaje de Comisión Conasev"
            Top             =   2295
            Width           =   1905
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   2
            X1              =   360
            X2              =   6780
            Y1              =   4125
            Y2              =   4125
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tasa Mensual (%)"
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
            Height          =   195
            Index           =   37
            Left            =   2640
            TabIndex        =   116
            Top             =   4260
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Vac Corrido"
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
            Height          =   195
            Index           =   72
            Left            =   2640
            TabIndex        =   115
            Top             =   3435
            Width           =   825
         End
      End
      Begin VB.Frame fraComisionMontoFL2 
         Caption         =   "Comisiones y Montos - Plazo (FL2)"
         Height          =   195
         Left            =   -67080
         TabIndex        =   69
         Top             =   9000
         Width           =   7605
         Begin VB.CheckBox chkAplicar 
            Caption         =   "Aplicar"
            Height          =   255
            Index           =   1
            Left            =   390
            TabIndex        =   79
            ToolTipText     =   "Aplicar Costos Negociación"
            Top             =   600
            Width           =   975
         End
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
            Index           =   1
            Left            =   4290
            MaxLength       =   45
            TabIndex        =   78
            Top             =   980
            Width           =   2025
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
            Index           =   1
            Left            =   4290
            MaxLength       =   45
            TabIndex        =   77
            Top             =   1308
            Width           =   2025
         End
         Begin VB.TextBox txtComisionCavali 
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
            Index           =   1
            Left            =   4290
            MaxLength       =   45
            TabIndex        =   76
            Top             =   1636
            Width           =   2025
         End
         Begin VB.TextBox txtComisionFondo 
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
            Index           =   1
            Left            =   4290
            MaxLength       =   45
            TabIndex        =   75
            Top             =   1964
            Width           =   2025
         End
         Begin VB.TextBox txtComisionConasev 
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
            Index           =   1
            Left            =   4290
            MaxLength       =   45
            TabIndex        =   74
            Top             =   2292
            Width           =   2025
         End
         Begin VB.TextBox txtPorcenAgente 
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
            Index           =   1
            Left            =   2625
            MaxLength       =   45
            TabIndex        =   73
            Top             =   980
            Width           =   1340
         End
         Begin VB.CommandButton cmdCalculo 
            Caption         =   "#"
            Height          =   285
            Left            =   360
            TabIndex        =   72
            ToolTipText     =   "Calcular TIRs de la orden"
            Top             =   4240
            Width           =   375
         End
         Begin VB.TextBox txtVacCorrido 
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
            Index           =   1
            Left            =   4290
            MaxLength       =   45
            TabIndex        =   71
            Top             =   3420
            Width           =   2025
         End
         Begin VB.CheckBox chkAjustePrecio 
            Caption         =   "Reajuste x Precio"
            Height          =   255
            Index           =   1
            Left            =   390
            TabIndex        =   70
            ToolTipText     =   "Aplicar Costos Negociación"
            Top             =   3440
            Width           =   1935
         End
         Begin TAMControls.TAMTextBox txtPrecio 
            Height          =   315
            Index           =   1
            Left            =   1470
            TabIndex        =   80
            Top             =   180
            Width           =   2235
            _ExtentX        =   3942
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
            Container       =   "frmOrdenRentaFijaLargoPlazo.frx":0EF9
            Text            =   "0.00000000"
            Decimales       =   8
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            AceptaNegativos =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   8
            MaximoValor     =   999999999
         End
         Begin TAMControls.TAMTextBox txtInteresCorrido 
            Height          =   315
            Index           =   1
            Left            =   4140
            TabIndex        =   187
            Top             =   3120
            Width           =   2235
            _ExtentX        =   3942
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
            Container       =   "frmOrdenRentaFijaLargoPlazo.frx":0F15
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            AceptaNegativos =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   12
            MaximoValor     =   999999999
         End
         Begin TAMControls.TAMTextBox txtSubTotal 
            Height          =   315
            Index           =   1
            Left            =   4290
            TabIndex        =   189
            Top             =   540
            Width           =   2025
            _ExtentX        =   3572
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
            Container       =   "frmOrdenRentaFijaLargoPlazo.frx":0F31
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            AceptaNegativos =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   12
            MaximoValor     =   999999999
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   5
            X1              =   360
            X2              =   6300
            Y1              =   4120
            Y2              =   4120
         End
         Begin VB.Label lblPorcenConasev 
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
            Left            =   2625
            TabIndex        =   103
            Tag             =   "0"
            ToolTipText     =   "Porcentaje de Comisión Conasev"
            Top             =   2292
            Width           =   1335
         End
         Begin VB.Label lblPorcenFondo 
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
            Left            =   2625
            TabIndex        =   102
            Tag             =   "0"
            ToolTipText     =   "Porcentaje de Comisión Fondo Liquidación"
            Top             =   1964
            Width           =   1335
         End
         Begin VB.Label lblPorcenCavali 
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
            Left            =   2625
            TabIndex        =   101
            Tag             =   "0"
            ToolTipText     =   "Porcentaje de Comisión Cavali"
            Top             =   1636
            Width           =   1335
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
            Index           =   1
            Left            =   2625
            TabIndex        =   100
            Tag             =   "0"
            ToolTipText     =   "Porcentaje de Comisión BVL"
            Top             =   1308
            Width           =   1335
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
            Index           =   1
            Left            =   4290
            TabIndex        =   99
            Tag             =   "0.00"
            ToolTipText     =   "Monto Total de la Orden"
            Top             =   3740
            Width           =   2025
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   4
            X1              =   360
            X2              =   6300
            Y1              =   3000
            Y2              =   3000
         End
         Begin VB.Label lblPorcenIgv 
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
            Left            =   2625
            TabIndex        =   98
            Tag             =   "0.00"
            ToolTipText     =   "Porcentaje de Comisión IGV"
            Top             =   2620
            Width           =   1335
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
            Index           =   1
            Left            =   4290
            TabIndex        =   97
            Tag             =   "0.00"
            ToolTipText     =   "Monto de Comisión IGV"
            Top             =   2620
            Width           =   2025
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   3
            X1              =   2580
            X2              =   6300
            Y1              =   880
            Y2              =   880
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   4365
            TabIndex        =   96
            Top             =   240
            Width           =   1845
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto Total"
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
            Height          =   195
            Index           =   38
            Left            =   2640
            TabIndex        =   95
            Top             =   3760
            Width           =   855
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Interés Corrido"
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
            Height          =   195
            Index           =   39
            Left            =   2640
            TabIndex        =   94
            Top             =   3120
            Width           =   1020
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "SubTotal"
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
            Height          =   195
            Index           =   47
            Left            =   2640
            TabIndex        =   93
            Top             =   580
            Width           =   645
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "IGV"
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
            Height          =   195
            Index           =   42
            Left            =   390
            TabIndex        =   92
            Top             =   2640
            Width           =   270
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisión Conasev"
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
            Height          =   195
            Index           =   41
            Left            =   390
            TabIndex        =   91
            Top             =   2302
            Width           =   1305
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisión Fondo Garantía"
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
            Height          =   195
            Index           =   40
            Left            =   390
            TabIndex        =   90
            Top             =   1984
            Width           =   1800
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisión Cavali"
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
            Height          =   195
            Index           =   30
            Left            =   390
            TabIndex        =   89
            Top             =   1656
            Width           =   1110
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisión BVL"
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
            Height          =   195
            Index           =   18
            Left            =   390
            TabIndex        =   88
            Top             =   1328
            Width           =   975
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisión SAB"
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
            Height          =   195
            Index           =   17
            Left            =   390
            TabIndex        =   87
            Top             =   1000
            Width           =   990
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Precio (%)"
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
            Height          =   195
            Index           =   14
            Left            =   390
            TabIndex        =   86
            Top             =   255
            Width           =   705
         End
         Begin VB.Label lblTirBruta 
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
            Left            =   2160
            TabIndex        =   85
            Tag             =   "0.00"
            Top             =   4240
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
            Left            =   4785
            TabIndex        =   84
            Tag             =   "0.00"
            Top             =   4240
            Width           =   1335
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "TIR Bruta"
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
            Height          =   195
            Index           =   48
            Left            =   1200
            TabIndex        =   83
            Top             =   4260
            Width           =   690
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "TIR Neta"
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
            Height          =   195
            Index           =   49
            Left            =   3720
            TabIndex        =   82
            Top             =   4260
            Width           =   660
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Vac Corrido"
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
            Height          =   195
            Index           =   73
            Left            =   2640
            TabIndex        =   81
            Top             =   3440
            Width           =   825
         End
      End
      Begin VB.Frame fraDatosNegociacion 
         Caption         =   "Negociación"
         Height          =   2715
         Left            =   -74640
         TabIndex        =   50
         Top             =   630
         Width           =   9135
         Begin VB.ComboBox cboCalculo 
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
            Left            =   5940
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   3810
            Width           =   2925
         End
         Begin VB.TextBox txtTirBrutaBase365 
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
            Left            =   5520
            MaxLength       =   45
            TabIndex        =   52
            Top             =   3930
            Width           =   1580
         End
         Begin VB.TextBox txtTirNetaBase365 
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
            Left            =   3690
            MaxLength       =   45
            TabIndex        =   51
            Top             =   3990
            Width           =   1580
         End
         Begin TAMControls.TAMTextBox txtCantidad 
            Height          =   315
            Left            =   2010
            TabIndex        =   54
            Top             =   840
            Width           =   2235
            _ExtentX        =   3942
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
            Container       =   "frmOrdenRentaFijaLargoPlazo.frx":0F4D
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            AceptaNegativos =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   2
            MaximoValor     =   999999999
         End
         Begin TAMControls.TAMTextBox txtMontoNominal 
            Height          =   315
            Left            =   2010
            TabIndex        =   55
            Top             =   450
            Width           =   2235
            _ExtentX        =   3942
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
            Container       =   "frmOrdenRentaFijaLargoPlazo.frx":0F69
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            AceptaNegativos =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   12
            MaximoValor     =   999999999
         End
         Begin TAMControls.TAMTextBox txtTirBrutaBaseBono 
            Height          =   315
            Left            =   2010
            TabIndex        =   56
            Top             =   1200
            Width           =   2235
            _ExtentX        =   3942
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
            Container       =   "frmOrdenRentaFijaLargoPlazo.frx":0F85
            Text            =   "0.000000000000"
            Decimales       =   12
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            AceptaNegativos =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   12
            MaximoValor     =   999999999
         End
         Begin TAMControls.TAMTextBox txtTirNetaBaseBono 
            Height          =   315
            Left            =   2010
            TabIndex        =   57
            Top             =   1980
            Width           =   2235
            _ExtentX        =   3942
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
            Container       =   "frmOrdenRentaFijaLargoPlazo.frx":0FA1
            Text            =   "0.00000000"
            Decimales       =   8
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   16776960
            AceptaNegativos =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   8
            MaximoValor     =   999999999
         End
         Begin TAMControls.TAMTextBox txtPrecio 
            Height          =   315
            Index           =   0
            Left            =   6510
            TabIndex        =   58
            Top             =   1230
            Width           =   2235
            _ExtentX        =   3942
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
            Container       =   "frmOrdenRentaFijaLargoPlazo.frx":0FBD
            Text            =   "0.000000000000"
            Decimales       =   12
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   16776960
            AceptaNegativos =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   12
            MaximoValor     =   999999999
         End
         Begin TAMControls.TAMTextBox txtValorActual 
            Height          =   315
            Left            =   6510
            TabIndex        =   59
            Top             =   450
            Width           =   2235
            _ExtentX        =   3942
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
            Container       =   "frmOrdenRentaFijaLargoPlazo.frx":0FD9
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   16776960
            AceptaNegativos =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   12
            MaximoValor     =   999999999
         End
         Begin TAMControls.TAMTextBox txtPrecioSucio 
            Height          =   315
            Left            =   6510
            TabIndex        =   182
            Top             =   840
            Width           =   2235
            _ExtentX        =   3942
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
            Container       =   "frmOrdenRentaFijaLargoPlazo.frx":0FF5
            Text            =   "0.000000000000"
            Decimales       =   12
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   16776960
            AceptaNegativos =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   12
            MaximoValor     =   999999999
         End
         Begin TAMControls.TAMTextBox txtTirBrutaLimpia 
            Height          =   315
            Left            =   2010
            TabIndex        =   184
            Top             =   1590
            Width           =   2235
            _ExtentX        =   3942
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
            Container       =   "frmOrdenRentaFijaLargoPlazo.frx":1011
            Text            =   "0.000000000000"
            Decimales       =   12
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   16776960
            AceptaNegativos =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   12
            MaximoValor     =   999999999
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tir Bruta Limpia (%)"
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
            Index           =   84
            Left            =   360
            TabIndex        =   185
            Top             =   1620
            Width           =   1350
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Precio Sucio (%)"
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
            Height          =   195
            Index           =   83
            Left            =   4860
            TabIndex        =   183
            Top             =   870
            Width           =   1155
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad Titulos"
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
            TabIndex        =   68
            Top             =   840
            Width           =   1140
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Calcular"
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
            Left            =   4920
            TabIndex        =   67
            Top             =   3780
            Width           =   570
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tir Bruta Sucia (%)"
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
            Index           =   57
            Left            =   360
            TabIndex        =   66
            Top             =   1230
            Width           =   1305
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tir Neta (%)"
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
            Index           =   58
            Left            =   360
            TabIndex        =   65
            Top             =   2010
            Width           =   825
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Base Bono"
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
            Index           =   59
            Left            =   6690
            TabIndex        =   64
            Top             =   3810
            Width           =   780
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Base 365"
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
            Index           =   60
            Left            =   5760
            TabIndex        =   63
            Top             =   3960
            Width           =   675
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000015&
            X1              =   4560
            X2              =   4560
            Y1              =   390
            Y2              =   2310
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Precio Limpio (%)"
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
            Height          =   195
            Index           =   9
            Left            =   4860
            TabIndex        =   62
            Top             =   1260
            Width           =   1200
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto Nominal"
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
            Height          =   195
            Index           =   81
            Left            =   360
            TabIndex        =   61
            Top             =   480
            Width           =   1065
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Valor Actual"
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
            Index           =   82
            Left            =   4860
            TabIndex        =   60
            Top             =   480
            Width           =   855
         End
      End
      Begin VB.Frame fraPosicion 
         Caption         =   "Datos Posición"
         Height          =   2745
         Left            =   -65280
         TabIndex        =   35
         Top             =   630
         Width           =   4545
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
            Height          =   315
            Left            =   2190
            MaxLength       =   45
            TabIndex        =   194
            Top             =   3135
            Width           =   1875
         End
         Begin VB.Label lblIndiceFinal 
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
            Left            =   2190
            TabIndex        =   201
            Tag             =   "0.00"
            ToolTipText     =   "Fecha de Liquidación"
            Top             =   3930
            Width           =   1875
         End
         Begin VB.Label lblIndiceInicial 
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
            Left            =   2190
            TabIndex        =   200
            Tag             =   "0.00"
            ToolTipText     =   "Fecha de Liquidación"
            Top             =   3525
            Width           =   1875
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Indice Final"
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
            Index           =   79
            Left            =   870
            TabIndex        =   199
            Top             =   3975
            Width           =   810
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Indice Inicial"
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
            Index           =   78
            Left            =   870
            TabIndex        =   198
            Top             =   3585
            Width           =   885
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Liquidación"
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
            Left            =   870
            TabIndex        =   197
            Top             =   2805
            Width           =   855
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
            Height          =   315
            Left            =   2190
            TabIndex        =   196
            Tag             =   "0.00"
            ToolTipText     =   "Fecha de Liquidación"
            Top             =   2760
            Width           =   1875
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Cambio"
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
            Left            =   870
            TabIndex        =   195
            Top             =   3180
            Width           =   885
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
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   51
            Left            =   360
            TabIndex        =   49
            Top             =   320
            Width           =   975
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Próximo Cupón"
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
            Height          =   195
            Index           =   52
            Left            =   360
            TabIndex        =   48
            Top             =   1035
            Width           =   1065
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Clasificación"
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
            Height          =   195
            Index           =   53
            Left            =   360
            TabIndex        =   47
            Top             =   1335
            Width           =   885
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Base - Tasa %"
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
            Height          =   195
            Index           =   54
            Left            =   360
            TabIndex        =   46
            Top             =   1665
            Width           =   1020
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Stock Nominal"
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
            Height          =   195
            Index           =   55
            Left            =   360
            TabIndex        =   45
            Top             =   1965
            Width           =   1035
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
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
            Height          =   195
            Index           =   56
            Left            =   360
            TabIndex        =   44
            Top             =   2295
            Width           =   585
         End
         Begin VB.Label lblValorNominal 
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
            Height          =   315
            Left            =   2190
            TabIndex        =   43
            Tag             =   "0.00"
            Top             =   300
            Width           =   2025
         End
         Begin VB.Label lblFechaCupon 
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
            Left            =   2190
            TabIndex        =   42
            Tag             =   "0.00"
            Top             =   960
            Width           =   2025
         End
         Begin VB.Label lblClasificacion 
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
            Left            =   2190
            TabIndex        =   41
            Tag             =   "0.00"
            Top             =   1290
            Width           =   2025
         End
         Begin VB.Label lblBaseTasaCupon 
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
            Left            =   2190
            TabIndex        =   40
            Tag             =   "0.00"
            Top             =   1620
            Width           =   2025
         End
         Begin VB.Label lblStockNominal 
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
            Height          =   315
            Left            =   2190
            TabIndex        =   39
            Tag             =   "0.00"
            Top             =   1950
            Width           =   2025
         End
         Begin VB.Label lblMoneda 
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
            Left            =   2190
            TabIndex        =   38
            Tag             =   "0.00"
            ToolTipText     =   "Moneda del Título"
            Top             =   2280
            Width           =   2025
         End
         Begin VB.Label lblInicioFechaCupon 
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
            Left            =   2190
            TabIndex        =   37
            Tag             =   "0.00"
            Top             =   630
            Width           =   2025
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Cupón Vigente"
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
            Height          =   195
            Index           =   80
            Left            =   360
            TabIndex        =   36
            Top             =   675
            Width           =   1050
         End
      End
      Begin VB.Frame fraResumen 
         Caption         =   "Resumen Negociación"
         Height          =   3135
         Left            =   -74640
         TabIndex        =   2
         Top             =   3840
         Width           =   13935
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Precio"
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
            Left            =   5280
            TabIndex        =   34
            Top             =   1180
            Width           =   450
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Precio"
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
            Left            =   480
            TabIndex        =   33
            Top             =   1180
            Width           =   450
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "SubTotal"
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
            Index           =   61
            Left            =   480
            TabIndex        =   32
            Top             =   1540
            Width           =   645
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisiones"
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
            Index           =   63
            Left            =   480
            TabIndex        =   31
            Top             =   1900
            Width           =   795
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Intereses Corridos"
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
            Index           =   64
            Left            =   480
            TabIndex        =   30
            Top             =   2260
            Width           =   1260
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto Total"
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
            Index           =   65
            Left            =   480
            TabIndex        =   29
            Top             =   2620
            Width           =   855
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "SubTotal"
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
            Index           =   66
            Left            =   5280
            TabIndex        =   28
            Top             =   1540
            Width           =   645
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisiones"
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
            Index           =   67
            Left            =   5280
            TabIndex        =   27
            Top             =   1900
            Width           =   795
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Intereses Corridos"
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
            Index           =   68
            Left            =   5280
            TabIndex        =   26
            Top             =   2260
            Width           =   1260
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto Total"
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
            Index           =   69
            Left            =   5280
            TabIndex        =   25
            Top             =   2620
            Width           =   855
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
            Left            =   2400
            TabIndex        =   24
            Tag             =   "0.00"
            Top             =   1160
            Width           =   2025
         End
         Begin VB.Label lblSubTotalResumen 
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
            Left            =   2400
            TabIndex        =   23
            Tag             =   "0.00"
            Top             =   1520
            Width           =   2025
         End
         Begin VB.Label lblComisionesResumen 
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
            Left            =   2400
            TabIndex        =   22
            Tag             =   "0.00"
            Top             =   1880
            Width           =   2025
         End
         Begin VB.Label lblInteresesResumen 
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
            Left            =   2400
            TabIndex        =   21
            Tag             =   "0.00"
            Top             =   2240
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
            Left            =   2400
            TabIndex        =   20
            Tag             =   "0.00"
            Top             =   2600
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
            Left            =   7320
            TabIndex        =   19
            Tag             =   "0.00"
            Top             =   1160
            Width           =   2025
         End
         Begin VB.Label lblSubTotalResumen 
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
            Left            =   7320
            TabIndex        =   18
            Tag             =   "0.00"
            Top             =   1520
            Width           =   2025
         End
         Begin VB.Label lblComisionesResumen 
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
            Left            =   7320
            TabIndex        =   17
            Tag             =   "0.00"
            Top             =   1880
            Width           =   2025
         End
         Begin VB.Label lblInteresesResumen 
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
            Left            =   7320
            TabIndex        =   16
            Tag             =   "0.00"
            Top             =   2240
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
            Left            =   7320
            TabIndex        =   15
            Tag             =   "0.00"
            Top             =   2600
            Width           =   2025
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Contado"
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
            Index           =   70
            Left            =   480
            TabIndex        =   14
            Top             =   840
            Width           =   600
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Plazo"
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
            Index           =   71
            Left            =   5280
            TabIndex        =   13
            Top             =   840
            Width           =   390
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tir Bruta"
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
            Index           =   74
            Left            =   10200
            TabIndex        =   12
            Top             =   1200
            Width           =   600
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tir Neta"
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
            Index           =   75
            Left            =   10200
            TabIndex        =   11
            Top             =   1560
            Width           =   570
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
            Left            =   11400
            TabIndex        =   10
            Tag             =   "0.00"
            Top             =   1200
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
            Left            =   11400
            TabIndex        =   9
            Tag             =   "0.00"
            Top             =   1560
            Width           =   2025
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000015&
            X1              =   4800
            X2              =   4800
            Y1              =   360
            Y2              =   2880
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
            Left            =   2400
            TabIndex        =   8
            Tag             =   "0.00"
            Top             =   360
            Width           =   2025
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad Facial"
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
            Index           =   77
            Left            =   480
            TabIndex        =   7
            Top             =   375
            Width           =   1095
         End
         Begin VB.Label lblDescripMonedaResumen 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Nuevos Soles (S/.)"
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
            Height          =   195
            Index           =   0
            Left            =   2400
            TabIndex        =   6
            Top             =   840
            Width           =   2025
         End
         Begin VB.Label lblDescripMonedaResumen 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Nuevos Soles (S/.)"
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
            Height          =   195
            Index           =   1
            Left            =   7320
            TabIndex        =   5
            Top             =   840
            Width           =   1845
         End
         Begin VB.Line Line3 
            BorderColor     =   &H80000015&
            X1              =   9720
            X2              =   9720
            Y1              =   360
            Y2              =   2880
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
            Left            =   11400
            TabIndex        =   4
            Tag             =   "0.00"
            Top             =   360
            Width           =   2025
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
            Index           =   62
            Left            =   10200
            TabIndex        =   3
            Top             =   360
            Width           =   630
         End
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
         Height          =   945
         Left            =   -72840
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   7140
         Width           =   8460
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmOrdenRentaFijaLargoPlazo.frx":102D
         Height          =   5085
         Left            =   360
         OleObjectBlob   =   "frmOrdenRentaFijaLargoPlazo.frx":1047
         TabIndex        =   180
         Top             =   2790
         Width           =   13965
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
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
         Height          =   195
         Index           =   50
         Left            =   -74280
         TabIndex        =   181
         Top             =   7155
         Width           =   1065
      End
   End
End
Attribute VB_Name = "frmOrdenRentaFijaLargoPlazo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Ordenes de Operaciones de Reporte con Acciones"
Option Explicit

Dim arrFondo()              As String, arrFondoOrden()              As String
Dim arrTipoInstrumento()    As String, arrTipoInstrumentoOrden()    As String
Dim arrEstado()             As String, arrTipoOrden()               As String
Dim arrOperacion()          As String, arrNegociacion()             As String
Dim arrEmisor()             As String, arrMoneda()                  As String
Dim arrBaseAnual()          As String, arrTipoTasa()                As String
Dim arrOrigen()             As String, arrClaseInstrumento()        As String
Dim arrTitulo()             As String, arrSubClaseInstrumento()     As String
Dim arrAgente()             As String, arrConceptoCosto()           As String
Dim strCodFondo             As String, strCodFondoOrden             As String
Dim strCodTipoInstrumento   As String, strCodTipoInstrumentoOrden   As String
Dim strCodEstado            As String, strCodTipoOrden              As String
Dim strCodOperacion         As String, strCodNegociacion            As String
Dim strCodEmisor            As String, strCodMoneda                 As String
Dim strCodBaseAnual         As String, strCodTipoTasa               As String
Dim strCodOrigen            As String, strCodClaseInstrumento       As String
Dim strCodTitulo            As String, strCodSubClaseInstrumento    As String
Dim strCodAgente            As String, strCodMonedaGarantia         As String
Dim strCodConcepto          As String, strCodGarantia               As String
Dim strEstado               As String, strNemonico                  As String

Dim strCodFile              As String, strCodAnalitica              As String
Dim strCodGrupo             As String, strCodCiiu                   As String
Dim strEstadoOrden          As String, strCodCategoria              As String
Dim strCodRiesgo            As String, strCodSubRiesgo              As String
Dim strCalcVcto             As String, strCodSector                 As String
Dim strCodTipoCostoBolsa    As String, strCodTipoCostoConasev       As String
Dim strCodTipoCostoFondo    As String, strCodTipoCavali             As String
Dim strIndCuponCero         As String, strIndPacto                  As String
Dim strIndNegociable        As String, strSQL                       As String
Dim strCodIndiceFinal       As String, strCodTipoAjuste             As String
Dim strCodPeriodoPago       As String, strCodIndiceInicial          As String
Dim strCodigosFile          As String, strIndAmortizacion           As String
Dim curSaldoAmortizacion    As Currency, curCantidadTitulo          As Currency
Dim dblTipoCambio           As Double, dblValorAmortizacion         As Double
Dim dblComisionBolsa        As Double, dblComisionConasev           As Double
Dim dblComisionFondo        As Double, dblComisionCavali            As Double
Dim dblVacCorrido           As Double
Dim datFechaEmision         As Date, intBaseAnual                   As Integer
Dim blnMonto                As Boolean, blnCantidad                 As Boolean
Public oExportacion As clsExportacion
Public indOk As Boolean
Dim adoExportacion As ADODB.Recordset
Dim adoConsulta             As ADODB.Recordset
Dim indSortAsc              As Boolean, indSortDesc                 As Boolean

Public Sub Accion(nAccion As ButtonAction)
    
    Select Case nAccion
                
        Case vNew
            Call Adicionar
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

Private Sub ActualizaComision(ctrlPorcentaje As Control, ctrlComision As Control)

    If Not IsNumeric(txtSubTotal(ctrlPorcentaje.Index).Text) Or Not IsNumeric(ctrlPorcentaje) Then Exit Sub
        
    If CDbl(ctrlPorcentaje) > 0 Then
        ctrlComision = CStr(CCur(txtSubTotal(ctrlPorcentaje.Index).Text) * CDbl(ctrlPorcentaje) / 100)
    Else
        ctrlComision = "0"
    End If
        
End Sub
Public Sub Modificar()

End Sub

Private Function TodoOK() As Boolean
        
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
    
    If cboAgente.ListIndex <= 0 Then
        MsgBox "Debe seleccionar el Agente.", vbCritical, Me.Caption
        If cboAgente.Enabled Then cboAgente.SetFocus
        Exit Function
    End If
                              
    If cboTitulo.ListIndex <= 0 Then
        MsgBox "Debe seleccionar el Título.", vbCritical, Me.Caption
        If cboTitulo.Enabled Then cboTitulo.SetFocus
        Exit Function
    End If
    
    If CVDate(dtpFechaOrden.Value) > CVDate(dtpFechaLiquidacion.Value) Then
        MsgBox "La Fecha de Liquidación debe ser mayor o igual a la Fecha de la ORDEN.", vbCritical, Me.Caption
        If dtpFechaLiquidacion.Enabled Then dtpFechaLiquidacion.SetFocus
        Exit Function
    End If
        
    If Trim(txtDescripOrden.Text) = Valor_Caracter Then
        MsgBox "Debe indicar la Descripción de la ORDEN.", vbCritical, Me.Caption
        If txtDescripOrden.Enabled Then txtDescripOrden.SetFocus
        Exit Function
    End If
    
    If CDbl(txtPrecio(0).Text) = 0 Then
        MsgBox "Debe indicar el Precio.", vbCritical, Me.Caption
        If txtPrecio(0).Enabled Then txtPrecio(0).SetFocus
        tabReporte.Tab = 2
        Exit Function
    End If
    
    If CDbl(txtTipoCambio.Text) = 0 Then
        MsgBox "Debe indicar el Tipo de Cambio.", vbCritical, Me.Caption
        If txtTipoCambio.Enabled Then txtTipoCambio.SetFocus
        tabReporte.Tab = 2
        Exit Function
    End If
    
    If CCur(txtCantidad.Text) = 0 Then
        MsgBox "Debe indicar la Cantidad.", vbCritical, Me.Caption
        If txtCantidad.Enabled Then txtCantidad.SetFocus
        tabReporte.Tab = 2
        Exit Function
    End If
        
    If CCur(lblMontoTotal(0).Caption) = 0 Then
        MsgBox "El Monto al contado es Cero.", vbCritical, Me.Caption
        If txtSubTotal(0).Enabled Then txtSubTotal(0).SetFocus
        Exit Function
    End If
            
    '*** Validación de STOCK ***
    If strCodTipoOrden = Codigo_Orden_Venta Then
        If CCur(txtCantidad.Text) > CCur(lblStockNominal.Caption) Then
            MsgBox "Stock insuficiente para Registrar la Orden de Venta.", vbCritical, Me.Caption
            If txtCantidad.Enabled Then txtCantidad.SetFocus
            tabReporte.Tab = 2
            Exit Function
        End If
    End If
            
    '*** Si todo paso OK ***
    TodoOK = True
  
End Function

Private Sub ActualizaPorcentaje(ctrlComision As Control, ctrlPorcentaje As Control)

    If Not IsNumeric(ctrlComision) Or Not IsNumeric(txtSubTotal(ctrlComision.Index).Text) Then Exit Sub
                
    If CCur(txtSubTotal(ctrlComision.Index).Text) = 0 Then
        ctrlPorcentaje = "0"
    Else
        If CCur(ctrlComision) > 0 Then
            ctrlPorcentaje = CStr((CCur(ctrlComision) / CCur(txtSubTotal(ctrlComision.Index).Text)) * 100)
        Else
            ctrlPorcentaje = "0"
        End If
    End If
                
End Sub

Public Sub Adicionar()

'    If Not EsDiaUtil(gdatFechaActual) Then
'        MsgBox "No se puede negociar en un día no útil !", vbCritical, Me.Caption
'        Exit Sub
'    End If
    
    If cboTipoInstrumento.ListCount > 1 Then
        frmMainMdi.stbMdi.Panels(3).Text = "Adicionar Orden..."
                    
        strEstado = Reg_Adicion
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabReporte
            .TabEnabled(0) = False
            .TabEnabled(2) = False
            .Tab = 1
        End With
    Else
        MsgBox "Acceso a Negociación Denegada", vbCritical, Me.Caption
    End If
    
End Sub

Private Sub AplicarCostos(Index As Integer)
    
    If strCodTipoCostoBolsa = Codigo_Tipo_Costo_Monto Then
        txtComisionBolsa(Index).Text = CStr(dblComisionBolsa)
    Else
        AsignaComision strCodTipoCostoBolsa, dblComisionBolsa, txtComisionBolsa
    End If
    
    If strCodTipoCostoConasev = Codigo_Tipo_Costo_Monto Then
        txtComisionConasev(Index).Text = CStr(dblComisionConasev)
    Else
        AsignaComision strCodTipoCostoConasev, dblComisionConasev, txtComisionConasev(Index)
    End If
    
    If strCodTipoCostoFondo = Codigo_Tipo_Costo_Monto Then
        txtComisionFondo(Index).Text = CStr(dblComisionFondo)
    Else
        AsignaComision strCodTipoCostoFondo, dblComisionFondo, txtComisionFondo(Index)
    End If
    
    If strCodTipoCavali = Codigo_Tipo_Costo_Monto Then
        txtComisionCavali(Index).Text = CStr(dblComisionCavali)
    Else
        AsignaComision strCodTipoCavali, dblComisionCavali, txtComisionCavali(Index)
    End If
                     
    Call CalculoTotal(Index)
    
End Sub

Private Sub AsignaComision(strTipoComision As String, dblValorComision As Double, ctrlValorComision As Control)
    
    If Not IsNumeric(txtSubTotal(ctrlValorComision.Index).Text) Then Exit Sub
    
    If dblValorComision > 0 Then
        ctrlValorComision.Text = CStr(CCur(txtSubTotal(ctrlValorComision.Index).Text) * dblValorComision / 100)
    End If
            
End Sub

Public Sub Buscar()

    Dim strFechaOrdenDesde          As String, strFechaOrdenHasta        As String
    Dim strFechaLiquidacionDesde    As String, strFechaLiquidacionHasta  As String
    Dim datFechaSiguiente           As Date
    
    Set adoConsulta = New ADODB.Recordset
    
    Me.MousePointer = vbHourglass
    
    If Not IsNull(dtpFechaOrdenDesde.Value) Or Not IsNull(dtpFechaOrdenHasta.Value) Then
        strFechaOrdenDesde = Convertyyyymmdd(dtpFechaOrdenDesde.Value)
        datFechaSiguiente = DateAdd("d", 1, dtpFechaOrdenHasta.Value)
        strFechaOrdenHasta = Convertyyyymmdd(datFechaSiguiente)
    End If
    
    If Not IsNull(dtpFechaLiquidacionDesde.Value) Or Not IsNull(dtpFechaLiquidacionHasta.Value) Then
        strFechaLiquidacionDesde = Convertyyyymmdd(dtpFechaLiquidacionDesde.Value)
        datFechaSiguiente = DateAdd("d", 1, dtpFechaLiquidacionHasta.Value)
        strFechaLiquidacionHasta = Convertyyyymmdd(datFechaSiguiente)
    End If
    
    strSQL = "SELECT NumOrden,FechaOrden,FechaLiquidacion,CodTitulo,Nemotecnico,EstadoOrden,CodFile,CodAnalitica,TipoOrden,IOR.CodMoneda," & _
        "DescripOrden,CantOrden,ValorNominal,PrecioUnitarioMFL1,MontoTotalMFL1, CodSigno DescripMoneda " & _
        "FROM InversionOrden IOR JOIN TipoOperacionNegociacion TON ON(TON.CodTipoOperacion=IOR.TipoOrden) " & _
        "JOIN Moneda MON ON(MON.CodMoneda=IOR.CodMoneda) " & _
        "WHERE CodAdministradora='" & gstrCodAdministradora & "' AND CodFondo='" & strCodFondo & "' "
        
    If strCodTipoInstrumento <> Valor_Caracter Then
        strSQL = strSQL & "AND CodFile='" & strCodTipoInstrumento & "' "
    Else
        strSQL = strSQL & "AND CodFile IN " & strCodigosFile & " "
    End If

    If Not IsNull(dtpFechaOrdenDesde.Value) Or Not IsNull(dtpFechaOrdenHasta.Value) Then
        strSQL = strSQL & "AND (FechaOrden >='" & strFechaOrdenDesde & "' AND FechaOrden <'" & strFechaOrdenHasta & "') "
    End If
    
    If Not IsNull(dtpFechaLiquidacionDesde.Value) Or Not IsNull(dtpFechaLiquidacionHasta.Value) Then
        strSQL = strSQL & "AND (FechaLiquidacion >='" & strFechaLiquidacionDesde & "' AND FechaLiquidacion <'" & strFechaLiquidacionHasta & "') "
    End If
    
    If strCodEstado <> Valor_Caracter Then
        strSQL = strSQL & "AND EstadoOrden='" & strCodEstado & "' "
    End If
    
    strSQL = strSQL & "ORDER BY NumOrden"
    
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

    Me.MousePointer = vbDefault
    
End Sub

Private Sub CalculoTotal(Index As Integer)

    Dim curComImp As Currency, curMonTotal As Currency

    If Not (IsNumeric(txtComisionAgente(Index).Text) And IsNumeric(txtComisionBolsa(Index).Text) And IsNumeric(txtComisionConasev(Index).Text) And IsNumeric(txtComisionCavali(Index).Text) And IsNumeric(txtComisionFondo(Index).Text)) Then Exit Sub
    
    curComImp = CCur(CCur(txtComisionAgente(Index).Text) + CCur(txtComisionBolsa(Index).Text) + CCur(txtComisionConasev(Index).Text) + CCur(txtComisionCavali(Index).Text) + CCur(txtComisionFondo(Index).Text)) * CDbl(lblPorcenIgv(Index).Caption)
    lblComisionIgv(Index).Caption = CStr(curComImp)

    curComImp = CCur(txtComisionAgente(Index).Text) + CCur(txtComisionBolsa(Index).Text) + CCur(txtComisionConasev(Index).Text) + CCur(txtComisionCavali(Index).Text) + CCur(txtComisionFondo(Index).Text) + CCur(lblComisionIgv(Index).Caption)
    
    lblComisionesResumen(Index).Caption = CStr(curComImp)

    If strCodTipoOrden = Codigo_Orden_Compra Or strCodTipoOrden = Codigo_Orden_Pacto Then  '*** Compra ***
        If Index = 0 Then
            curMonTotal = CCur(txtSubTotal(Index).Text) + curComImp
        Else
            curMonTotal = CCur(txtSubTotal(Index).Text) - curComImp
        End If
    ElseIf strCodTipoOrden = Codigo_Orden_Venta Then '*** Venta ***
        curMonTotal = CCur(txtSubTotal(Index).Text) - curComImp
    End If
        
    curMonTotal = curMonTotal + CCur(txtInteresCorrido(Index).Text) + CCur(txtVacCorrido(Index).Text)
    lblMontoTotal(Index).Caption = CStr(curMonTotal)
    
End Sub


Public Sub Cancelar()

    cmdOpcion.Visible = True
    With tabReporte
        .TabEnabled(0) = True
        .Tab = 0
    End With
    Call Buscar
    
End Sub

Public Sub CargarComisiones(ByVal strCodComision As String, Index As Integer)
     
     Call AplicarCostos(Index)
     
End Sub

Private Sub CargarListas()

    Dim intRegistro As Integer
    
    '*** Fondos ***
    strSQL = "{ call up_ACSelDatosParametro(74,'" & gstrCodAdministradora & "','" & gstrCodFondoContable & "') }"
    CargarControlLista strSQL, cboFondo, arrFondo(), Valor_Caracter
    CargarControlLista strSQL, cboFondoOrden, arrFondoOrden(), Valor_Caracter
        
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
            
    '*** Estados de la Orden ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='ESTORD' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboEstado, arrEstado(), Sel_Todos
    
    intRegistro = ObtenerItemLista(arrEstado(), Estado_Orden_Ingresada)
    If intRegistro >= 0 Then cboEstado.ListIndex = intRegistro
        
    '*** Tipo de Orden ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPORD' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboTipoOrden, arrTipoOrden(), Valor_Caracter

    '*** Tipo Liquidación Operación ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPLIQ' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboOperacion, arrOperacion(), Valor_Caracter
    
    '*** Mecanismos de Negociación ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='MECNEG' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboNegociacion, arrNegociacion(), Valor_Caracter

    '*** Conceptos de Costos de Negociación ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPCCO' AND ValorParametro='RF' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboConceptoCosto, arrConceptoCosto(), Sel_Defecto
    
    '*** Agente ***
    strSQL = "SELECT (CodPersona + CodGrupo + CodCiiu) CODIGO,DescripPersona DESCRIP FROM InstitucionPersona WHERE TipoPersona='" & Codigo_Tipo_Persona_Agente & "' AND IndVigente='X' ORDER BY DescripPersona"
    CargarControlLista strSQL, cboAgente, arrAgente(), Sel_Defecto

    '*** Mercado de Negociación ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='MDONEG' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboOrigen, arrOrigen(), Valor_Caracter
                    
    '*** Tipo de Cálculo ***
    cboCalculo.AddItem Sel_Defecto, 0
    cboCalculo.AddItem "Tir Bruta", 1
    cboCalculo.AddItem "Tir Neta", 2
    cboCalculo.AddItem "Precio", 3
        
End Sub
Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Ordenes de Inversión"
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Text = "Papeleta de Inversión"
    
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

Public Sub Eliminar()

    If strEstado = Reg_Consulta Or strEstado = Reg_Edicion Then
        Dim strMensaje  As String
        
        strMensaje = "Se procederá a eliminar la ORDEN " & tdgConsulta.Columns(1) & " por la " & _
            tdgConsulta.Columns(3) & vbNewLine & vbNewLine & vbNewLine & "¿ Seguro de continuar ?"
        
        If MsgBox(strMensaje, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
    
            '*** Anular Orden ***
            adoComm.CommandText = "UPDATE InversionOrden SET EstadoOrden='" & Estado_Orden_Anulada & "' WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodTitulo='" & Trim(tdgConsulta.Columns(2)) & "' AND NumOrden='" & Trim(tdgConsulta.Columns(1)) & "'"
                
            adoConn.Execute adoComm.CommandText
            
            '*** Anular Título si corresponde ***
            adoComm.CommandText = "UPDATE InstrumentoInversion SET IndVigente='' WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodTitulo='" & Trim(tdgConsulta.Columns(2)) & "'"
                
            adoConn.Execute adoComm.CommandText
            
            MsgBox Mensaje_Eliminacion_Exitosa, vbExclamation, Me.Caption
            
            tabReporte.TabEnabled(0) = True
            tabReporte.Tab = 0
            Call Buscar
            
            Exit Sub
        End If
    End If
    
End Sub

Public Sub Grabar()

    Dim adoRegistro         As ADODB.Recordset
    Dim strFechaOrden       As String, strFechaLiquidacion      As String
    Dim strFechaEmision     As String, strFechaVencimiento      As String
    Dim strMensaje          As String, strIndTitulo             As String
    Dim strCodReportado     As String
    Dim intRegistro         As Integer, intAccion               As Integer
    Dim lngNumError         As Long
    Dim strDescripOrden     As String
    
    
    'On Error GoTo CtrlError
    
    If strEstado = Reg_Consulta Then Exit Sub
    
    If strEstado = Reg_Adicion Then
        If TodoOK() Then
            strEstadoOrden = Estado_Orden_Ingresada
            
            '*** Validación del Límite de Inversión con respecto al Activo ***
'            If strCodTipoOrden = Codigo_Orden_Compra Then
'                Me.MousePointer = vbHourglass
'                intRes = ValidLimActivo(strCodFile, strCodFondoOrden, Convertyyyymmdd(dtpFechaOrden.Value), strCodMoneda, CCur(lblMontoTotal.Caption), CDbl(txtTipoCambio.Text), strCodEmisor)
'                Me.MousePointer = vbDefault
'
'                If intRes = 0 Then
'                    strEstadoOrden = Estado_Orden_PorAutorizar
'                End If
'            End If

            '*** Validación del Límite de Línea de Crédito Vigente ***
'            If strCodTipoOrden = Codigo_Orden_Compra Then
'                Me.MousePointer = vbHourglass
'                intRes = ValidLimCobertura(strCodEmisor, Convertyyyymmdd(dtpFechaOrden.Text), lblDescripMoneda.Tag, CCur(lblMontoTotal.Caption), CDbl(txtTipoCambio.Text))
'                Me.MousePointer = vbDefault
'
'                If intRes = 0 Then
'                    strEstadoOrden = Estado_Orden_PorAutorizar
'                End If
'            End If

        
            strMensaje = "Para proceder al Registro de la Orden Confirme lo siquiente : " & Chr(vbKeyReturn) & Chr(vbKeyReturn) & _
                "Fecha de Operación" & Space(4) & ">" & Space(2) & CStr(dtpFechaOrden.Value) & Chr(vbKeyReturn) & _
                "Fecha de Liquidación" & Space(3) & ">" & Space(2) & CStr(dtpFechaLiquidacion.Value) & Chr(vbKeyReturn) & Chr(vbKeyReturn)
                                    
            strMensaje = strMensaje & "Cantidad" & Space(22) & ">" & Space(2) & txtCantidad.Text & Chr(vbKeyReturn) & _
                "Precio Unitario (%)" & Space(6) & ">" & Space(2) & txtPrecio(0).Text & Chr(vbKeyReturn) & Chr(vbKeyReturn) & _
                "Monto Total" & Space(17) & ">" & Space(2) & Trim(lblDescripMonedaResumen(0).Caption) & Space(1) & lblMontoTotal(0).Caption & Chr(vbKeyReturn) & Chr(vbKeyReturn) & _
                "Tir Bruta (%)" & Space(23) & ">" & Space(2) & txtTirBrutaBaseBono.Text & Chr(vbKeyReturn) & _
                "Tir Neta (%)" & Space(23) & ">" & Space(2) & txtTirNetaBaseBono.Text & Chr(vbKeyReturn) & _
                Chr(vbKeyReturn) & Chr(vbKeyReturn) & _
                "¿ Seguro de continuar ?"

            If MsgBox(strMensaje, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
               Me.Refresh: Exit Sub
            End If

        
            Me.MousePointer = vbHourglass
            
            strFechaOrden = Convertyyyymmdd(dtpFechaOrden.Value)
            strFechaEmision = Convertyyyymmdd(dtpFechaOrden.Value)
            strFechaLiquidacion = Convertyyyymmdd(dtpFechaLiquidacion.Value)
            strFechaVencimiento = strFechaLiquidacion
            
            strDescripOrden = UCase(Trim(cboTipoOrden.Text) & " " & Trim(cboTipoInstrumentoOrden.Text) & " " & Trim(Left(cboTitulo.Text, 15)) & " CANT: " & (txtCantidad.Text) & " PRECIO: " & (txtPrecio(0).Text) & "%")
            
            
            Set adoRegistro = New ADODB.Recordset
            '*** Guardar Orden de Inversion ***
            With adoComm
                If strCodTipoOrden = Codigo_Orden_Pacto Then
                    strIndTitulo = Valor_Caracter
                    strCodAnalitica = NumAleatorio(8)
                    strCodTitulo = NumAleatorio(15)
                    strCodTipoTasa = Codigo_Tipo_Tasa_Efectiva
                    strCodBaseAnual = Codigo_Base_Actual_365
                    strCodRiesgo = "00" ' Sin Clasificacion
                    strCodReportado = strCodAgente
                    strCodFile = Left(Trim(lblAnalitica.Caption), 3)
                Else
                    strIndTitulo = Valor_Indicador
                    strCodTitulo = strCodGarantia
                    strCodGarantia = Valor_Caracter
                    strCodMoneda = lblMoneda.Tag
                    strFechaVencimiento = Convertyyyymmdd(Valor_Fecha)
                    strCodReportado = Valor_Caracter
                End If
                                            
'                .CommandText = "BEGIN TRAN ProcOrden"
'                adoConn.Execute .CommandText
                                
'                .CommandText = "{ call up_IVAdicInversionOrden('" & strCodFondoOrden & "','" & _
'                    gstrCodAdministradora & "','','" & strFechaOrden & "','" & _
'                    strCodTitulo & "','" & strNemonico & "','" & gstrPeriodoActual & "','" & gstrMesActual & "','" & _
'                    "','" & strEstadoOrden & "','" & strCodAnalitica & "','" & strCodFile & "','" & _
'                    strCodAnalitica & "','" & strCodClaseInstrumento & "','" & strCodSubClaseInstrumento & "','" & strCodTipoOrden & "','" & _
'                    strCodOperacion & "','" & strCodNegociacion & "','" & strCodOrigen & "','" & Trim(txtDescripOrden.Text) & "','" & strCodEmisor & "','" & _
'                    strCodAgente & "','" & strCodGarantia & "','','" & Convertyyyymmdd(CVDate(Valor_Fecha)) & "','" & strFechaVencimiento & "','" & strFechaLiquidacion & "','" & _
'                    strFechaEmision & "','" & strCodMoneda & "','" & strCodMoneda & "','" & strCodMoneda & "'," & CDec(txtCantidad.Text) & "," & _
'                    CDec(txtTipoCambio.Text) & "," & CDec(txtTipoCambio.Text) & "," & CDec(lblValorNominal.Caption) & "," & _
'                    CDec(txtPrecio(0).Text) & "," & CDec(txtSubTotal(0).Text) & "," & CDec(txtSubTotal(0).Text) & "," & CDec(txtInteresCorrido(0).Text) & "," & _
'                    CDec(txtComisionAgente(0).Text) & "," & CDec(txtComisionCavali(0).Text) & "," & CDec(txtComisionConasev(0).Text) & "," & _
'                    CDec(txtComisionBolsa(0).Text) & "," & CDec(txtComisionFondo(0).Text) & ",0,0,0," & CDec(lblComisionIgv(0).Caption) & "," & CDec(lblMontoTotal(0).Caption) & "," & CDec(lblMontoTotal(0).Caption) & "," & _
'                    CDec(txtPrecio(1).Text) & "," & CDec(txtSubTotal(1).Text) & "," & CDec(txtInteresCorrido(1).Text) & "," & CDec(txtComisionAgente(1).Text) & "," & _
'                    CDec(txtComisionCavali(1).Text) & "," & CDec(txtComisionConasev(1).Text) & "," & CDec(txtComisionBolsa(1).Text) & "," & _
'                    CDec(txtComisionFondo(1).Text) & ",0,0,0," & CDec(lblComisionIgv(1).Caption) & "," & CDec(lblMontoTotal(1).Caption) & "," & _
'                    CDec(lblMontoTotal(1).Caption) & ",0,'','','','" & strCodReportado & "','',''," & CDec(txtVacCorrido(0).Text) & ",'','','" & strIndTitulo & "','" & _
'                    strCodTipoTasa & "','" & strCodBaseAnual & "'," & CDec(txtTasaMensual.Text) & "," & CDec(lblTirBrutaResumen.Caption) & "," & CDec(lblTirNetaResumen.Caption) & ",'" & _
'                    strCodRiesgo & "','" & strCodSubRiesgo & "','" & Trim(txtObservacion.Text) & "','" & gstrLogin & "') }"
'                adoConn.Execute .CommandText
                
                .CommandText = "{ call up_IVAdicInversionOrden('" & strCodFondoOrden & "','" & _
                    gstrCodAdministradora & "','','" & strFechaOrden & "','" & _
                    strCodTitulo & "','" & strNemonico & "','" & gstrPeriodoActual & "','" & gstrMesActual & "','" & _
                    "','" & strEstadoOrden & "','" & strCodAnalitica & "','" & strCodFile & "','" & _
                    strCodAnalitica & "','" & strCodClaseInstrumento & "','" & strCodSubClaseInstrumento & "','" & strCodTipoOrden & "','" & _
                    strCodOperacion & "','" & strCodNegociacion & "','" & strCodOrigen & "','" & strDescripOrden & "','" & strCodEmisor & "','" & _
                    strCodAgente & "','" & strCodGarantia & "','" & Convertyyyymmdd(CVDate(Valor_Fecha)) & "','" & strFechaVencimiento & "','" & strFechaLiquidacion & "','" & _
                    strFechaEmision & "','" & strCodMoneda & "'," & CDec(txtCantidad.Text) & "," & _
                    CDbl(txtTipoCambio.Text) & "," & CDbl(lblValorNominal.Caption) & ",100," & CDbl(lblValorNominal.Caption) & "," & _
                    CDbl(txtPrecio(0).Text) & "," & CDbl(txtPrecioSucio.Value) & "," & CDec(txtSubTotal(0).Text) & "," & CDec(txtInteresCorrido(0).Text) & "," & _
                    CDec(txtComisionAgente(0).Text) & "," & CDec(txtComisionCavali(0).Text) & "," & CDec(txtComisionConasev(0).Text) & "," & _
                    CDec(txtComisionBolsa(0).Text) & "," & CDec(txtComisionFondo(0).Text) & ",0,0,0," & CDec(lblComisionIgv(0).Caption) & "," & CDec(lblMontoTotal(0).Caption) & "," & _
                    CDbl(txtPrecio(1).Text) & "," & CDbl(txtPrecio(1).Text) & "," & CDec(txtSubTotal(1).Text) & "," & CDec(txtInteresCorrido(1).Text) & "," & CDec(txtComisionAgente(1).Text) & "," & _
                    CDec(txtComisionCavali(1).Text) & "," & CDec(txtComisionConasev(1).Text) & "," & CDec(txtComisionBolsa(1).Text) & "," & _
                    CDec(txtComisionFondo(1).Text) & ",0,0,0," & CDec(lblComisionIgv(1).Caption) & "," & CDec(lblMontoTotal(1).Caption) & "," & _
                    "0,0,'','','','','','" & strCodReportado & "','','','','',''," & CDec(txtVacCorrido(0).Text) & ",'','','" & strIndTitulo & "','" & _
                    strCodTipoTasa & "','" & strCodBaseAnual & "'," & CDec(txtTasaMensual.Text) & "," & CDbl(txtTirBrutaBaseBono.Value) & "," & CDbl(txtTirBrutaLimpia.Value) & "," & CDbl(txtTirNetaBaseBono.Value) & ",'" & _
                    strCodRiesgo & "','" & strCodSubRiesgo & "','" & Trim(txtObservacion.Text) & "') }"
                adoConn.Execute .CommandText

'                .CommandText = "COMMIT TRAN ProcOrden"
'                adoConn.Execute .CommandText
                                                                                                      
            End With
                                                                                    
            Me.MousePointer = vbDefault
        
            MsgBox Mensaje_Adicion_Exitosa, vbExclamation
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
            cmdOpcion.Visible = True
            With tabReporte
                .TabEnabled(0) = True
                .Tab = 0
            End With
            Call Buscar
        End If
    End If
    Exit Sub
        
CtrlError:
    If strCodTipoOrden <> Codigo_Orden_Pacto Then strCodGarantia = strCodTitulo
    
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
    
'    MsgBox adoConn.Errors.Item(0).Description & vbNewLine & vbNewLine & Mensaje_Proceso_NoExitoso, vbCritical
'    Me.MousePointer = vbDefault
        
End Sub

Public Sub Imprimir()

End Sub

Public Sub SubImprimir(Index As Integer)

    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    Dim strFechaDesde           As String, strFechaHasta        As String
    Dim strSeleccionRegistro    As String

    If tabReporte.Tab = 1 Then Exit Sub
    
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
                aReportParamF(1) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10)
                aReportParamF(2) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10)
                aReportParamF(3) = Format(Time(), "hh:mm:ss")
                aReportParamF(4) = Trim(cboFondo.Text)
                aReportParamF(5) = gstrNombreEmpresa & Space(1)
                            
                aReportParamS(0) = strCodFondo
                aReportParamS(1) = gstrCodAdministradora
                aReportParamS(2) = Convertyyyymmdd(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10))
                aReportParamS(3) = Convertyyyymmdd(DateAdd("d", 1, CVDate(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10))))
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
                ReDim aReportParamFn(5)
                ReDim aReportParamF(5)
                            
                aReportParamFn(0) = "Usuario"
                aReportParamFn(1) = "FechaDesde"
                aReportParamFn(2) = "FechaHasta"
                aReportParamFn(3) = "Hora"
                aReportParamFn(4) = "Fondo"
                aReportParamFn(5) = "NombreEmpresa"
                            
                aReportParamF(0) = gstrLogin
                aReportParamF(1) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10)
                aReportParamF(2) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10)
                aReportParamF(3) = Format(Time(), "hh:mm:ss")
                aReportParamF(4) = Trim(cboFondo.Text)
                aReportParamF(5) = gstrNombreEmpresa & Space(1)
                            
                aReportParamS(0) = strCodFondo
                aReportParamS(1) = gstrCodAdministradora
                aReportParamS(2) = Convertyyyymmdd(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10))
                aReportParamS(3) = Convertyyyymmdd(DateAdd("d", 1, CVDate(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10))))
                aReportParamS(4) = strCodMoneda
                aReportParamS(5) = strCodTipoInstrumento
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

Private Sub InicializarValores()
    
    Dim adoRegistro As ADODB.Recordset
    
    '*** Valores Iniciales ***
    strEstado = Reg_Defecto
    tabReporte.Tab = 0

    dtpFechaOrdenDesde.Value = gdatFechaActual
    dtpFechaOrdenHasta.Value = dtpFechaOrdenDesde.Value
    dtpFechaLiquidacionDesde.Value = Null
    dtpFechaLiquidacionHasta.Value = Null
    
    lblPorcenIgv(0).Caption = CStr(gdblTasaIgv)
    lblPorcenIgv(1).Caption = CStr(gdblTasaIgv)
    chkAjustePrecio(0).Value = vbUnchecked
    
    Set adoRegistro = New ADODB.Recordset
    With adoComm
        .CommandText = "SELECT CodFile FROM InversionFile  " & _
            "WHERE TipoValor='" & Codigo_Valor_RentaFija & "' AND TipoPlazo='" & Codigo_Valor_LargoPlazo & "' AND IndInstrumento='X' AND IndVigente='X' " & _
            "ORDER BY DescripFile"
        Set adoRegistro = .Execute
                
        strCodigosFile = Valor_Caracter
        Do While Not adoRegistro.EOF
            If strCodigosFile <> Valor_Caracter Then strCodigosFile = strCodigosFile & ",'"
            
            strCodigosFile = strCodigosFile & Trim(adoRegistro("CodFile")) & "'"
        
            adoRegistro.MoveNext
        Loop
        adoRegistro.Close: Set adoRegistro = Nothing
        
        strCodigosFile = "('" & strCodigosFile & ",'009')"
    End With
    
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgConsulta.Columns(1).Width = tdgConsulta.Width * 0.01 * 8
    tdgConsulta.Columns(8).Width = tdgConsulta.Width * 0.01 * 12
    tdgConsulta.Columns(3).Width = tdgConsulta.Width * 0.01 * 32
    tdgConsulta.Columns(0).Width = tdgConsulta.Width * 0.01 * 6
    tdgConsulta.Columns(4).Width = tdgConsulta.Width * 0.01 * 13
    tdgConsulta.Columns(6).Width = tdgConsulta.Width * 0.01 * 11
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
                
End Sub

Private Sub IniciarComisiones()

    Dim intContador As Integer
    
    For intContador = 0 To 1
        txtComisionAgente(intContador).Text = "0"
        txtComisionBolsa(intContador).Text = "0"
        txtComisionCavali(intContador).Text = "0"
        txtComisionFondo(intContador).Text = "0"
        txtComisionConasev(intContador).Text = "0"
        lblComisionIgv(intContador).Caption = "0"
        
        txtPorcenAgente(intContador).Text = "0"
        lblPorcenBolsa(intContador).Caption = "0"
        lblPorcenCavali(intContador).Caption = "0"
        lblPorcenFondo(intContador).Caption = "0"
        lblPorcenConasev(intContador).Caption = "0"
        
        lblPrecioResumen(intContador).Caption = "0"
        lblSubTotalResumen(intContador).Caption = "0"
        lblComisionesResumen(intContador).Caption = "0"
        lblInteresesResumen(intContador).Caption = "0"
        lblTotalResumen(intContador).Caption = "0"
        
    Next
    
End Sub

Private Sub LlenarFormulario(strModo As String)

    Dim adoRecord   As ADODB.Recordset
    Dim strSQL      As String
    Dim intRegistro As Integer
    
    Select Case strModo
        Case Reg_Adicion
        
            blnMonto = False
            blnCantidad = False
            
            intRegistro = ObtenerItemLista(arrFondoOrden(), strCodFondo)
            If intRegistro >= 0 Then cboFondoOrden.ListIndex = intRegistro
        
            cboTipoInstrumentoOrden.ListIndex = -1
            If cboTipoInstrumentoOrden.ListCount > 0 Then cboTipoInstrumentoOrden.ListIndex = 0
                                    
            cboTipoOrden.ListIndex = -1
            If cboTipoOrden.ListCount > 0 Then cboTipoOrden.ListIndex = 0
    
            cboOperacion.ListIndex = -1
            If cboOperacion.ListCount > 0 Then cboOperacion.ListIndex = 0
        
            cboNegociacion.ListIndex = -1
            If cboNegociacion.ListCount > 0 Then cboNegociacion.ListIndex = 0
            
            cboConceptoCosto.ListIndex = -1
            If cboConceptoCosto.ListCount > 0 Then cboConceptoCosto.ListIndex = 0
            
            cboAgente.ListIndex = -1
            If cboAgente.ListCount > 0 Then cboAgente.ListIndex = 0
            
            intRegistro = ObtenerItemLista(arrOrigen(), Codigo_Negociacion_Local)
            If intRegistro >= 0 Then cboOrigen.ListIndex = intRegistro
            
            dtpFechaOrden.Value = gdatFechaActual
            dtpFechaLiquidacion.Value = dtpFechaOrden.Value
            lblFechaLiquidacion.Caption = CStr(dtpFechaOrden.Value)
                                    
            txtDescripOrden.Text = Valor_Caracter
            txtObservacion.Text = Valor_Caracter
            txtPrecio(0).Text = "0"
            txtPrecio(1).Text = "0"
            txtInteresCorrido(0).Text = "0"
            txtInteresCorrido(1).Text = "0"
            txtVacCorrido(0).Text = "0"
            txtVacCorrido(1).Text = "0"
            txtCantidad.Text = "0"
            lblAnalitica.Caption = "??? - ????????"
            lblIndiceInicial.Caption = Valor_Caracter
            lblIndiceFinal.Caption = Valor_Caracter
            lblIndiceInicial.ToolTipText = Valor_Caracter
            lblIndiceFinal.ToolTipText = Valor_Caracter
                                    
            txtSubTotal(0).Text = "0"
            txtSubTotal(1).Text = "0"
            
            cboCalculo.ListIndex = -1
            If cboCalculo.ListCount > 0 Then cboCalculo.ListIndex = 0
                        
            txtTasaMensual.Text = "0"
                                                
            chkAplicar(0).Value = vbUnchecked
            chkAplicar(1).Value = vbUnchecked
                        
            Call IniciarComisiones
            
            txtTirBrutaBaseBono.Text = "0"
            txtTirNetaBaseBono.Text = "0"
            txtTirBrutaBase365.Text = "0"
            txtTirNetaBase365.Text = "0"
            lblValorNominal.Caption = "0"
            lblFechaCupon.Caption = Valor_Caracter
            lblInicioFechaCupon.Caption = Valor_Caracter
            lblClasificacion.Caption = Valor_Caracter
            lblBaseTasaCupon.Caption = Valor_Caracter
            lblStockNominal.Caption = "0"
            lblMoneda.Caption = Valor_Caracter
            lblCantidadResumen.Caption = "0"
                        
            lblMontoTotal(0).Caption = "0"
            lblMontoTotal(1).Caption = "0"
            lblTirBruta.Caption = "0"
            lblTirNeta.Caption = "0"
            lblTirBrutaResumen.Caption = "0"
            lblTirNetaResumen.Caption = "0"
            
            chkAjustePrecio(0).Value = vbChecked
                        
            cboFondoOrden.SetFocus
                        
        Case Reg_Edicion
    
    End Select
    
    chkAplicar(0).Enabled = False
    
End Sub

Private Function PosicionLimites() As Boolean

    PosicionLimites = False
        
    If cboTipoInstrumentoOrden.ListIndex <= 0 Then
        MsgBox "Debe seleccionar el Tipo de Instrumento.", vbCritical, Me.Caption
        'cboEmisor.ListIndex = -1: cboTitulo.ListIndex = -1
        If cboTipoInstrumentoOrden.Enabled Then cboTipoInstrumentoOrden.SetFocus
        Exit Function
    End If

'    If strCodTipoOrden = Codigo_Orden_Compra Then ValidLimites strCodEmisor, Convertyyyymmdd(dtpFechaOrden.Value), CDbl(txtTipoCambio.Text), strCodFile, strCodFondoOrden

    '*** Si todo pasó OK ***
    PosicionLimites = True
    
End Function
Public Sub Salir()

    Unload Me
    
End Sub

Private Sub cboAgente_Click()

    strCodAgente = Valor_Caracter
    If cboAgente.ListIndex < 0 Then Exit Sub
    
    strCodAgente = Trim(arrAgente(cboAgente.ListIndex))
    
End Sub
Private Sub CalcularPrecioTIRVAN()


    Dim dblTir  As Double
    Dim adoTemporal As ADODB.Recordset
    
    '*** Calculo del PRECIO a partir de la TIR BRUTA(365) ***
    Dim intRes As Integer
    Dim dblPrecio As Double, dblValTir As Double
    Dim dblMtoNomi As Double
    Dim strFchCalc As String
            
    If CDbl(txtCantidad.Text) <= 0 Then Exit Sub
    
    If cboTitulo.ListIndex <= 0 Then Exit Sub
            
    If CDbl(txtTirBrutaBaseBono.Text) <= 0 Then Exit Sub

    dblValTir = CDbl(txtTirBrutaBaseBono.Text)
    
    dblPrecio = 0
    dblMtoNomi = CDbl(txtCantidad.Text)
    strFchCalc = Convertyyyymmdd(dtpFechaLiquidacion.Value)
                              
    '*** Inicio de los calculos ***
    If dblValTir > 0 Then
        '*** Calculo del PRECIO para BONOS ***
                
        Set adoTemporal = New ADODB.Recordset
        adoComm.CommandText = "SELECT dbo.uf_ACVNANoPer('" & strCodGarantia & "','" & dtpFechaLiquidacion & "','" & CVDate(lblFechaCupon.Caption) & "'," & txtMontoNominal.Value & "," & curCantidadTitulo & "," & dblValTir & ",'" & strCodTipoAjuste & "','" & strCodIndiceInicial & "','" & strCodIndiceFinal & "') AS 'ValorActual'"
        Set adoTemporal = adoComm.Execute
        
        If Not adoTemporal.EOF Then
            txtValorActual.Text = adoTemporal("ValorActual")
        End If
        adoTemporal.Close
        
        'txtValorActual.Text = VNANoPer(strCodGarantia, dtpFechaLiquidacion, CVDate(lblFechaCupon.Caption), txtMontoNominal.Value, curCantidadTitulo, dblValTir, strCodTipoAjuste, strCodIndiceInicial, strCodIndiceFinal)
        dblPrecio = ((txtValorActual.Value) - CDbl(txtInteresCorrido(0).Value)) / (txtMontoNominal.Value)
       
        '*** PRECIO a TIR BRUTA ***
        txtPrecio(0).Text = CStr(dblPrecio * 100) 'PRECIO LIMPIO
        
        txtSubTotal(0).Text = txtMontoNominal.Value * dblPrecio 'curCantidad * dblPreUni
        
        dblPrecio = ((txtValorActual.Value)) / (txtMontoNominal.Value)
        
        txtPrecioSucio.Text = CStr(dblPrecio * 100) 'PRECIO SUCIO
    End If

    'TIR BRUTA SUCIA
    'dblTir = TirNoPer(strCodGarantia, dtpFechaLiquidacion, lblFechaCupon.Caption, CDbl(txtValorActual.Value), 0, txtMontoNominal.Value, txtCantidad.Value, 0.1, strCodTipoAjuste, strCodIndiceInicial, strCodIndiceFinal)

    'TIR BRUTA LIMPIA
    'dblTir = TirNoPer(strCodGarantia, dtpFechaLiquidacion, lblFechaCupon.Caption, CDbl(txtValorActual.Value) - CDbl(txtInteresCorrido(0).Text), 0, txtMontoNominal.Value, txtCantidad.Value, 0.1, strCodTipoAjuste, strCodIndiceInicial, strCodIndiceFinal)
    
    adoComm.CommandText = "SELECT dbo.uf_ACTirNoPer('" & strCodGarantia & "','" & dtpFechaLiquidacion & "','" & lblFechaCupon.Caption & "'," & CDbl(txtValorActual.Value - CDbl(txtInteresCorrido(0).Value)) & "," & 0 & "," & txtMontoNominal.Value & "," & txtCantidad.Value & "," & 0.1 & ",'" & strCodTipoAjuste & "','" & strCodIndiceInicial & "','" & strCodIndiceFinal & "') AS 'TirNoPer'"
    Set adoTemporal = adoComm.Execute
    
    If Not adoTemporal.EOF Then
        dblTir = adoTemporal("TirNoPer")
    End If
    adoTemporal.Close
    
    txtTirBrutaLimpia.Text = CStr(dblTir)

    'TIR NETA
    'dblTir = TirNoPer(strCodGarantia, dtpFechaLiquidacion, lblFechaCupon.Caption, CCur(txtSubTotal(0).Text) + CCur(lblComisionesResumen(0).Caption), CDbl(txtInteresCorrido(0).Value) + CDbl(txtVacCorrido(0).Text), txtMontoNominal.Value, txtCantidad.Value, 0.1, strCodTipoAjuste, strCodIndiceInicial, strCodIndiceFinal)
    
    adoComm.CommandText = "SELECT dbo.uf_ACTirNoPer('" & strCodGarantia & "','" & dtpFechaLiquidacion & "','" & lblFechaCupon.Caption & "'," & CDbl(CDbl(txtSubTotal(0).Value) + CDbl(lblComisionesResumen(0).Caption)) & "," & CDbl(CDbl(txtInteresCorrido(0).Value) + CDbl(txtVacCorrido(0).Text)) & "," & txtMontoNominal.Value & "," & txtCantidad.Value & "," & 0.1 & ",'" & strCodTipoAjuste & "','" & strCodIndiceInicial & "','" & strCodIndiceFinal & "') AS 'TirNoPer'"
    Set adoTemporal = adoComm.Execute
    
    If Not adoTemporal.EOF Then
        dblTir = adoTemporal("TirNoPer")
    End If
    adoTemporal.Close
    
    txtTirNetaBaseBono.Text = CStr(dblTir)


End Sub

Private Sub cboCalculo_Click()

    Dim dblTir  As Double
    Dim adoTemporal As ADODB.Recordset
    
    If cboCalculo.ListIndex < 0 Then Exit Sub

    Select Case UCase(Trim(cboCalculo.Text))
        Case "PRECIO"
            '*** Calculo del PRECIO a partir de la TIR BRUTA(365) ***
            Dim intRes As Integer
            Dim dblPrecio As Double, dblValTir As Double
            Dim dblMtoNomi As Double
            Dim strFchCalc As String
            
            If cboTitulo.ListIndex <= 0 Then
               MsgBox "Por favor seleccione el TITULO de la ORDEN.", vbCritical, Me.Caption
               cboTitulo.SetFocus
               Exit Sub
            End If
            
            If CInt(Left(lblBaseTasaCupon.Caption, 3)) = 360 Then
               If CDbl(txtTirBrutaBaseBono.Text) <= 0 Then
                  MsgBox "Por favor ingrese la TIR BRUTA de la ORDEN.", vbCritical, Me.Caption
                  If txtTirBrutaBaseBono.Enabled = True Then txtTirBrutaBaseBono.SetFocus
                  Exit Sub
               Else
                  dblValTir = CDbl(txtTirBrutaBaseBono.Text)
               End If
            ElseIf CInt(Left(lblBaseTasaCupon.Caption, 3)) = 365 Then
               If CDbl(txtTirBrutaBase365.Text) <= 0 Then
                  MsgBox "Por favor ingrese la TIR BRUTA de la ORDEN.", vbCritical, Me.Caption
                  If txtTirBrutaBase365.Enabled = True Then txtTirBrutaBase365.SetFocus
                  Exit Sub
               Else
                  dblValTir = CDbl(txtTirBrutaBase365.Text)
               End If
            End If
            
            If CDbl(txtCantidad.Text) <= 0 Then
               MsgBox "Por favor ingrese el Monto Nominal de la ORDEN.", vbCritical, Me.Caption
               txtCantidad.SetFocus
               Exit Sub
            End If
            
                
            dblPrecio = 0
            dblMtoNomi = CDbl(txtCantidad.Text)
            strFchCalc = Convertyyyymmdd(dtpFechaLiquidacion.Value)
            
            '*** Inicio de los calculos ***
            If dblValTir > 0 Then
                '*** Calculo del PRECIO para BONOS ***
                If strCodFile = "005" Then
                    If strCodTipoAjuste = Codigo_Tipo_Ajuste_Vac Then   '*** BONOS VAC ***
                        'dblPrecio = CalculaVANBonosVAC(strCodFile, strCodAnal, strFchCalc, dblValTir, strTipVac)
                    Else                                       '*** BONOS No VAC ***
                        
                        Set adoTemporal = New ADODB.Recordset
                        adoComm.CommandText = "SELECT dbo.uf_ACVNANoPer('" & strCodGarantia & "','" & dtpFechaLiquidacion & "','" & CVDate(lblFechaCupon.Caption) & "'," & txtMontoNominal.Value & "," & curCantidadTitulo & "," & dblValTir & ",'" & strCodTipoAjuste & "','" & strCodIndiceInicial & "','" & strCodIndiceFinal & "') AS 'ValorActual'"
                        Set adoTemporal = adoComm.Execute
                        
                        If Not adoTemporal.EOF Then
                            txtValorActual.Text = adoTemporal("ValorActual")
                        End If
                        adoTemporal.Close
                        
                        'txtValorActual.Text = VNANoPer(strCodGarantia, dtpFechaLiquidacion, CVDate(lblFechaCupon.Caption), txtMontoNominal.Value, curCantidadTitulo, dblValTir, strCodTipoAjuste, strCodIndiceInicial, strCodIndiceFinal)
                        dblPrecio = ((txtValorActual.Value) - CDbl(txtInteresCorrido(0).Value)) / (txtMontoNominal.Value)
                    End If
                '*** Calculo del PRECIO para LETRAS HIPOTECARIAS ***
                Else
                    'dblPrecio = CalculaVANLH(strCodFile, strCodAnal, strFchCalc, dblValTir)
                End If
               
                '*** PRECIO a TIR BRUTA ***
                txtPrecio(0).Text = CStr(dblPrecio * 100) 'PRECIO LIMPIO
                
                dblPrecio = ((txtValorActual.Value)) / (txtMontoNominal.Value)
                
                txtPrecioSucio.Text = CStr(dblPrecio * 100) 'PRECIO SUCIO
            End If

            'TIR BRUTA SUCIA
            'dblTir = TirNoPer(strCodGarantia, dtpFechaLiquidacion, lblFechaCupon.Caption, CDbl(txtValorActual.Value), 0, txtMontoNominal.Value, txtCantidad.Value, 0.1, strCodTipoAjuste, strCodIndiceInicial, strCodIndiceFinal)

            'TIR BRUTA LIMPIA
            'dblTir = TirNoPer(strCodGarantia, dtpFechaLiquidacion, lblFechaCupon.Caption, CDbl(txtValorActual.Value) - CDbl(txtInteresCorrido(0).Text), 0, txtMontoNominal.Value, txtCantidad.Value, 0.1, strCodTipoAjuste, strCodIndiceInicial, strCodIndiceFinal)
            
            adoComm.CommandText = "SELECT dbo.uf_ACTirNoPer('" & strCodGarantia & "','" & dtpFechaLiquidacion & "','" & lblFechaCupon.Caption & "'," & CDbl(txtValorActual.Value - CDbl(txtInteresCorrido(0).Value)) & "," & 0 & "," & txtMontoNominal.Value & "," & txtCantidad.Value & "," & 0.1 & ",'" & strCodTipoAjuste & "','" & strCodIndiceInicial & "','" & strCodIndiceFinal & "') AS 'TirNoPer'"
            Set adoTemporal = adoComm.Execute
            
            If Not adoTemporal.EOF Then
                dblTir = adoTemporal("TirNoPer")
            End If
            adoTemporal.Close
            
            txtTirBrutaLimpia.Text = CStr(dblTir)

            'TIR NETA
            'dblTir = TirNoPer(strCodGarantia, dtpFechaLiquidacion, lblFechaCupon.Caption, CCur(txtSubTotal(0).Text) + CCur(lblComisionesResumen(0).Caption), CDbl(txtInteresCorrido(0).Value) + CDbl(txtVacCorrido(0).Text), txtMontoNominal.Value, txtCantidad.Value, 0.1, strCodTipoAjuste, strCodIndiceInicial, strCodIndiceFinal)
            
            adoComm.CommandText = "SELECT dbo.uf_ACTirNoPer('" & strCodGarantia & "','" & dtpFechaLiquidacion & "','" & lblFechaCupon.Caption & "'," & CDbl(CDbl(txtSubTotal(0).Value) + CDbl(lblComisionesResumen(0).Caption)) & "," & CDbl(CDbl(txtInteresCorrido(0).Value) + CDbl(txtVacCorrido(0).Text)) & "," & txtMontoNominal.Value & "," & txtCantidad.Value & "," & 0.1 & ",'" & strCodTipoAjuste & "','" & strCodIndiceInicial & "','" & strCodIndiceFinal & "') AS 'TirNoPer'"
            Set adoTemporal = adoComm.Execute
            
            If Not adoTemporal.EOF Then
                dblTir = adoTemporal("TirNoPer")
            End If
            adoTemporal.Close
            
            txtTirNetaBaseBono.Text = CStr(dblTir)
           

        Case "TIR BRUTA"
            '*** Calculo de la TIR BRUTA(365) a partir del Precio ***
            If cboTitulo.ListIndex <= 0 Then
               MsgBox "Por favor seleccione el TITULO de la ORDEN.", vbCritical, "Aviso"
               cboTitulo.SetFocus
               Exit Sub
            End If
            
            If CDbl(txtCantidad.Text) <= 0 Then
               MsgBox "Por favor ingrese el MONTO NOMINAL de la ORDEN.", vbCritical, "Aviso"
               txtCantidad.SetFocus
               Exit Sub
            End If
            
            If CDbl(txtPrecio(0).Text) <= 0 Then
               MsgBox "Por favor ingrese el PRECIO de la ORDEN.", vbCritical, "Aviso"
               txtPrecio(0).SetFocus
               Exit Sub
            End If
            
            '*** Inicio de los calculos ***
            If CDbl(txtPrecio(0).Text) > 0 And CDbl(txtCantidad.Text) > 0 Then
                '*** Calculo de la TIR BRUTA en base 365 a partir del Precio ***
                If strCodTipoAjuste = Codigo_Tipo_Ajuste_Vac Then          '*** VAC Periodico ***
'                    dblTir = TirNoPer(strCodGarantia, dtpFechaLiquidacion, lblFechaCupon.Caption, CDbl(txtSubTotal(0).Text), CDbl(txtInteresCorrido(0).Text) - (curDifReaCap + curIntCapRea), CCur(txtCantidad.Text), CCur(lblMontoTotal(0).Caption), 0.1, strCodIndiceInicial)
                    dblTir = TirNoPer(strCodGarantia, dtpFechaLiquidacion, lblFechaCupon.Caption, CDbl(txtSubTotal(0).Value), CDbl(txtInteresCorrido(0).Value) + CDbl(txtVacCorrido(0).Text), CCur(txtCantidad.Text), CCur(lblMontoTotal(0).Caption), 0.1, strCodTipoAjuste, strCodIndiceInicial, strCodIndiceFinal)
'                ElseIf strCodIndiceInicial = Codigo_Vac_Liquidacion Then   '*** VAC al Vcto. ***
'                    dblTir = TirNoPer(strCodGarantia, dtpFechaLiquidacion, lblFechaCupon.Caption, CDbl(txtSubTotal(0).Text), CDbl(txtInteresCorrido(0).Text) - curIntCapRea, CCur(txtCantidad.Text), curCantidadTitulo, 0.1, strCodIndiceInicial)
                Else                         '*** No VAC ***
                    dblTir = TirNoPer(strCodGarantia, dtpFechaLiquidacion, CVDate(lblFechaCupon.Caption), CDbl(txtSubTotal(0).Value), CDbl(txtInteresCorrido(0).Value), CCur(txtCantidad.Text), curCantidadTitulo, 0.1, strCodTipoAjuste, strCodIndiceInicial, strCodIndiceFinal)
                End If
        
                '*** TIR BRUTA a 365 ***
                txtTirBrutaBase365.Text = CStr(dblTir)
                lblTirBrutaResumen.Caption = CStr(dblTir)
                If strCodTipoOrden = Codigo_Orden_Pacto Then lblTirBrutaResumen.Caption = "0"
               
                '*** Calculo de la TIR BRUTA en la base del Bono ***
                txtTirBrutaBaseBono.Text = "0"
                If CDbl(txtTirBrutaBase365.Text) > 0 Then
                  '*** Base del Bono es 360 ====> Convertir a 360 ***
                  If strCodBaseAnual <> Codigo_Base_Actual_Actual Or strCodBaseAnual <> Codigo_Base_30_365 Or strCodBaseAnual <> Codigo_Base_Actual_365 Then
                     txtTirBrutaBaseBono.Text = CStr(((1 + CDbl(txtTirBrutaBase365.Text) * 0.01) ^ (360 / 365) - 1) * 100)
                  '*** Base de Bono es 365 ====> Mantener en 365 ***
                  Else
                     txtTirBrutaBaseBono.Text = CStr(txtTirBrutaBase365.Text)
                  End If
               End If
            End If
        Case "TIR NETA"
            '*** Calculo de la TIR NETA(365) a partir del PRECIO ***
            If cboTitulo.ListIndex <= 0 Then
               MsgBox "Por favor seleccione el TITULO de la ORDEN.", vbCritical, Me.Caption
               cboTitulo.SetFocus
               Exit Sub
            End If
            
            If CDbl(txtCantidad.Text) <= 0 Then
               MsgBox "Por favor ingrese el MONTO NOMINAL de la ORDEN.", vbCritical, Me.Caption
               txtCantidad.SetFocus
               Exit Sub
            End If
            
            If CDbl(txtPrecio(0).Text) <= 0 Then
               MsgBox "Por favor ingrese el PRECIO de la ORDEN.", vbCritical, Me.Caption
               txtPrecio(0).SetFocus
               Exit Sub
            End If
    
            '*** Inicio de los calculos ***
            If CDbl(txtPrecio(0).Text) > 0 And CDbl(txtCantidad.Text) > 0 Then
                '*** Calculo de la TIR NETA en base 365 a partir del Precio ***
                If strCodTipoAjuste = Codigo_Tipo_Ajuste_Vac Then         '*** VAC Periodico ***
                    If strCodTipoOrden = Codigo_Orden_Venta Then
                        dblTir = TirNoPer(strCodGarantia, dtpFechaLiquidacion, lblFechaCupon.Caption, CDbl(txtSubTotal(0).Value) - CCur(lblComisionesResumen(0).Caption), CDbl(txtInteresCorrido(0).Value) + CDbl(txtVacCorrido(0).Text), CCur(txtCantidad.Text), CCur(lblMontoTotal(0).Caption), 0.1, strCodTipoAjuste, strCodIndiceInicial, strCodIndiceFinal)
                    Else
'                        dblTir = TirNoPer(strCodGarantia, dtpFechaLiquidacion, lblFechaCupon.Caption, CCur(txtSubTotal(0).Text) + CCur(lblComisionesResumen(0).Caption), CDbl(txtInteresCorrido(0).Text) - (curDifReaCap + curIntCapRea), CCur(txtCantidad.Text), CCur(lblMontoTotal(0).Caption), 0.1, strCodIndiceInicial)
                        dblTir = TirNoPer(strCodGarantia, dtpFechaLiquidacion, lblFechaCupon.Caption, CCur(txtSubTotal(0).Value) + CCur(lblComisionesResumen(0).Caption), CDbl(txtInteresCorrido(0).Value) + CDbl(txtVacCorrido(0).Text), CCur(txtCantidad.Text), CCur(lblMontoTotal(0).Caption), 0.1, strCodTipoAjuste, strCodIndiceInicial, strCodIndiceFinal)
                    End If
'                ElseIf strCodIndiceInicial = Codigo_Vac_Liquidacion Then   '*** VAC al Vcto. ***
'                    If strCodTipoOrden = Codigo_Orden_Venta Then
'                        dblTir = TirNoPer(strCodGarantia, dtpFechaLiquidacion, lblFechaCupon.Caption, CDbl(txtSubTotal(0).Text) - CCur(lblComisionesResumen(0).Caption), CDbl(txtInteresCorrido(0).Text) - curIntCapRea, CCur(txtCantidad.Text), curCantidadTitulo, 0.1, strCodTipoAjuste, strCodIndiceInicial, strCodIndiceFinal)
'                    Else
'                        dblTir = TirNoPer(strCodGarantia, dtpFechaLiquidacion, lblFechaCupon.Caption, CDbl(txtSubTotal(0).Text) + CCur(lblComisionesResumen(0).Caption), CDbl(txtInteresCorrido(0).Text) - curIntCapRea, CCur(txtCantidad.Text), curCantidadTitulo, 0.1, strCodTipoAjuste, strCodIndiceInicial, strCodIndiceFinal)
'                    End If
                Else                         '*** No VAC ***
                    If strCodTipoOrden = Codigo_Orden_Venta Then
                        dblTir = TirNoPer(strCodGarantia, dtpFechaLiquidacion, CVDate(lblFechaCupon.Caption), CDbl(txtSubTotal(0).Value) - CCur(lblComisionesResumen(0).Caption), CDbl(txtInteresCorrido(0).Value), CCur(txtCantidad.Text), curCantidadTitulo, 0.1, strCodTipoAjuste, strCodIndiceInicial, strCodIndiceFinal)
                    Else
                        dblTir = TirNoPer(strCodGarantia, dtpFechaLiquidacion, CVDate(lblFechaCupon.Caption), CDbl(txtSubTotal(0).Value) + CCur(lblComisionesResumen(0).Caption), CDbl(txtInteresCorrido(0).Value), CCur(txtCantidad.Text), curCantidadTitulo, 0.1, strCodTipoAjuste, strCodIndiceInicial, strCodIndiceFinal)
                    End If
                End If
        
                '*** TIR NETA a 365 ***
                txtTirNetaBase365.Text = CStr(dblTir)
                lblTirNetaResumen.Caption = CStr(dblTir)
                If strCodTipoOrden = Codigo_Orden_Pacto Then lblTirNetaResumen.Caption = "0"
               
                '*** Calculo de la TIR NETA en la base del Bono ***
                txtTirNetaBaseBono.Text = "0"
                If CDbl(txtTirNetaBase365.Text) > 0 Then
                  '*** Base del Bono es 360 ====> Convertir a 360 ***
                  If strCodBaseAnual <> Codigo_Base_Actual_Actual Or strCodBaseAnual <> Codigo_Base_30_365 Or strCodBaseAnual <> Codigo_Base_Actual_365 Then
                     txtTirNetaBaseBono.Text = CStr(((1 + CDbl(txtTirNetaBase365.Text) * 0.01) ^ (360 / 365) - 1) * 100)
                  '*** Base de Bono es 365 ====> Mantener en 365 ***
                  Else
                     txtTirNetaBaseBono.Text = CStr(txtTirNetaBase365.Text)
                  End If
               End If
            End If
    End Select
    
End Sub


Private Sub cboClaseInstrumento_Click()

    strCodClaseInstrumento = Valor_Caracter
    If cboClaseInstrumento.ListIndex < 0 Then Exit Sub
    
    strCodClaseInstrumento = Trim(arrClaseInstrumento(cboClaseInstrumento.ListIndex))
    
    cboTipoOrden_Click
            
End Sub


Private Sub cboConceptoCosto_Click()

    Dim adoRegistro As ADODB.Recordset
    
    strCodConcepto = Valor_Caracter
    If cboConceptoCosto.ListIndex < 0 Then Exit Sub
    
    strCodConcepto = Trim(arrConceptoCosto(cboConceptoCosto.ListIndex))
    
    If cboConceptoCosto.ListIndex > 0 Then chkAplicar(0).Enabled = True
    
    strCodTipoCostoBolsa = Valor_Caracter: strCodTipoCostoConasev = Valor_Caracter
    strCodTipoCavali = Valor_Caracter: strCodTipoCostoFondo = Valor_Caracter
    dblComisionBolsa = 0: dblComisionConasev = 0
    dblComisionCavali = 0: dblComisionFondo = 0
        
    With adoComm
        Set adoRegistro = New ADODB.Recordset
                
        .CommandText = "SELECT CodCosto,TipoCosto,ValorCosto FROM CostoNegociacion WHERE TipoOperacion='" & strCodConcepto & "' AND TipoValor='" & Codigo_Valor_RentaFija & "' ORDER BY CodCosto"
        Set adoRegistro = .Execute

        Do Until adoRegistro.EOF
            Select Case Trim(adoRegistro("CodCosto"))
                Case Codigo_Costo_Bolsa
                    strCodTipoCostoBolsa = Trim(adoRegistro("TipoCosto"))
                    dblComisionBolsa = CDbl(adoRegistro("ValorCosto"))
                Case Codigo_Costo_Conasev
                    strCodTipoCostoConasev = Trim(adoRegistro("TipoCosto"))
                    dblComisionConasev = CDbl(adoRegistro("ValorCosto"))
                Case Codigo_Costo_Cavali
                    strCodTipoCavali = Trim(adoRegistro("TipoCosto"))
                    dblComisionCavali = CDbl(adoRegistro("ValorCosto"))
                Case Codigo_Costo_FLiquidacion
                    strCodTipoCostoFondo = Trim(adoRegistro("TipoCosto"))
                    dblComisionFondo = CDbl(adoRegistro("ValorCosto"))
           End Select
           adoRegistro.MoveNext
        Loop
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
End Sub


Private Sub cboEstado_Click()

    strCodEstado = Valor_Caracter
    If cboEstado.ListIndex < 0 Then Exit Sub
    
    strCodEstado = Trim(arrEstado(cboEstado.ListIndex))
    
End Sub


Private Sub cboFondo_Click()

    Dim adoRegistro As ADODB.Recordset
    
    strCodFondo = Valor_Caracter
    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(arrFondo(cboFondo.ListIndex))
    
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Fecha Vigente, Moneda ***
        .CommandText = "{ call up_ACSelDatosParametro(23,'" & strCodFondo & "','" & gstrCodAdministradora & "') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            gdatFechaActual = CVDate(adoRegistro("FechaCuota"))
            dtpFechaOrdenDesde.Value = gdatFechaActual
            dtpFechaOrdenHasta.Value = dtpFechaOrdenDesde.Value
            strCodMoneda = Trim(adoRegistro("CodMoneda"))
                       
            frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    '*** Tipo de Instrumento ***
    strSQL = "SELECT FIF.CodFile CODIGO,DescripFile DESCRIP " & _
        "FROM FondoInversionFile FIF JOIN InversionFile IVF ON(IVF.CodFile=FIF.CodFile) " & _
        "WHERE TipoValor='" & Codigo_Valor_RentaFija & "' AND TipoPlazo='" & Codigo_Valor_LargoPlazo & "' AND IndInstrumento='X' AND IndVigente='X' AND " & _
        "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' ORDER BY DescripFile"
    CargarControlLista strSQL, cboTipoInstrumento, arrTipoInstrumento(), Sel_Todos
    
    If cboTipoInstrumento.ListCount > 0 Then cboTipoInstrumento.ListIndex = 0
    
End Sub


Private Sub cboFondoOrden_Click()

    Dim adoRegistro As ADODB.Recordset
    
    strCodFondoOrden = Valor_Caracter
    If cboFondoOrden.ListIndex < 0 Then Exit Sub
    
    strCodFondoOrden = Trim(arrFondoOrden(cboFondoOrden.ListIndex))

    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Fecha Vigente, Moneda, Tipo de Cambio ***
        .CommandText = "{ call up_ACSelDatosParametro(23,'" & strCodFondoOrden & "','" & gstrCodAdministradora & "') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            gdatFechaActual = CVDate(adoRegistro("FechaCuota"))
            gstrFechaActual = Convertyyyymmdd(gdatFechaActual)
            dtpFechaOrden.Value = gdatFechaActual
            
            gstrPeriodoActual = CStr(Year(gdatFechaActual))
            gstrMesActual = Format(Month(gdatFechaActual), "00")
            
            dtpFechaLiquidacion.Value = dtpFechaOrden.Value
            strCodMoneda = Trim(adoRegistro("CodMoneda"))
            txtTipoCambio.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, dtpFechaOrden.Value, strCodMoneda, Codigo_Moneda_Local))
            If CDbl(txtTipoCambio.Text) = 0 Then txtTipoCambio.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, DateAdd("d", -1, dtpFechaOrden.Value), strCodMoneda, Codigo_Moneda_Local))
            dblTipoCambio = CDbl(txtTipoCambio.Text)
                       
            frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
            Me.Refresh
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    '*** Tipo de Instrumento ***
    strSQL = "SELECT FIF.CodFile CODIGO,DescripFile DESCRIP " & _
        "FROM FondoInversionFile FIF JOIN InversionFile IVF ON(IVF.CodFile=FIF.CodFile) " & _
        "WHERE TipoValor='" & Codigo_Valor_RentaFija & "' AND TipoPlazo='" & Codigo_Valor_LargoPlazo & "' AND IndInstrumento='X' AND IndVigente='X' AND " & _
        "CodFondo='" & strCodFondoOrden & "' AND CodAdministradora='" & gstrCodAdministradora & "' ORDER BY DescripFile"
    CargarControlLista strSQL, cboTipoInstrumentoOrden, arrTipoInstrumentoOrden(), Sel_Defecto
            
End Sub

Private Sub cboNegociacion_Click()

    strCodNegociacion = Valor_Caracter
    If cboNegociacion.ListIndex < 0 Then Exit Sub
    
    strCodNegociacion = Trim(arrNegociacion(cboNegociacion.ListIndex))
            
    cboConceptoCosto.ListIndex = -1
    If cboConceptoCosto.ListCount > 0 Then cboConceptoCosto.ListIndex = 0
    
    cboConceptoCosto.Enabled = False
    If strCodNegociacion = Codigo_Mecanismo_Rueda Then cboConceptoCosto.Enabled = True

End Sub

Private Sub cboOperacion_Click()

    strCodOperacion = Valor_Caracter
    If cboOperacion.ListIndex < 0 Then Exit Sub
    
    strCodOperacion = Trim(arrOperacion(cboOperacion.ListIndex))
    
End Sub

Private Sub cboOrigen_Click()

    strCodOrigen = Valor_Caracter
    If cboOrigen.ListIndex < 0 Then Exit Sub
    
    strCodOrigen = Trim(arrOrigen(cboOrigen.ListIndex))
    
End Sub

Private Sub cboTipoInstrumento_Click()

    strCodTipoInstrumento = Valor_Caracter
    If cboTipoInstrumento.ListIndex < 0 Then Exit Sub
    
    strCodTipoInstrumento = Trim(arrTipoInstrumento(cboTipoInstrumento.ListIndex))
    
End Sub

Private Sub cboTipoInstrumentoOrden_Click()

'    Dim adoRegistro As ADODB.Recordset
    
    strCodTipoInstrumentoOrden = Valor_Caracter
    strIndPacto = Valor_Caracter: strIndNegociable = Valor_Caracter
    If cboTipoInstrumentoOrden.ListIndex < 0 Then Exit Sub
    
    strCodTipoInstrumentoOrden = Trim(arrTipoInstrumentoOrden(cboTipoInstrumentoOrden.ListIndex))

'    Set adoRegistro = New ADODB.Recordset
'    With adoComm
'        .CommandText = "SELECT IndPacto,IndNegociable FROM InversionFile WHERE CodFile='" & strCodTipoInstrumentoOrden & "'"
'        Set adoRegistro = .Execute
'
'        If Not adoRegistro.EOF Then
'            strIndPacto = Trim(adoRegistro("IndPacto"))
'            strIndNegociable = Trim(adoRegistro("IndNegociable"))
'        End If
'        adoRegistro.Close: Set adoRegistro = Nothing
'    End With
    
    '*** Tipo de Orden ***
    strSQL = "SELECT IFTON.CodTipoOperacion CODIGO,DescripTipoOperacion DESCRIP " & _
        "FROM InversionFileTipoOperacionNegociacion IFTON JOIN TipoOperacionNegociacion TON ON(TON.CodTipoOperacion=IFTON.CodTipoOperacion)" & _
        "WHERE IFTON.CodFile='" & strCodTipoInstrumentoOrden & "' ORDER BY DescripTipoOperacion"
    CargarControlLista strSQL, cboTipoOrden, arrTipoOrden(), Valor_Caracter

    If cboTipoOrden.ListCount > 0 Then cboTipoOrden.ListIndex = 0
        
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

Private Sub cboTipoOrden_Click()

    Dim strSQL  As String
    
    strCodTipoOrden = Valor_Caracter
    If cboTipoOrden.ListIndex < 0 Then Exit Sub

    strCodTipoOrden = Trim(arrTipoOrden(cboTipoOrden.ListIndex))

    Me.MousePointer = vbHourglass
    Select Case strCodTipoOrden
        Case Codigo_Orden_Compra
            strSQL = "SELECT CodTitulo CODIGO,(Nemotecnico + ' ' + DescripTitulo)DESCRIP FROM InstrumentoInversion " & _
                "WHERE CodFile='" & strCodFile & "' AND CodDetalleFile='" & strCodClaseInstrumento & "' AND IndVigente='X' ORDER BY DescripTitulo"
            CargarControlLista strSQL, cboTitulo, arrTitulo(), Sel_Defecto
        
            If cboTitulo.ListCount > 0 Then cboTitulo.ListIndex = 0
                                    
            txtTasaMensual.Enabled = False
            fraComisionMontoFL2.Visible = False
            
        Case Codigo_Orden_Venta
            strSQL = "SELECT II.CodTitulo CODIGO," & _
                "(RTRIM(II.Nemotecnico) + ' ' + RTRIM(II.DescripTitulo)) DESCRIP " & _
                "FROM InstrumentoInversion II JOIN InversionKardex IK ON(IK.CodTitulo=II.CodTitulo) " & _
                "WHERE SaldoFinal > 0 AND IndUltimoMovimiento='X' AND " & _
                "II.CodFile='" & strCodFile & "' AND II.CodDetalleFile='" & strCodClaseInstrumento & "' AND " & _
                "IK.CodFondo='" & strCodFondoOrden & "' AND IK.CodAdministradora='" & gstrCodAdministradora & "' " & _
                "ORDER BY II.Nemotecnico"
            CargarControlLista strSQL, cboTitulo, arrTitulo(), Sel_Defecto
        
            If cboTitulo.ListCount > 0 Then cboTitulo.ListIndex = 0
                                    
            txtTasaMensual.Enabled = False
            fraComisionMontoFL2.Visible = False
            
        Case Codigo_Orden_Pacto
            strSQL = "SELECT CodTitulo CODIGO,(Nemotecnico + ' ' + DescripTitulo)DESCRIP FROM InstrumentoInversion " & _
                "WHERE CodFile='" & strCodFile & "' AND CodDetalleFile='" & strCodClaseInstrumento & "' AND IndVigente='X' ORDER BY DescripTitulo"
            CargarControlLista strSQL, cboTitulo, arrTitulo(), Sel_Defecto
        
            If cboTitulo.ListCount > 0 Then cboTitulo.ListIndex = 0
            
            txtTasaMensual.Enabled = True
            fraComisionMontoFL2.Visible = True
                            
    End Select
    Me.MousePointer = vbDefault
    
End Sub

Private Sub cboTitulo_Click()

    Dim adoRegistro     As ADODB.Recordset
    Dim adoTemporal     As ADODB.Recordset
    Dim strIndAjuste    As String
    Dim intRegistro     As Integer
    
    strCodGarantia = Valor_Caracter: txtDescripOrden = Valor_Caracter: strCodAnalitica = Valor_Caracter
    lblAnalitica = strCodTipoInstrumentoOrden & "-????????"
    strCodEmisor = Valor_Caracter: strCodGrupo = Valor_Caracter
    strNemonico = Valor_Caracter: strIndAmortizacion = Valor_Caracter
    
    If cboTitulo.ListIndex < 0 Then Exit Sub

    strCodGarantia = Trim(arrTitulo(cboTitulo.ListIndex))

    With adoComm
        Set adoRegistro = New ADODB.Recordset

        .CommandText = "SELECT CodAnalitica,CodSubDetalleFile,ValorNominal,CodMoneda,CodEmisor,CodGrupo," & _
                        "BaseAnual,CodRiesgo,CodSubRiesgo,FechaEmision,IndTasaAjustada," & _
                        "CuponCalculo,PeriodoPago,CodTipoVac,BaseAnual,CodTipoTasa,Nemotecnico," & _
                        "IndAmortizacion,CodTipoAjuste " & _
            "FROM InstrumentoInversion WHERE CodTitulo='" & strCodGarantia & "'"
        Set adoRegistro = .Execute

        If Not adoRegistro.EOF Then
            strCodAnalitica = Trim(adoRegistro("CodAnalitica"))
            strCodMonedaGarantia = Trim(adoRegistro("CodMoneda"))
            strCodEmisor = Trim(adoRegistro("CodEmisor")): strCodGrupo = Trim(adoRegistro("CodGrupo"))
            datFechaEmision = adoRegistro("FechaEmision")
            strCodIndiceFinal = Trim(adoRegistro("CuponCalculo"))
            strCodPeriodoPago = Trim(adoRegistro("PeriodoPago"))
            strCodIndiceInicial = Trim(adoRegistro("CodTipoVac"))
            strCodTipoTasa = Trim(adoRegistro("CodTipoTasa"))
            strNemonico = Trim(adoRegistro("Nemotecnico"))
            strCodRiesgo = Trim(adoRegistro("CodRiesgo"))
            strCodSubRiesgo = Trim(adoRegistro("CodSubRiesgo"))
            strIndAmortizacion = Trim(adoRegistro("IndAmortizacion"))
            strCodSubClaseInstrumento = Trim(adoRegistro("CodSubDetalleFile"))
            
            lblAnalitica = strCodTipoInstrumentoOrden & "-" & strCodAnalitica
            If strCodTipoOrden = Codigo_Orden_Pacto Then lblAnalitica = "009" & "-????????"
            
            lblBaseTasaCupon.Caption = "360"
            intBaseAnual = 360
            If adoRegistro("BaseAnual") = Codigo_Base_Actual_Actual Or adoRegistro("BaseAnual") = Codigo_Base_Actual_365 Or adoRegistro("BaseAnual") = Codigo_Base_30_365 Then
                lblBaseTasaCupon.Caption = "365"
                intBaseAnual = 365
            End If
            strCodBaseAnual = Trim(adoRegistro("BaseAnual"))
                                    
            strCodTipoAjuste = Valor_Caracter
            strIndAjuste = Valor_Caracter
            If Trim(adoRegistro("IndTasaAjustada")) = Valor_Indicador Then
                strCodTipoAjuste = adoRegistro("CodTipoAjuste")
                strIndAjuste = Valor_Indicador
            End If
            lblValorNominal.Caption = CStr(adoRegistro("ValorNominal"))
            
            Set adoTemporal = New ADODB.Recordset
            .CommandText = "SELECT dbo.uf_IVObtenerValorNominalCupon('" & strCodGarantia & "','" & gstrFechaActual & "') AS 'ValorNominal'"
            Set adoTemporal = .Execute
            
            If Not adoTemporal.EOF Then
                lblValorNominal.Caption = adoTemporal("ValorNominal")
'            Else
'                lblValorNominal.Caption = "0.00"
            End If
            adoTemporal.Close
            
            lblMoneda.Caption = ObtenerDescripcionMoneda(strCodMonedaGarantia)
            lblMoneda.Tag = strCodMonedaGarantia
            lblDescripMoneda(0).Caption = ObtenerSignoMoneda(strCodMonedaGarantia)
            lblDescripMoneda(1).Caption = lblDescripMoneda(0).Caption
            lblDescripMonedaResumen(0).Caption = lblDescripMoneda(0).Caption
            lblDescripMonedaResumen(1).Caption = lblDescripMoneda(0).Caption
            lblClasificacion.Caption = Valor_Caracter
            tabReporte.TabEnabled(2) = True
        End If
        adoRegistro.Close
        
        '*** Obtener el Riesgo ***
        .CommandText = "SELECT ValorParametro FROM AuxiliarParametro WHERE CodTipoParametro='TIPRIE' AND CodParametro='" & strCodRiesgo & "'"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            strCodCategoria = Trim(adoRegistro("ValorParametro"))
        End If
        adoRegistro.Close
        
        lblClasificacion.Caption = strCodCategoria & Space(1) & strCodSubRiesgo
        
        '*** Obtener kardex de titulos ***
        .CommandText = "SELECT TirPromedio,SaldoFinal,SaldoAmortizacion FROM InversionKardex " & _
            "WHERE CodTitulo='" & strCodGarantia & "' AND SaldoFinal > 0 AND IndUltimoMovimiento='X' AND " & _
            "CodFondo='" & strCodFondoOrden & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
        Set adoRegistro = .Execute

        If Not adoRegistro.EOF Then
            lblStockNominal.Caption = CStr(adoRegistro("SaldoFinal"))
            curSaldoAmortizacion = CCur(adoRegistro("SaldoAmortizacion"))
        End If
        adoRegistro.Close
        
        .CommandText = "SELECT FechaInicio,FechaVencimiento,TasaInteres,ValorAmortizacion,FechaInicioIndice,FechaFinIndice " & _
            "FROM InstrumentoInversionCalendario " & _
            "WHERE CodTitulo='" & strCodGarantia & "' AND FechaVencimiento>='" & Convertyyyymmdd(dtpFechaLiquidacion.Value) & "' ORDER BY FechaVencimiento  "
        Set adoRegistro = .Execute

        If Not adoRegistro.EOF Then
            lblBaseTasaCupon.Caption = lblBaseTasaCupon.Caption & " - " & Format(adoRegistro("TasaInteres"), "0.000000")
            lblFechaCupon.Caption = CStr(adoRegistro("FechaVencimiento"))
            lblInicioFechaCupon.Caption = CStr(adoRegistro("FechaInicio"))
            
            '*** La tasa es ajustada ? ***
            If strIndAjuste = Valor_Indicador Then
                Dim dblVacIniCupon          As Double, dblVacFinCupon           As Double
                Dim strFechaInicuponMas1    As String, strFechaFinCuponMas1     As String
                
                If strCodTipoAjuste = Codigo_Tipo_Ajuste_Vac Then
                    lblIndiceInicial.Caption = CStr(adoRegistro("FechaInicioIndice"))
                    If strCodIndiceFinal = Codigo_Vac_Liquidacion Then
                        lblIndiceFinal.Caption = CStr(dtpFechaLiquidacion.Value)
                    Else
                        lblIndiceFinal.Caption = CStr(adoRegistro("FechaFinIndice"))
                    End If
                    
                    dblVacIniCupon = 0: dblVacFinCupon = 0
                    
                    strFechaInicuponMas1 = Convertyyyymmdd(DateAdd("d", 1, CVDate(lblIndiceInicial.Caption)))
                    strFechaFinCuponMas1 = Convertyyyymmdd(DateAdd("d", 1, CVDate(lblIndiceFinal.Caption)))
                    
                    '*** Obtener las Tasas VAC: Emisión, Liquidación, Cupón Inicial y Cupón Final ***
                    dblVacIniCupon = ObtenerTasaAjuste(Codigo_Tipo_Ajuste_Vac, "00", Convertyyyymmdd(CVDate(lblIndiceInicial.Caption)), strFechaInicuponMas1)
                    dblVacFinCupon = ObtenerTasaAjuste(Codigo_Tipo_Ajuste_Vac, "00", Convertyyyymmdd(CVDate(lblIndiceFinal.Caption)), strFechaFinCuponMas1)
                    
                    lblIndiceInicial.ToolTipText = CStr(dblVacIniCupon)
                    lblIndiceFinal.ToolTipText = CStr(dblVacFinCupon)
                Else
                    lblIndiceInicial.Caption = Valor_Caracter
                    lblIndiceFinal.Caption = CStr(dtpFechaLiquidacion.Value)
                    
                    dblVacIniCupon = 0: dblVacFinCupon = 0
                    
                    strFechaFinCuponMas1 = Convertyyyymmdd(DateAdd("d", 1, CVDate(lblIndiceFinal.Caption)))
                    
                    '*** Obtener la Tasa de Ajuste ***
                    dblVacFinCupon = ObtenerTasaAjuste(strCodTipoAjuste, strCodIndiceInicial, Convertyyyymmdd(CVDate(lblIndiceFinal.Caption)), strFechaFinCuponMas1)
                    
                    lblIndiceInicial.ToolTipText = Valor_Caracter
                    lblIndiceFinal.ToolTipText = CStr(dblVacFinCupon)
                End If
            End If
        End If
        adoRegistro.Close
        
        .CommandText = "SELECT SUM(ValorAmortizacion) ValorAmortizacion " & _
            "FROM InstrumentoInversionCalendario " & _
            "WHERE CodTitulo='" & strCodGarantia & "' AND FechaVencimiento<'" & Convertyyyymmdd(dtpFechaLiquidacion.Value) & "'"
        Set adoRegistro = .Execute

        If Not adoRegistro.EOF Then
            If Not IsNull(adoRegistro("ValorAmortizacion")) Then
                dblValorAmortizacion = CDbl(adoRegistro("ValorAmortizacion")) / CDbl(lblValorNominal.Caption)
            End If
        End If
        adoRegistro.Close: Set adoRegistro = Nothing

        If strCodGarantia <> Valor_Caracter Then
            '*** Validar Limites ***
            If Not PosicionLimites() Then Exit Sub
        End If
    End With

    txtDescripOrden.Text = UCase(Trim(cboTipoOrden.Text) & " " & Trim(cboTipoInstrumentoOrden.Text) & " - " & Left(cboTitulo.Text, 15))
    
    
End Sub

Private Sub chkAjustePrecio_Click(Index As Integer)

    txtCantidad_Change
    Call CalculoTotal(Index)
    
End Sub

Private Sub chkAplicar_Click(Index As Integer)

    If chkAplicar(Index).Value Then
        Call AplicarCostos(Index)
    Else
        Call IniciarComisiones
        Call CalculoTotal(Index)
    End If
    
End Sub

Private Sub cmdCalculo_Click()

    Dim dblFactor As Double

    '*** Tir Bruta ***
'    If CInt(txtDiasPlazo.Text) > 0 And CCur(txtSubTotal(0).Text) > 0 Then
'        If CCur(txtSubTotal(1).Text) = 0 Or CCur(txtSubTotal(0).Text) = 0 Then
'            MsgBox "Por favor verificar que el SubTotal al Contado y a Plazo tengan valores.", vbExclamation, Me.Caption
'            Exit Sub
'        End If
'        'dblFactor = (CCur(txtSubTotal(1).Text) / CCur(txtSubTotal(0).Text)) ^ (365 / CInt(txtDiasPlazo.Text))
'        dblFactor = TirNoPerPlazo(strCodGarantia, dtpFechaLiquidacion, lblFechaCupon.Caption, dtpFechaVencimiento.Value, CDbl(txtSubTotal(1).Text) + CDbl(txtInteresCorrido(1).Text), CDbl(txtSubTotal(0).Text), CDbl(txtInteresCorrido(0).Text) - curIntCapRea, CCur(txtCantidad.Text) * CCur(lblValorNominal.Caption), curCantidadTitulo * CCur(lblValorNominal.Caption), 0.1, strCodIndiceInicial, CInt(txtDiasPlazo.Text))
'        lblTirBruta.Caption = CStr(dblFactor)
'        lblTirBrutaResumen.Caption = CStr(dblFactor)
'    End If
'
'    '*** Tir Neta ***
'    If CInt(txtDiasPlazo.Text) > 0 And CCur(lblMontoTotal(0).Caption) > 0 Then
'        'dblFactor = (CCur(lblMontoTotal(1).Caption) / CCur(lblMontoTotal(0).Caption)) ^ (365 / CInt(txtDiasPlazo.Text))
'        dblFactor = TirNoPerPlazo(strCodGarantia, dtpFechaLiquidacion, lblFechaCupon.Caption, dtpFechaVencimiento.Value, CDbl(lblMontoTotal(1).Caption), CDbl(txtSubTotal(0).Text) + CDbl(lblComisionesResumen(0).Caption), CDbl(txtInteresCorrido(0).Text) - curIntCapRea, CCur(txtCantidad.Text) * CCur(lblValorNominal.Caption), curCantidadTitulo * CCur(lblValorNominal.Caption), 0.1, strCodIndiceInicial, CInt(txtDiasPlazo.Text))
'        lblTirNeta.Caption = CStr(dblFactor)
'        lblTirNetaResumen.Caption = CStr(dblFactor)
'    End If
    
End Sub

Private Sub cmdEnviar_Click()

    Dim strFechaDesde       As String, strFechaHasta        As String
    Dim intRegistro         As Integer, intContador         As Integer
    Dim datFecha            As Date
    
    If adoConsulta.RecordCount = 0 Then Exit Sub
    
    strFechaDesde = Convertyyyymmdd(dtpFechaOrdenDesde.Value)
    datFecha = DateAdd("d", 1, dtpFechaOrdenHasta.Value)
    strFechaHasta = Convertyyyymmdd(datFecha)
    
    intContador = tdgConsulta.SelBookmarks.Count - 1
    
    If intContador < 0 Then
        MsgBox "No se ha seleccionado ningún registro", vbCritical, Me.Caption
        Exit Sub
    End If
        
    For intRegistro = 0 To intContador
        tdgConsulta.Row = tdgConsulta.SelBookmarks(intRegistro) - 1
               
        If strCodEstado = Estado_Orden_Ingresada Then
            adoComm.CommandText = "UPDATE InversionOrden SET EstadoOrden='" & Estado_Orden_Enviada & "'," & _
                "UsuarioEdicion='" & gstrLogin & "',FechaEdicion='" & strFechaDesde & Space(1) & Format(Time, "hh:mm") & "' " & _
                "WHERE NumOrden='" & Trim(tdgConsulta.Columns(1)) & "' AND CodFondo='" & strCodFondo & "' AND " & _
                "CodAdministradora='" & gstrCodAdministradora & "' AND EstadoOrden='" & Estado_Orden_Ingresada & "'"
        ElseIf strCodEstado = Estado_Orden_Enviada Then
            adoComm.CommandText = "UPDATE InversionOrden SET EstadoOrden='" & Estado_Orden_Ingresada & "'," & _
                "UsuarioEdicion='" & gstrLogin & "',FechaEdicion='" & strFechaDesde & Space(1) & Format(Time, "hh:mm") & "' " & _
                "WHERE NumOrden='" & Trim(tdgConsulta.Columns(1)) & "' AND CodFondo='" & strCodFondo & "' AND " & _
                "CodAdministradora='" & gstrCodAdministradora & "' AND EstadoOrden='" & Estado_Orden_Enviada & "'"
        End If
        adoConn.Execute adoComm.CommandText
    Next
    
    If strCodEstado = Estado_Orden_Ingresada Then
        MsgBox Mensaje_Envio_Exitoso, vbExclamation, gstrNombreEmpresa
    Else
        MsgBox Mensaje_Desenvio_Exitoso, vbExclamation, gstrNombreEmpresa
    End If
    Call Buscar
    
End Sub

Private Sub cmdExportarExcel_Click()
    Call ExportarExcel
End Sub

Private Sub ExportarExcel()
    
    Dim adoRegistro As ADODB.Recordset
    Dim execSQL As String
    Dim rutaExportacion As String

    Dim datFechaSiguiente As Date
    Dim strFechaLiquidacionHasta As String

    Set frmFormulario = frmOrdenRentaFijaLargoPlazo

    Set adoRegistro = New ADODB.Recordset

    'If TodoOK() Then

        Dim strNameProc As String

        gstrNameRepo = "OrdenRentaFijaLargoPlazo"

        strNameProc = ObtenerBaseReporte(gstrNameRepo)

        Dim arrParmS(6)

        arrParmS(0) = Trim(strCodFondo)
        arrParmS(1) = Trim(gstrCodAdministradora)

        If strCodTipoInstrumento <> Valor_Caracter Then
            arrParmS(2) = Trim(strCodTipoInstrumento)
        Else
            arrParmS(2) = "%"
        End If

        If IsNull(dtpFechaOrdenDesde.Value) And IsNull(dtpFechaOrdenHasta.Value) Then
            arrParmS(3) = Convertyyyymmdd(dtpFechaLiquidacionDesde.Value)
            datFechaSiguiente = DateAdd("d", 1, dtpFechaLiquidacionHasta.Value)
            strFechaLiquidacionHasta = Convertyyyymmdd(datFechaSiguiente)
            arrParmS(4) = strFechaLiquidacionHasta
            arrParmS(5) = "L"
        Else
            arrParmS(3) = Convertyyyymmdd(dtpFechaOrdenDesde.Value)
            arrParmS(4) = Convertyyyymmdd(dtpFechaOrdenHasta.Value)
            arrParmS(5) = "O"
        End If
        
        If strCodEstado <> Valor_Caracter Then
            arrParmS(6) = strCodEstado
        Else
            MsgBox "Debe seleccionar un Estado.", vbCritical, Me.Caption
            If cboEstado.Enabled Then cboEstado.SetFocus
            Exit Sub
        End If

        execSQL = ObtenerCommandText(strNameProc, arrParmS())

        With adoComm

            .CommandText = execSQL

            Set adoRegistro = .Execute

        End With

        Set oExportacion = New clsExportacion

        Call ConfiguraRecordsetExportacion

        Call LlenarRecordsetExportacion(adoRegistro)

        If adoExportacion.RecordCount > 0 Then

            frmRutaGrabar.Show vbModal

            If indOk = True Then

                Screen.MousePointer = vbHourglass

                rutaExportacion = gs_FormName

                If oExportacion.ExportaRecordSetExcel(adoExportacion, gstrNameRepo, rutaExportacion) Then
                    MsgBox "Exportacion realizada", vbInformation
                Else
                    MsgBox "Fallo en exportacion", vbCritical
                End If

                Set oExportacion = Nothing

            End If

        Else
            MsgBox "No existen registros, exportacion a excel cancelada", vbExclamation
        End If

        Screen.MousePointer = vbDefault

    'End If
        
End Sub

Private Function ObtenerBaseReporte(ByVal strNombreReporte As String) As String
    
    ObtenerBaseReporte = Valor_Caracter
    
    Dim crxAplicacion As CRAXDRT.Application
    Dim crxReporte As CRAXDRT.Report
    Dim strReportPath As String
    Dim strBase As String
    Dim intIndex As Integer
        
    strReportPath = gstrRptPath & strNombreReporte & ".RPT"
    
    On Error GoTo Ctrl_Error
    
    Set crxAplicacion = New CRAXDRT.Application

    Set crxReporte = crxAplicacion.OpenReport(strReportPath)

    strBase = crxReporte.Database.Tables(1).Name

    intIndex = InStr(1, strBase, ";", vbBinaryCompare)
        
    strBase = Mid(strBase, 1, intIndex - 1)
    
    ObtenerBaseReporte = strBase

    Set crxReporte = Nothing
    Set crxAplicacion = Nothing
    
    Exit Function
    
Ctrl_Error:
MsgBox "Error al obtener la base del Reporte", vbCritical
Exit Function

End Function

Private Function ObtenerCommandText(ByVal strCadena As String, ByRef arrParametros()) As String
    
    Dim strParametros As String
    Dim i As Integer
    
    strParametros = "{ call " & strCadena & " ("
    
    For i = 0 To UBound(arrParametros)
    
        strParametros = strParametros & "'" & arrParametros(i) & "'" & ","
    
    Next
    
    strParametros = Mid(strParametros, 1, Len(strParametros) - 1)
    
    strParametros = strParametros & ") }"
    
    ObtenerCommandText = strParametros

End Function

Private Sub ConfiguraRecordsetExportacion()

    Set adoExportacion = New ADODB.Recordset

    With adoExportacion
       .CursorLocation = adUseClient
       .Fields.Append "NumOrden", adChar, 10
       .Fields.Append "FechaOrden", adDate
       .Fields.Append "FechaLiquidacion", adDate
       .Fields.Append "CodTitulo", adChar, 15
       .Fields.Append "Nemotecnico", adChar, 15
       .Fields.Append "EstadoOrden", adChar, 2
       
       .Fields.Append "CodFile", adChar, 3
       .Fields.Append "CodAnalitica", adChar, 8
       .Fields.Append "TipoOrden", adChar, 2
       
       .Fields.Append "CodMoneda", adChar, 2
       .Fields.Append "DescripOrden", adVarChar, 100
       .Fields.Append "CantOrden", adDecimal
       .Fields.Append "ValorNominal", adDecimal
       .Fields.Append "PrecioUnitarioMFL1", adDecimal
       .Fields.Append "MontoTotalMFL1", adDecimal
       .Fields.Append "DescripMoneda", adChar, 3
'       .CursorType = adOpenStatic

       .LockType = adLockBatchOptimistic
    End With
    
'    adoExportacion.Fields.Item("Cantidad").Precision = 19
'    adoExportacion.Fields.Item("Cantidad").NumericScale = 2
'
'    adoExportacion.Fields.Item("Cotiza").Precision = 23
'    adoExportacion.Fields.Item("Cotiza").NumericScale = 6
'
'    adoExportacion.Fields.Item("Bruto").Precision = 19
'    adoExportacion.Fields.Item("Bruto").NumericScale = 2
'
'    adoExportacion.Fields.Item("S.A.B").Precision = 19
'    adoExportacion.Fields.Item("S.A.B").NumericScale = 2
'
'    adoExportacion.Fields.Item("BVL").Precision = 19
'    adoExportacion.Fields.Item("BVL").NumericScale = 2
'
'    adoExportacion.Fields.Item("Fondo").Precision = 19
'    adoExportacion.Fields.Item("Fondo").NumericScale = 2
'
'    adoExportacion.Fields.Item("Cavali").Precision = 19
'    adoExportacion.Fields.Item("Cavali").NumericScale = 2
'
'    adoExportacion.Fields.Item("Fdo. Cavali").Precision = 19
'    adoExportacion.Fields.Item("Fdo. Cavali").NumericScale = 2
'
'    adoExportacion.Fields.Item("Conasev").Precision = 19
'    adoExportacion.Fields.Item("Conasev").NumericScale = 2
'
'    adoExportacion.Fields.Item("Com. Broker").Precision = 19
'    adoExportacion.Fields.Item("Com. Broker").NumericScale = 2
'
'    adoExportacion.Fields.Item("Tot. Com").Precision = 19
'    adoExportacion.Fields.Item("Tot. Com").NumericScale = 2
'
'    adoExportacion.Fields.Item("IGV").Precision = 19
'    adoExportacion.Fields.Item("IGV").NumericScale = 2
'
'    adoExportacion.Fields.Item("IGV Cavali").Precision = 19
'    adoExportacion.Fields.Item("IGV Cavali").NumericScale = 2
'
'    adoExportacion.Fields.Item("Neto").Precision = 19
'    adoExportacion.Fields.Item("Neto").NumericScale = 2
'
'    adoExportacion.Fields.Item("RUT").Precision = 19
'    adoExportacion.Fields.Item("RUT").NumericScale = 2
    
    adoExportacion.Open
    
End Sub

Private Sub LlenarRecordsetExportacion(ByRef adoRecords As ADODB.Recordset)
        
    Dim dblTipoCambio As Double, dblBruto As Double, dblSAB As Double, dblBVL As Double, dblFondo As Double
    Dim dblCavali As Double, dblFdoCavali As Double, dblConasev As Double, dblTotCom As Double, dblComBroker As Double
    Dim dblCotiza As Double, dblIGV As Double, dblIGVCavali As Double, dblNeto As Double
        
    'dblTipoCambio = CDbl(txtTipoCambio.Text)
        
    If Not adoRecords.EOF Then
    
        Do Until adoRecords.EOF
                
                adoExportacion.AddNew
                
                adoExportacion.Fields("NumOrden") = Trim(adoRecords.Fields("NumOrden"))
                adoExportacion.Fields("FechaOrden") = Trim(adoRecords.Fields("FechaOrden"))
                adoExportacion.Fields("FechaLiquidacion") = adoRecords.Fields("FechaLiquidacion")
                adoExportacion.Fields("CodTitulo") = Trim(adoRecords.Fields("CodTitulo"))
                adoExportacion.Fields("Nemotecnico") = Trim(adoRecords.Fields("Nemotecnico"))
                adoExportacion.Fields("EstadoOrden") = Trim(adoRecords.Fields("EstadoOrden"))
                
                adoExportacion.Fields("CodFile") = Trim(adoRecords.Fields("CodFile"))
                adoExportacion.Fields("CodAnalitica") = Trim(adoRecords.Fields("CodAnalitica"))
                adoExportacion.Fields("TipoOrden") = adoRecords.Fields("TipoOrden")
                
                adoExportacion.Fields("CodMoneda") = adoRecords.Fields("CodMoneda")
                adoExportacion.Fields("DescripOrden") = adoRecords.Fields("DescripOrden")
                adoExportacion.Fields("CantOrden") = adoRecords.Fields("CantOrden")
                adoExportacion.Fields("ValorNominal") = adoRecords.Fields("ValorNominal")
                adoExportacion.Fields("PrecioUnitarioMFL1") = adoRecords.Fields("PrecioUnitarioMFL1")
                adoExportacion.Fields("MontoTotalMFL1") = adoRecords.Fields("MontoTotalMFL1")
                adoExportacion.Fields("DescripMoneda") = adoRecords.Fields("DescripMoneda")
                
                adoExportacion.Update
    
                adoRecords.MoveNext
                
        Loop
        
        adoRecords.Close: Set adoRecords = Nothing
    
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
    
    txtCantidad_Change
    
End Sub

Private Sub dtpFechaLiquidacionDesde_Click()

    If IsNull(dtpFechaLiquidacionDesde.Value) Then
        dtpFechaLiquidacionHasta.Value = Null
        dtpFechaOrdenDesde.Value = gdatFechaActual
        dtpFechaOrdenHasta.Value = dtpFechaOrdenDesde.Value
    Else
        dtpFechaLiquidacionDesde.Value = gdatFechaActual
        dtpFechaLiquidacionHasta.Value = dtpFechaLiquidacionDesde.Value
        dtpFechaOrdenDesde.Value = Null
        dtpFechaOrdenHasta.Value = dtpFechaOrdenDesde.Value
    End If
    
End Sub


Private Sub dtpFechaLiquidacionHasta_Click()

    If IsNull(dtpFechaLiquidacionHasta.Value) Then
        dtpFechaLiquidacionDesde.Value = Null
        dtpFechaOrdenDesde.Value = gdatFechaActual
        dtpFechaOrdenHasta.Value = dtpFechaOrdenDesde.Value
    Else
        dtpFechaLiquidacionDesde.Value = gdatFechaActual
        dtpFechaLiquidacionHasta.Value = dtpFechaLiquidacionDesde.Value
        dtpFechaOrdenDesde.Value = Null
        dtpFechaOrdenHasta.Value = dtpFechaOrdenDesde.Value
    End If
    
End Sub


Private Sub dtpFechaOrdenDesde_Click()

    If IsNull(dtpFechaOrdenDesde.Value) Then
        dtpFechaOrdenHasta.Value = Null
        dtpFechaLiquidacionDesde.Value = gdatFechaActual
        dtpFechaLiquidacionHasta.Value = dtpFechaLiquidacionDesde.Value
    Else
        dtpFechaOrdenDesde.Value = gdatFechaActual
        dtpFechaOrdenHasta.Value = dtpFechaOrdenDesde.Value
        dtpFechaLiquidacionDesde.Value = Null
        dtpFechaLiquidacionHasta.Value = dtpFechaLiquidacionDesde.Value
    End If
    
End Sub


Private Sub dtpFechaOrdenHasta_Click()

    If IsNull(dtpFechaOrdenHasta.Value) Then
        dtpFechaOrdenDesde.Value = Null
        dtpFechaLiquidacionDesde.Value = gdatFechaActual
        dtpFechaLiquidacionHasta.Value = dtpFechaLiquidacionDesde.Value
    Else
        dtpFechaOrdenDesde.Value = gdatFechaActual
        dtpFechaOrdenHasta.Value = dtpFechaOrdenDesde.Value
        dtpFechaLiquidacionDesde.Value = Null
        dtpFechaLiquidacionHasta.Value = dtpFechaLiquidacionDesde.Value
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

Private Sub Form_Unload(Cancel As Integer)

    Set frmOrdenReporteRentaVariable = Nothing
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

Private Sub lblPorcenBolsa_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblPorcenBolsa(Index), Decimales_Tasa)
    
End Sub

Private Sub lblPorcenCavali_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblPorcenCavali(Index), Decimales_Tasa)
    
End Sub

Private Sub lblPorcenConasev_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblPorcenConasev(Index), Decimales_Tasa)
    
End Sub

Private Sub lblPorcenFondo_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblPorcenFondo(Index), Decimales_Tasa)
    
End Sub

Private Sub lblPorcenIgv_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblPorcenIgv(Index), Decimales_Monto)
    
End Sub

Private Sub lblPrecioResumen_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblPrecioResumen(Index), Decimales_Precio)
    
End Sub

Private Sub lblStockNominal_Change()

    Call FormatoMillarEtiqueta(lblStockNominal, Decimales_Monto)
    
End Sub

Private Sub lblSubTotalResumen_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblSubTotalResumen(Index), Decimales_Monto)
    
End Sub

Private Sub lblTirBruta_Change()

    Call FormatoMillarEtiqueta(lblTirBruta, Decimales_TasaDiaria)
    
End Sub

Private Sub lblTirBrutaResumen_Change()

    Call FormatoMillarEtiqueta(lblTirBrutaResumen, Decimales_TasaDiaria)
    
End Sub

Private Sub lblTirNeta_Change()

    Call FormatoMillarEtiqueta(lblTirNeta, Decimales_TasaDiaria)
    
End Sub

Private Sub lblTirNetaResumen_Change()

    Call FormatoMillarEtiqueta(lblTirNetaResumen, Decimales_TasaDiaria)
    
End Sub

Private Sub lblTotalResumen_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblTotalResumen(Index), Decimales_Monto)
    
End Sub

Private Sub lblValorNominal_Change()

    Call FormatoMillarEtiqueta(lblValorNominal, Decimales_TasaDiaria)
    
End Sub

Private Sub tabReporte_Click(PreviousTab As Integer)

    Select Case tabReporte.Tab
        Case 1, 2
            'If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vQuery)
            If PreviousTab = 0 And strEstado = Reg_Consulta Then tabReporte.Tab = 0
            If strEstado = Reg_Defecto Then tabReporte.Tab = 0
            If tabReporte.Tab = 2 Then
                fraDatosNegociacion.Caption = "Negociación" & Space(1) & "-" & Space(1) & _
                    Trim(cboTipoOrden.Text) & Space(1) & Trim(Left(cboTitulo.Text, 15))
            End If
    End Select
    
End Sub

Private Sub TAMTextBox1_Click()

End Sub

Private Sub tdgConsulta_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)

    If ColIndex = 4 Then
        Call DarFormatoValor(Value, Decimales_Monto)
    End If
    
    If ColIndex = 6 Then
        Call DarFormatoValor(Value, Decimales_Precio)
    End If
    
    If ColIndex = 7 Then
        Call DarFormatoValor(Value, Decimales_Monto)
    End If
    
End Sub

Private Sub txtCantidad_Change()

    Dim curCantidad As Currency, dblPreUni  As Double
    Dim curSubTotal As Currency
    
    If Trim(txtCantidad.Text) = Valor_Caracter Then Exit Sub
    
    If txtCantidad.Value > 0 And cboTitulo.ListIndex > 0 And Not blnMonto Then
        
        blnCantidad = True
        
        If IsNumeric(txtCantidad.Text) Then
            curCantidad = CCur(txtCantidad.Value)
        Else
            curCantidad = 0
        End If
        
        lblCantidadResumen.Caption = CStr(curCantidad)
        
        txtMontoNominal.Text = curCantidad * CDbl(lblValorNominal.Caption)
        
        curCantidadTitulo = curCantidad
            
        txtInteresCorrido(0).Text = "0"
        
        txtInteresCorrido(0).Text = CStr(CalculoInteresCorrido(strCodGarantia, txtMontoNominal.Value, CVDate(lblInicioFechaCupon.Caption), dtpFechaLiquidacion.Value, strCodIndiceFinal, strCodTipoAjuste, strCodTipoTasa, strCodPeriodoPago, strCodIndiceInicial, strCodBaseAnual, intBaseAnual))
        
        Call CalcularPrecioTIRVAN
                
        Call CalculoTotal(0)

        blnCantidad = False
    
    End If
    
End Sub




Private Sub txtComisionAgente_Change(Index As Integer)

    Call FormatoCajaTexto(txtComisionAgente(Index), Decimales_Monto)
    
    If chkAplicar(Index).Value Then
        ActualizaPorcentaje txtComisionAgente(Index), txtPorcenAgente(Index)
    End If
    
    Call CalculoTotal(Index)
    
End Sub

Private Sub txtComisionAgente_KeyPress(Index As Integer, KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtComisionAgente(Index), Decimales_Monto)
    
    If KeyAscii = vbKeyReturn Then
        If chkAplicar(Index).Value Then
            ActualizaPorcentaje txtComisionAgente(Index), txtPorcenAgente(Index)
        End If
        Call CalculoTotal(Index)
    End If
    
End Sub

Private Sub txtComisionBolsa_Change(Index As Integer)

    Call FormatoCajaTexto(txtComisionBolsa(Index), Decimales_Monto)
    
    If chkAplicar(Index).Value Then
        ActualizaPorcentaje txtComisionBolsa(Index), lblPorcenBolsa(Index)
    End If
    
    Call CalculoTotal(Index)
    
End Sub

Private Sub txtComisionBolsa_KeyPress(Index As Integer, KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtComisionBolsa(Index), Decimales_Monto)
    
    If KeyAscii = vbKeyReturn Then
        Call CalculoTotal(Index)
    End If
    
End Sub

Private Sub txtComisionCavali_Change(Index As Integer)

    Call FormatoCajaTexto(txtComisionCavali(Index), Decimales_Monto)
    
    If chkAplicar(Index).Value Then
        ActualizaPorcentaje txtComisionCavali(Index), lblPorcenCavali(Index)
    End If
    
    Call CalculoTotal(Index)
    
End Sub

Private Sub txtComisionCavali_KeyPress(Index As Integer, KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtComisionCavali(Index), Decimales_Monto)
    
    If KeyAscii = vbKeyReturn Then
        Call CalculoTotal(Index)
    End If
    
End Sub

Private Sub txtComisionConasev_Change(Index As Integer)

    Call FormatoCajaTexto(txtComisionConasev(Index), Decimales_Monto)
    
    If chkAplicar(Index).Value Then
        ActualizaPorcentaje txtComisionConasev(Index), lblPorcenConasev(Index)
    End If
    
    Call CalculoTotal(Index)
    
End Sub

Private Sub txtComisionConasev_KeyPress(Index As Integer, KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtComisionConasev(Index), Decimales_Monto)
    
    If KeyAscii = vbKeyReturn Then
        Call CalculoTotal(Index)
    End If
    
End Sub

Private Sub txtComisionFondo_Change(Index As Integer)

    Call FormatoCajaTexto(txtComisionFondo(Index), Decimales_Monto)
    
    If chkAplicar(Index).Value Then
        ActualizaPorcentaje txtComisionFondo(Index), lblPorcenFondo(Index)
    End If
    
    Call CalculoTotal(Index)
    
End Sub


Private Sub txtComisionFondo_KeyPress(Index As Integer, KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtComisionFondo(Index), Decimales_Monto)
    
    If KeyAscii = vbKeyReturn Then
        Call CalculoTotal(Index)
    End If
    
End Sub

Private Sub txtInteresCorrido_Change(Index As Integer)

    'Call FormatoCajaTexto(txtInteresCorrido(Index), Decimales_Monto)
    
    If Trim(txtInteresCorrido(Index).Text) <> Valor_Caracter And Trim(txtVacCorrido(Index).Text) <> Valor_Caracter Then
        lblInteresesResumen(Index).Caption = CStr(CCur(txtInteresCorrido(Index).Text) + CCur(txtVacCorrido(Index).Text))
    End If
    
End Sub

Private Sub txtInteresCorrido_KeyPress(Index As Integer, KeyAscii As Integer)

    'Call ValidaCajaTexto(KeyAscii, "M", txtInteresCorrido(Index), Decimales_Monto)
    
    If KeyAscii = vbKeyReturn Then
        Call CalculoTotal(Index)
    End If
    
End Sub


Private Sub txtMontoNominal_Change()

    Dim curMonto As Currency, dblPreUni  As Double
    Dim curSubTotal As Currency

    If Trim(txtMontoNominal.Text) = Valor_Caracter Then Exit Sub
    
    If txtMontoNominal.Value > 0 And cboTitulo.ListIndex > 0 And Not blnCantidad Then
        
        blnMonto = True
        
        If IsNumeric(txtMontoNominal.Value) Then
            curMonto = CCur(txtMontoNominal.Value)
        Else
            curMonto = 0
        End If
        
        txtCantidad.Text = CStr(Round(txtMontoNominal.Value / CDbl(lblValorNominal.Caption), 0))
        
        lblCantidadResumen.Caption = CStr(txtCantidad.Value)
        
        curCantidadTitulo = txtCantidad.Value
            
        txtInteresCorrido(0).Text = "0"
        
        txtInteresCorrido(0).Text = CStr(CalculoInteresCorrido(strCodGarantia, txtMontoNominal.Value, CVDate(lblInicioFechaCupon.Caption), dtpFechaLiquidacion.Value, strCodIndiceFinal, strCodTipoAjuste, strCodTipoTasa, strCodPeriodoPago, strCodIndiceInicial, strCodBaseAnual, intBaseAnual))
        
        Call CalcularPrecioTIRVAN
        
        Call CalculoTotal(0)

        blnMonto = False
    
    End If
    
End Sub

Private Sub txtPorcenAgente_Change(Index As Integer)

    Call FormatoCajaTexto(txtPorcenAgente(Index), Decimales_Tasa)
    
End Sub

Private Sub txtPorcenAgente_KeyPress(Index As Integer, KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtPorcenAgente(Index), Decimales_Tasa)
    
    If KeyAscii = vbKeyReturn Then
        If chkAplicar(Index).Value Then
            ActualizaComision txtPorcenAgente(Index), txtComisionAgente(Index)
        End If
        Call CalculoTotal(Index)
    End If
    
End Sub

Private Sub txtPrecio_Change(Index As Integer)

    'Call FormatoCajaTexto(txtPrecio(index), Decimales_Precio)
    
    txtCantidad_Change
    Call CalculoTotal(Index)
    
    lblPrecioResumen(Index).Caption = CStr(txtPrecio(Index).Text)
    
End Sub

Private Sub txtPrecio_KeyPress(Index As Integer, KeyAscii As Integer)

    'Call ValidaCajaTexto(KeyAscii, "M", txtPrecio(index), Decimales_Precio)
    
End Sub

Private Sub txtSubTotal_Change(Index As Integer)

    'Call FormatoCajaTexto(txtSubTotal(Index), Decimales_Monto)
    
'    If CLng(txtCantidad.Text) > 0 And CCur(txtSubTotal(Index).Text) > 0 Then
'        txtPrecio(Index).Text = CStr(CCur(txtSubTotal(Index).Text) / (CLng(txtCantidad.Text) * CLng(lblValorNominal.Caption)))
'    End If
    
    lblSubTotalResumen(Index).Caption = CStr(txtSubTotal(Index).Value)
    
End Sub

Private Sub txtSubTotal_KeyPress(Index As Integer, KeyAscii As Integer)
    
    'Call ValidaCajaTexto(KeyAscii, "M", txtSubTotal(Index), Decimales_Monto)
    
    If KeyAscii = vbKeyReturn Then
        Call CalculoTotal(Index)
    End If
    
End Sub

Private Sub txtTasaMensual_Change()

'    Call FormatoCajaTexto(txtTasaMensual, Decimales_Tasa)
'
'    If strCodMoneda = strCodMonedaGarantia Then
'        txtSubTotal(1).Text = CStr(((CDbl(txtTasaMensual.Text) / 100 + 1) ^ (CLng(txtDiasPlazo.Text) / 30)) * CDbl(txtPrecio(0).Text) * CLng(txtCantidad.Text) * CLng(lblValorNominal.Caption))
'        txtSubTotal(1).Text = CStr(((CDbl(txtTasaMensual.Text) / 100 + 1) ^ (CLng(txtDiasPlazo.Text) / 360)) * CDbl(txtPrecio(0).Text) * CLng(txtCantidad.Text) * CLng(lblValorNominal.Caption))
'    ElseIf strCodMoneda = Codigo_Moneda_Local And strCodMonedaGarantia <> Codigo_Moneda_Local Then
'            txtSubTotal(1).Text = CStr(((CDbl(txtTasaMensual.Text) / 100 + 1) ^ (CLng(txtDiasPlazo.Text) / 30)) * (CDbl(txtPrecio(0).Text) * CDbl(txtTipoCambio.Text)) * CLng(txtCantidad.Text) * CLng(lblValorNominal.Caption))
'        ElseIf strCodMoneda <> Codigo_Moneda_Local And strCodMonedaGarantia = Codigo_Moneda_Local Then
'                txtSubTotal(1).Text = CStr(((CDbl(txtTasaMensual.Text) / 100 + 1) ^ (CLng(txtDiasPlazo.Text) / 30)) * (CDbl(txtPrecio(0).Text) / CDbl(txtTipoCambio.Text)) * CLng(txtCantidad.Text) * CLng(lblValorNominal.Caption))
'            End If
    
End Sub

Private Sub txtTasaMensual_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtTasaMensual, Decimales_Tasa)
    
End Sub


Private Sub txtTipoCambio_Change()

    Call FormatoCajaTexto(txtTipoCambio, Decimales_TipoCambio)
    
End Sub


Private Sub txtTipoCambio_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtTipoCambio, Decimales_TipoCambio)
    
End Sub


Private Sub txtTirBrutaBase365_Change()

    Call FormatoCajaTexto(txtTirBrutaBase365, Decimales_Tasa)
    
End Sub

Private Sub txtTirBrutaBase365_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtTirBrutaBase365, Decimales_Tasa)
    
End Sub


Private Sub txtTirBrutaBaseBono_Change()

    Dim curCantidad As Currency, dblPreUni  As Double
    Dim curSubTotal As Currency
    
    If Trim(txtTirBrutaBaseBono.Text) = Valor_Caracter Then Exit Sub
    
    If txtTirBrutaBaseBono.Value > 0 And cboTitulo.ListIndex > 0 And txtMontoNominal.Value > 0 And txtCantidad.Value > 0 Then
        
        txtInteresCorrido(0).Text = "0"
        
        txtInteresCorrido(0).Text = CStr(CalculoInteresCorrido(strCodGarantia, txtMontoNominal.Value, CVDate(lblInicioFechaCupon.Caption), dtpFechaLiquidacion.Value, strCodIndiceFinal, strCodTipoAjuste, strCodTipoTasa, strCodPeriodoPago, strCodIndiceInicial, strCodBaseAnual, intBaseAnual))
        
        Call CalcularPrecioTIRVAN
                
        Call CalculoTotal(0)

    End If
    
End Sub


Private Sub txtTirBrutaBaseBono_KeyPress(KeyAscii As Integer)

    'Call ValidaCajaTexto(KeyAscii, "M", txtTirBrutaBaseBono, Decimales_Tasa)
    
End Sub


Private Sub txtTirNetaBase365_Change()

    Call FormatoCajaTexto(txtTirNetaBase365, Decimales_Tasa)
    
End Sub

Private Sub txtTirNetaBase365_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtTirNetaBase365, Decimales_Tasa)
    
End Sub

Private Sub txtTirNetaBaseBono_Change()

    'Call FormatoCajaTexto(txtTirNetaBaseBono, Decimales_Tasa)
    
End Sub

Private Sub txtTirNetaBaseBono_KeyPress(KeyAscii As Integer)

    'Call ValidaCajaTexto(KeyAscii, "M", txtTirNetaBaseBono, Decimales_Tasa)
    
End Sub

Private Sub txtVacCorrido_Change(Index As Integer)

    Call FormatoCajaTexto(txtVacCorrido(Index), Decimales_Monto)
    
    lblInteresesResumen(Index).Caption = CStr(CCur(txtInteresCorrido(Index).Text) + CCur(txtVacCorrido(Index).Text))
    
End Sub

Private Sub txtVacCorrido_KeyPress(Index As Integer, KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtVacCorrido(Index), Decimales_Monto)
    
    If KeyAscii = vbKeyReturn Then
        Call CalculoTotal(Index)
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
