VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmOrdenRentaVariable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ordenes - Al Contado Renta Variable"
   ClientHeight    =   9045
   ClientLeft      =   1365
   ClientTop       =   1830
   ClientWidth     =   14430
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
   Icon            =   "frmOrdenRentaVariable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9045
   ScaleWidth      =   14430
   Begin TAMControls2.ucBotonEdicion2 cmdOpcion 
      Height          =   735
      Left            =   780
      TabIndex        =   227
      Top             =   8280
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
      Left            =   12450
      TabIndex        =   225
      Top             =   8280
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
      Caption0        =   "&Salir"
      Tag0            =   "9"
      ToolTipText0    =   "Salir"
      UserControlWidth=   1200
   End
   Begin TabDlg.SSTab tabRentaVariable 
      Height          =   8265
      Left            =   0
      TabIndex        =   39
      Top             =   0
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   14579
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Lista"
      TabPicture(0)   =   "frmOrdenRentaVariable.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "tdgConsulta"
      Tab(0).Control(1)=   "fraCriterio"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Datos Orden Inversión"
      TabPicture(1)   =   "frmOrdenRentaVariable.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblDescrip(30)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraDatosBasicos"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "fraDatosTitulo"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "fraResumen"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtObservacion"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdAccion"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Negociación"
      TabPicture(2)   =   "frmOrdenRentaVariable.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraComisionMontoFL1"
      Tab(2).Control(1)=   "fraPosicion"
      Tab(2).Control(2)=   "fraDatosNegociacion"
      Tab(2).Control(3)=   "fraComisionMontoFL2"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Análisis de Riesgo"
      TabPicture(3)   =   "frmOrdenRentaVariable.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraRiesgoMercado"
      Tab(3).ControlCount=   1
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   10950
         TabIndex        =   226
         Top             =   7350
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
      Begin VB.Frame fraRiesgoMercado 
         Caption         =   "Análisis de VaR"
         Height          =   4725
         Left            =   -74430
         TabIndex        =   198
         Top             =   810
         Width           =   8355
         Begin VB.CommandButton Command1 
            Caption         =   "Calcular VaR"
            Height          =   405
            Left            =   540
            TabIndex        =   208
            Top             =   3990
            Width           =   1365
         End
         Begin VB.TextBox txtNivelConfianza 
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
            Left            =   2820
            MaxLength       =   45
            TabIndex        =   206
            Top             =   1350
            Width           =   705
         End
         Begin VB.TextBox txtUnidadesPeriodo 
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
            Left            =   6780
            MaxLength       =   45
            TabIndex        =   204
            Top             =   840
            Width           =   885
         End
         Begin VB.ComboBox cboTipoVaR 
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
            Left            =   2820
            Style           =   2  'Dropdown List
            TabIndex        =   203
            ToolTipText     =   "Fondo"
            Top             =   360
            Width           =   1725
         End
         Begin VB.ComboBox cboPeriodo 
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
            Left            =   2820
            Style           =   2  'Dropdown List
            TabIndex        =   200
            ToolTipText     =   "Fondo"
            Top             =   840
            Width           =   1725
         End
         Begin MSComCtl2.DTPicker dtpFechaAnalisisVaR 
            Height          =   315
            Left            =   2820
            TabIndex        =   215
            Top             =   1830
            Width           =   1395
            _ExtentX        =   2461
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
            Format          =   293404673
            CurrentDate     =   38776
         End
         Begin VB.Line Line4 
            X1              =   270
            X2              =   6780
            Y1              =   2280
            Y2              =   2280
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
            Index           =   8
            Left            =   5250
            TabIndex        =   219
            Top             =   2490
            Width           =   1215
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Valor de Operación"
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
            Index           =   84
            Left            =   240
            TabIndex        =   218
            Top             =   2490
            Width           =   1365
         End
         Begin VB.Label lblValorOperacion 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Left            =   2820
            TabIndex        =   217
            Tag             =   "0.00"
            ToolTipText     =   "Porcentaje de Comisión IGV"
            Top             =   2460
            Width           =   2295
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Análisis"
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
            Left            =   240
            TabIndex        =   216
            Top             =   1860
            Width           =   1245
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
            Index           =   7
            Left            =   5250
            TabIndex        =   214
            Top             =   3420
            Width           =   1215
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Left            =   2820
            TabIndex        =   213
            Tag             =   "0.00"
            ToolTipText     =   "Porcentaje de Comisión IGV"
            Top             =   3390
            Width           =   2295
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Máxima Pérdida Esperada"
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
            Index           =   82
            Left            =   240
            TabIndex        =   212
            Top             =   3420
            Width           =   1845
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
            Index           =   6
            Left            =   5250
            TabIndex        =   211
            Top             =   2970
            Width           =   1215
         End
         Begin VB.Label lblValorMercadoCartera 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Left            =   2820
            TabIndex        =   210
            Tag             =   "0.00"
            ToolTipText     =   "Porcentaje de Comisión IGV"
            Top             =   2940
            Width           =   2295
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Valor de Mercado de Cartera"
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
            Left            =   240
            TabIndex        =   209
            Top             =   2970
            Width           =   2040
         End
         Begin VB.Label Label1 
            Caption         =   "%"
            Height          =   255
            Left            =   3600
            TabIndex        =   207
            Top             =   1380
            Width           =   285
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Nivel de Confianza"
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
            Left            =   240
            TabIndex        =   205
            Top             =   1380
            Width           =   1335
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de VaR"
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
            Index           =   79
            Left            =   240
            TabIndex        =   202
            Top             =   420
            Width           =   900
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Unidades de Periodo"
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
            Index           =   78
            Left            =   4950
            TabIndex        =   201
            Top             =   870
            Width           =   1485
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Periodo de Análisis"
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
            Index           =   77
            Left            =   240
            TabIndex        =   199
            Top             =   900
            Width           =   1335
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
         Height          =   720
         Left            =   2040
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   7395
         Width           =   7920
      End
      Begin VB.Frame fraComisionMontoFL2 
         Caption         =   "Comisiones y Montos - Contado (FL1)"
         Height          =   375
         Left            =   -66150
         TabIndex        =   121
         Top             =   3270
         Width           =   6735
         Begin VB.TextBox txtComisionIgv 
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
            TabIndex        =   223
            Top             =   3870
            Width           =   1995
         End
         Begin VB.TextBox txtComisionGastoBancario 
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
            TabIndex        =   163
            Top             =   3150
            Width           =   2025
         End
         Begin VB.TextBox txtComisionxAccion 
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
            TabIndex        =   162
            Top             =   3510
            Width           =   2025
         End
         Begin VB.TextBox txtComisionFondoG 
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
            Left            =   4305
            MaxLength       =   45
            TabIndex        =   155
            Top             =   2535
            Width           =   2025
         End
         Begin VB.CommandButton cmdCalculo 
            Caption         =   "#"
            Height          =   285
            Left            =   360
            TabIndex        =   38
            ToolTipText     =   "Calcular TIRs de la orden"
            Top             =   4860
            Width           =   375
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
            TabIndex        =   32
            Top             =   1065
            Width           =   1340
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
            TabIndex        =   37
            Top             =   2850
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
            TabIndex        =   36
            Top             =   2130
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
            TabIndex        =   35
            Top             =   1770
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
            TabIndex        =   34
            Top             =   1410
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
            Index           =   1
            Left            =   4290
            MaxLength       =   45
            TabIndex        =   33
            Top             =   1065
            Width           =   2025
         End
         Begin VB.CheckBox chkAplicar 
            Caption         =   "Aplicar"
            Height          =   255
            Index           =   1
            Left            =   390
            TabIndex        =   31
            ToolTipText     =   "Aplicar Costos Negociación"
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox txtPrecioUnitario 
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
            Index           =   1
            Left            =   2625
            MaxLength       =   45
            TabIndex        =   30
            Top             =   240
            Width           =   1340
         End
         Begin VB.Label lblPorcenComisionAccion 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   1
            Left            =   2625
            TabIndex        =   168
            Top             =   3540
            Width           =   1335
         End
         Begin VB.Label lblPorcenGastoBancario 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   1
            Left            =   2625
            TabIndex        =   167
            Top             =   3180
            Width           =   1335
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Gastos Bancarios"
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
            Index           =   71
            Left            =   420
            TabIndex        =   161
            Top             =   3180
            Width           =   1245
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisión Especial"
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
            Index           =   70
            Left            =   420
            TabIndex        =   160
            Top             =   3510
            Width           =   1275
         End
         Begin VB.Label lblPorcenFondoG 
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
            Left            =   2640
            TabIndex        =   154
            Tag             =   "0"
            ToolTipText     =   "Porcentaje de Comisión Fondo Liquidación"
            Top             =   2520
            Width           =   1335
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
            Index           =   67
            Left            =   360
            TabIndex        =   152
            Top             =   2520
            Width           =   1800
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   5
            X1              =   360
            X2              =   6300
            Y1              =   4740
            Y2              =   4740
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
            TabIndex        =   142
            Tag             =   "0.00"
            Top             =   4860
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
            TabIndex        =   141
            Tag             =   "0.00"
            Top             =   4860
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
            Index           =   65
            Left            =   1200
            TabIndex        =   140
            Top             =   4875
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
            Index           =   64
            Left            =   3720
            TabIndex        =   139
            Top             =   4875
            Width           =   660
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
            Index           =   47
            Left            =   390
            TabIndex        =   137
            Top             =   1080
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
            Index           =   46
            Left            =   390
            TabIndex        =   136
            Top             =   1440
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
            Index           =   45
            Left            =   390
            TabIndex        =   135
            Top             =   1815
            Width           =   1110
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisión Fondo Liquidación"
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
            Index           =   44
            Left            =   360
            TabIndex        =   134
            Top             =   2175
            Width           =   1980
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
            Index           =   43
            Left            =   390
            TabIndex        =   133
            Top             =   2910
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
            Index           =   42
            Left            =   390
            TabIndex        =   132
            Top             =   3930
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
            Index           =   41
            Left            =   2640
            TabIndex        =   131
            Top             =   620
            Width           =   645
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
            Index           =   40
            Left            =   2640
            TabIndex        =   130
            Top             =   4380
            Width           =   855
         End
         Begin VB.Label lblSubTotal 
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
            TabIndex        =   129
            Tag             =   "0.00"
            ToolTipText     =   "SubTotal de la Orden"
            Top             =   600
            Width           =   2025
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   3
            X1              =   2580
            X2              =   6300
            Y1              =   960
            Y2              =   960
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
            TabIndex        =   128
            Tag             =   "0.00"
            ToolTipText     =   "Porcentaje de Comisión IGV"
            Top             =   3870
            Width           =   1335
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   2
            X1              =   360
            X2              =   6300
            Y1              =   4245
            Y2              =   4245
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
            TabIndex        =   127
            Tag             =   "0.00"
            ToolTipText     =   "Monto Total de la Orden"
            Top             =   4365
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
            Index           =   1
            Left            =   2625
            TabIndex        =   126
            Tag             =   "0"
            ToolTipText     =   "Porcentaje de Comisión BVL"
            Top             =   1410
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
            TabIndex        =   125
            Tag             =   "0"
            ToolTipText     =   "Porcentaje de Comisión Cavali"
            Top             =   1770
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
            TabIndex        =   124
            Tag             =   "0"
            ToolTipText     =   "Porcentaje de Comisión Fondo Liquidación"
            Top             =   2130
            Width           =   1335
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
            TabIndex        =   123
            Tag             =   "0"
            ToolTipText     =   "Porcentaje de Comisión Conasev"
            Top             =   2850
            Width           =   1335
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Precio Unitario"
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
            Left            =   390
            TabIndex        =   122
            Top             =   260
            Width           =   1035
         End
      End
      Begin VB.Frame fraResumen 
         Caption         =   "Resumen Negociación"
         Height          =   2805
         Left            =   240
         TabIndex        =   94
         Top             =   4470
         Width           =   13935
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
            Index           =   48
            Left            =   10200
            TabIndex        =   148
            Top             =   360
            Width           =   630
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
            TabIndex        =   147
            Tag             =   "0.00"
            Top             =   360
            Width           =   2025
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
            Index           =   38
            Left            =   5280
            TabIndex        =   120
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
            Index           =   37
            Left            =   480
            TabIndex        =   119
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
            TabIndex        =   118
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
            TabIndex        =   117
            Top             =   1900
            Width           =   795
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
            Index           =   60
            Left            =   480
            TabIndex        =   116
            Top             =   2260
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
            Index           =   59
            Left            =   5280
            TabIndex        =   115
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
            Index           =   58
            Left            =   5280
            TabIndex        =   114
            Top             =   1900
            Width           =   795
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
            Index           =   57
            Left            =   5280
            TabIndex        =   113
            Top             =   2260
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
            TabIndex        =   112
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
            TabIndex        =   111
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
            TabIndex        =   110
            Tag             =   "0.00"
            Top             =   1880
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
            TabIndex        =   109
            Tag             =   "0.00"
            Top             =   2240
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
            TabIndex        =   108
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
            TabIndex        =   107
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
            TabIndex        =   106
            Tag             =   "0.00"
            Top             =   1880
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
            TabIndex        =   105
            Tag             =   "0.00"
            Top             =   2240
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
            Index           =   55
            Left            =   480
            TabIndex        =   104
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
            Index           =   54
            Left            =   5280
            TabIndex        =   103
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
            Index           =   52
            Left            =   10200
            TabIndex        =   102
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
            Index           =   51
            Left            =   10200
            TabIndex        =   101
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
            TabIndex        =   100
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
            TabIndex        =   99
            Tag             =   "0.00"
            Top             =   1560
            Width           =   2025
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000015&
            X1              =   4800
            X2              =   4800
            Y1              =   360
            Y2              =   2500
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
            TabIndex        =   98
            Tag             =   "0.00"
            Top             =   360
            Width           =   2025
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad Nominal"
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
            Index           =   49
            Left            =   480
            TabIndex        =   97
            Top             =   380
            Width           =   1245
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
            TabIndex        =   96
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
            TabIndex        =   95
            Top             =   840
            Width           =   1845
         End
         Begin VB.Line Line3 
            BorderColor     =   &H80000015&
            X1              =   9720
            X2              =   9720
            Y1              =   360
            Y2              =   2500
         End
      End
      Begin VB.Frame fraDatosNegociacion 
         Caption         =   "Negociación"
         Height          =   1785
         Left            =   -74640
         TabIndex        =   91
         Top             =   720
         Width           =   8805
         Begin VB.CheckBox chkAplicarMon 
            Caption         =   "Aplicar Moneda de Pago para Comisiones"
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
            Height          =   255
            Index           =   0
            Left            =   4320
            TabIndex        =   180
            ToolTipText     =   "Aplicar Costos Negociación"
            Top             =   2160
            Visible         =   0   'False
            Width           =   3855
         End
         Begin VB.TextBox txtTipoCambioConversion 
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
            Left            =   5520
            TabIndex        =   178
            Top             =   2700
            Visible         =   0   'False
            Width           =   1605
         End
         Begin VB.TextBox txtPrecioUnitario 
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
            Index           =   0
            Left            =   2520
            TabIndex        =   169
            Top             =   810
            Width           =   1695
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
            Left            =   7080
            MaxLength       =   45
            TabIndex        =   21
            Top             =   810
            Width           =   1605
         End
         Begin VB.TextBox txtCantidad 
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
            Left            =   2520
            MaxLength       =   45
            TabIndex        =   20
            Top             =   450
            Width           =   1695
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
            Left            =   2160
            TabIndex        =   183
            Top             =   2100
            Width           =   435
         End
         Begin VB.Label lblMonPagoAMonOrigen 
            Alignment       =   2  'Center
            Caption         =   "lblMonPagoAMonOrigen"
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
            Height          =   225
            Left            =   4470
            TabIndex        =   182
            Top             =   2760
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.Label lblMonOrigenAMonPago 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "lblMonOrigenAMonPago"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   5955
            TabIndex        =   181
            Top             =   840
            Width           =   1080
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
            Left            =   2130
            TabIndex        =   179
            Top             =   1200
            Width           =   345
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Cambio Pactado"
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
            Index           =   75
            Left            =   2760
            TabIndex        =   177
            Top             =   2760
            Visible         =   0   'False
            Width           =   1530
         End
         Begin VB.Label lblSubTotal 
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
            Left            =   2520
            TabIndex        =   175
            Tag             =   "0.00"
            ToolTipText     =   "SubTotal de la Orden"
            Top             =   1170
            Width           =   1695
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "SubTotal Moneda Pago"
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
            Left            =   480
            TabIndex        =   172
            Top             =   2400
            Width           =   1725
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
            Index           =   13
            Left            =   210
            TabIndex        =   171
            Top             =   1200
            Width           =   645
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Precio Unitario"
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
            Left            =   210
            TabIndex        =   170
            Top             =   825
            Width           =   1035
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Liquidación"
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
            Left            =   4740
            TabIndex        =   145
            Top             =   450
            Width           =   1530
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
            Left            =   7080
            TabIndex        =   144
            Tag             =   "0.00"
            ToolTipText     =   "Fecha de Liquidación"
            Top             =   450
            Width           =   1605
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000015&
            X1              =   4500
            X2              =   4500
            Y1              =   420
            Y2              =   1470
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
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   17
            Left            =   4740
            TabIndex        =   93
            Top             =   870
            Width           =   885
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad Nominal"
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
            Left            =   210
            TabIndex        =   92
            Top             =   465
            Width           =   1245
         End
      End
      Begin VB.Frame fraPosicion 
         Caption         =   "Datos Posición"
         Height          =   2085
         Left            =   -65790
         TabIndex        =   80
         Top             =   720
         Width           =   4995
         Begin VB.Label lblMonedaPago 
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
            Left            =   2370
            TabIndex        =   174
            Tag             =   "0.00"
            ToolTipText     =   "Moneda del Título"
            Top             =   1740
            Width           =   2025
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
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   74
            Left            =   630
            TabIndex        =   173
            Top             =   1770
            Width           =   1005
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
            Index           =   12
            Left            =   630
            TabIndex        =   90
            Top             =   255
            Width           =   975
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Ultimo Precio"
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
            Index           =   5
            Left            =   630
            TabIndex        =   89
            Top             =   555
            Width           =   930
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
            Left            =   630
            TabIndex        =   88
            Top             =   855
            Width           =   885
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
            Index           =   16
            Left            =   630
            TabIndex        =   87
            Top             =   1155
            Width           =   1035
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
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   56
            Left            =   630
            TabIndex        =   86
            Top             =   1455
            Width           =   1170
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
            Height          =   285
            Left            =   2370
            TabIndex        =   85
            Tag             =   "0.00"
            Top             =   240
            Width           =   2025
         End
         Begin VB.Label lblUltimoPrecio 
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
            Left            =   2370
            TabIndex        =   84
            Tag             =   "0.00"
            Top             =   540
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
            Height          =   285
            Left            =   2370
            TabIndex        =   83
            Tag             =   "0.00"
            Top             =   840
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
            Height          =   285
            Left            =   2370
            TabIndex        =   82
            Tag             =   "0.00"
            Top             =   1140
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
            Height          =   285
            Left            =   2370
            TabIndex        =   81
            Tag             =   "0.00"
            ToolTipText     =   "Moneda del Título"
            Top             =   1440
            Width           =   2025
         End
      End
      Begin VB.Frame fraComisionMontoFL1 
         Caption         =   "ComisionAgente"
         Height          =   4665
         Left            =   -74640
         TabIndex        =   66
         Top             =   2700
         Width           =   6555
         Begin VB.TextBox txtComisionIgv 
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
            Left            =   4290
            MaxLength       =   45
            TabIndex        =   222
            Top             =   3660
            Width           =   1905
         End
         Begin VB.TextBox txtComisionxAccion 
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
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   4290
            TabIndex        =   159
            Top             =   3300
            Width           =   1905
         End
         Begin VB.TextBox txtComisionGastoBancario 
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
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   4290
            TabIndex        =   158
            Top             =   2940
            Width           =   1905
         End
         Begin VB.TextBox txtComisionFondoG 
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
            Left            =   4290
            MaxLength       =   45
            TabIndex        =   153
            Top             =   2280
            Width           =   1905
         End
         Begin MSAdodcLib.Adodc adoCostoFL1 
            Height          =   330
            Left            =   3540
            Top             =   5250
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
            Left            =   1470
            MaxLength       =   45
            TabIndex        =   29
            Top             =   5520
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CheckBox chkAplicar 
            Caption         =   "Aplicar"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   22
            ToolTipText     =   "Aplicar Costos Negociación"
            Top             =   420
            Width           =   2205
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
            Left            =   4290
            MaxLength       =   45
            TabIndex        =   24
            Top             =   840
            Width           =   1905
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
            Left            =   4290
            MaxLength       =   45
            TabIndex        =   25
            Top             =   1170
            Width           =   1905
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
            Left            =   4290
            MaxLength       =   45
            TabIndex        =   26
            Top             =   1530
            Width           =   1905
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
            Left            =   4290
            MaxLength       =   45
            TabIndex        =   27
            Top             =   1890
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
            Left            =   4290
            MaxLength       =   45
            TabIndex        =   28
            Top             =   2610
            Width           =   1905
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
            Index           =   0
            Left            =   2310
            MaxLength       =   45
            TabIndex        =   23
            Top             =   855
            Width           =   1905
         End
         Begin TrueOleDBGrid60.TDBGrid tdgCostoFL1 
            Height          =   615
            Left            =   3570
            OleObjectBlob   =   "frmOrdenRentaVariable.frx":04B2
            TabIndex        =   164
            Top             =   5370
            Visible         =   0   'False
            Width           =   2385
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   1
            X1              =   1890
            X2              =   6360
            Y1              =   4080
            Y2              =   4080
         End
         Begin VB.Label lblComisionxAccionOrigen 
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
            Left            =   180
            TabIndex        =   197
            Tag             =   "0.00"
            ToolTipText     =   "Monto de Comisión IGV"
            Top             =   7320
            Width           =   1365
         End
         Begin VB.Label lblComisionGastoBancarioOrigen 
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
            Left            =   180
            TabIndex        =   196
            Tag             =   "0.00"
            ToolTipText     =   "Monto de Comisión IGV"
            Top             =   6990
            Width           =   1365
         End
         Begin VB.Label lblComisionConasevOrigen 
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
            Left            =   180
            TabIndex        =   195
            Tag             =   "0.00"
            ToolTipText     =   "Monto de Comisión IGV"
            Top             =   6630
            Width           =   1365
         End
         Begin VB.Label lblComisionFondoGOrigen 
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
            Left            =   180
            TabIndex        =   194
            Tag             =   "0.00"
            ToolTipText     =   "Monto de Comisión IGV"
            Top             =   6270
            Width           =   1365
         End
         Begin VB.Label lblComisionFondoOrigen 
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
            Left            =   180
            TabIndex        =   193
            Tag             =   "0.00"
            ToolTipText     =   "Monto de Comisión IGV"
            Top             =   5880
            Width           =   1365
         End
         Begin VB.Label lblComisionCavaliOrigen 
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
            Left            =   180
            TabIndex        =   192
            Tag             =   "0.00"
            ToolTipText     =   "Monto de Comisión IGV"
            Top             =   5520
            Width           =   1365
         End
         Begin VB.Label lblComisionBolsaOrigen 
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
            Left            =   180
            TabIndex        =   191
            Tag             =   "0.00"
            ToolTipText     =   "Monto de Comisión IGV"
            Top             =   5190
            Width           =   1365
         End
         Begin VB.Label lblComisionAgenteOrigen 
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
            Left            =   1590
            TabIndex        =   190
            Tag             =   "0.00"
            ToolTipText     =   "Monto de Comisión IGV"
            Top             =   5190
            Width           =   1365
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
            Index           =   5
            Left            =   2550
            TabIndex        =   189
            Top             =   4890
            Width           =   285
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
            Index           =   4
            Left            =   3840
            TabIndex        =   188
            Top             =   4260
            Width           =   375
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
            Index           =   76
            Left            =   2040
            TabIndex        =   187
            Top             =   4950
            Width           =   1005
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
            Index           =   3
            Left            =   1740
            TabIndex        =   186
            Top             =   5460
            Width           =   1155
         End
         Begin VB.Label lblComisionIgvOrigen 
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
            Left            =   180
            TabIndex        =   185
            Tag             =   "0.00"
            ToolTipText     =   "Monto de Comisión IGV"
            Top             =   7650
            Width           =   1365
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
            Index           =   2
            Left            =   4710
            TabIndex        =   184
            Top             =   480
            Width           =   1425
         End
         Begin VB.Label lblDescrip 
            Height          =   60
            Index           =   72
            Left            =   150
            TabIndex        =   176
            Top             =   330
            Width           =   45
         End
         Begin VB.Label lblPorcenComisionAccion 
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
            Left            =   2310
            TabIndex        =   166
            Tag             =   "0"
            Top             =   3360
            Width           =   1905
         End
         Begin VB.Label lblPorcenGastoBancario 
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
            Left            =   2310
            TabIndex        =   165
            Tag             =   "0"
            Top             =   3000
            Width           =   1905
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisión Especial"
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
            Index           =   69
            Left            =   240
            TabIndex        =   157
            Top             =   3360
            Width           =   1275
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Gastos Bancarios"
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
            Index           =   68
            Left            =   240
            TabIndex        =   156
            Top             =   3030
            Width           =   1245
         End
         Begin VB.Label lblPorcenFondoG 
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
            Left            =   2310
            TabIndex        =   151
            Tag             =   "0"
            ToolTipText     =   "Porcentaje de Comisión Fondo Garantía"
            Top             =   2295
            Width           =   1905
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
            Index           =   66
            Left            =   240
            TabIndex        =   150
            Top             =   2310
            Width           =   1800
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
            Index           =   62
            Left            =   1470
            TabIndex        =   143
            Top             =   5250
            Visible         =   0   'False
            Width           =   1260
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
            Left            =   2310
            TabIndex        =   79
            Tag             =   "0"
            ToolTipText     =   "Porcentaje de Comisión Conasev"
            Top             =   2640
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
            Left            =   2310
            TabIndex        =   78
            Tag             =   "0"
            ToolTipText     =   "Porcentaje de Comisión Fondo Liquidación"
            Top             =   1920
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
            Left            =   2310
            TabIndex        =   77
            Tag             =   "0"
            ToolTipText     =   "Porcentaje de Comisión Cavali"
            Top             =   1560
            Width           =   1905
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
            Left            =   2310
            TabIndex        =   76
            Tag             =   "0"
            ToolTipText     =   "Porcentaje de Comisión BVL"
            Top             =   1200
            Width           =   1905
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
            Left            =   4290
            TabIndex        =   75
            Tag             =   "0.00"
            ToolTipText     =   "Monto Total de la Orden"
            Top             =   4200
            Width           =   1905
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
            Left            =   2310
            TabIndex        =   74
            Tag             =   "0.00"
            ToolTipText     =   "Porcentaje de Comisión IGV"
            Top             =   3690
            Width           =   1905
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   0
            X1              =   2190
            X2              =   6390
            Y1              =   750
            Y2              =   750
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
            Index           =   11
            Left            =   2520
            TabIndex        =   73
            Top             =   4260
            Width           =   1215
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
            Index           =   10
            Left            =   240
            TabIndex        =   72
            Top             =   3720
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
            Index           =   28
            Left            =   240
            TabIndex        =   71
            Top             =   2700
            Width           =   1305
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisión Fondo Liquidación"
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
            Left            =   240
            TabIndex        =   70
            Top             =   1965
            Width           =   1980
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
            Left            =   240
            TabIndex        =   69
            Top             =   1605
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
            Index           =   26
            Left            =   240
            TabIndex        =   68
            Top             =   1230
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
            Index           =   14
            Left            =   240
            TabIndex        =   67
            Top             =   870
            Width           =   990
         End
      End
      Begin VB.Frame fraCriterio 
         Caption         =   "Criterios de Búsqueda"
         Height          =   2175
         Left            =   -74760
         TabIndex        =   54
         Top             =   780
         Width           =   13935
         Begin VB.CommandButton cmdExportarExcel 
            Caption         =   "Excel"
            Height          =   735
            Left            =   10680
            Picture         =   "frmOrdenRentaVariable.frx":38F6
            Style           =   1  'Graphical
            TabIndex        =   224
            Top             =   1200
            Width           =   1200
         End
         Begin VB.CommandButton cmdEnviar 
            Caption         =   "En&viar"
            Height          =   735
            Left            =   12180
            Picture         =   "frmOrdenRentaVariable.frx":3EFE
            Style           =   1  'Graphical
            TabIndex        =   149
            ToolTipText     =   "Enviar a BackOffice"
            Top             =   1200
            Width           =   1200
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
            TabIndex        =   0
            Top             =   360
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
            TabIndex        =   1
            Top             =   780
            Width           =   5145
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
            TabIndex        =   2
            Top             =   1200
            Width           =   5145
         End
         Begin MSComCtl2.DTPicker dtpFechaOrdenDesde 
            Height          =   285
            Left            =   9600
            TabIndex        =   3
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
            TabIndex        =   4
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
            TabIndex        =   5
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
            TabIndex        =   6
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
            Index           =   35
            Left            =   11280
            TabIndex        =   63
            Top             =   800
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
            Index           =   34
            Left            =   8880
            TabIndex        =   62
            Top             =   800
            Width           =   465
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
            Index           =   33
            Left            =   7200
            TabIndex        =   61
            Top             =   800
            Width           =   1305
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
            Index           =   32
            Left            =   7200
            TabIndex        =   60
            Top             =   380
            Width           =   930
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
            TabIndex        =   59
            Top             =   380
            Width           =   450
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
            TabIndex        =   58
            Top             =   380
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
            Index           =   9
            Left            =   11280
            TabIndex        =   57
            Top             =   380
            Width           =   420
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
            Index           =   8
            Left            =   480
            TabIndex        =   56
            Top             =   800
            Width           =   825
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
            TabIndex        =   55
            Top             =   1220
            Width           =   495
         End
      End
      Begin VB.Frame fraDatosTitulo 
         Caption         =   "Datos de la Orden"
         Height          =   1155
         Left            =   240
         TabIndex        =   48
         Top             =   3225
         Width           =   13935
         Begin VB.TextBox txtNroPoliza 
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
            Left            =   9300
            MaxLength       =   15
            TabIndex        =   221
            Top             =   660
            Width           =   1890
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
            Left            =   1740
            MaxLength       =   45
            TabIndex        =   18
            Top             =   690
            Width           =   4920
         End
         Begin MSComCtl2.DTPicker dtpFechaOrden 
            Height          =   285
            Left            =   1740
            TabIndex        =   50
            Top             =   315
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
            Left            =   9300
            TabIndex        =   17
            Top             =   285
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
            Caption         =   "Nro. Poliza"
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
            Index           =   85
            Left            =   6960
            TabIndex        =   220
            Top             =   690
            Width           =   765
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
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   6
            Left            =   480
            TabIndex        =   53
            Top             =   345
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
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   4
            Left            =   6960
            TabIndex        =   52
            Top             =   315
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
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   7
            Left            =   480
            TabIndex        =   51
            Top             =   750
            Width           =   840
         End
      End
      Begin VB.Frame fraDatosBasicos 
         Caption         =   "Datos Básicos"
         Height          =   2380
         Left            =   240
         TabIndex        =   40
         Top             =   750
         Width           =   13935
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
            Left            =   9330
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   1455
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
            TabIndex        =   16
            Top             =   1815
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
            TabIndex        =   10
            Top             =   1455
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
            TabIndex        =   8
            ToolTipText     =   "Agente"
            Top             =   735
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
            Left            =   1740
            Style           =   2  'Dropdown List
            TabIndex        =   7
            ToolTipText     =   "Fondo"
            Top             =   360
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
            Left            =   9330
            Style           =   2  'Dropdown List
            TabIndex        =   14
            ToolTipText     =   "Mecanismo de Negociación"
            Top             =   1086
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
            TabIndex        =   9
            ToolTipText     =   "Instrumento de Inversión"
            Top             =   1080
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
            TabIndex        =   11
            ToolTipText     =   "Orden de..."
            Top             =   1815
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
            Left            =   9330
            Style           =   2  'Dropdown List
            TabIndex        =   13
            ToolTipText     =   "Tipo de Operación"
            Top             =   723
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
            Left            =   9330
            Style           =   2  'Dropdown List
            TabIndex        =   12
            ToolTipText     =   "Títulos"
            Top             =   360
            Width           =   4185
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
            Height          =   195
            Index           =   3
            Left            =   6960
            TabIndex        =   146
            Top             =   1469
            Width           =   1530
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
            Index           =   50
            Left            =   6960
            TabIndex        =   138
            Top             =   1835
            Width           =   1575
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
            Left            =   480
            TabIndex        =   49
            Top             =   1455
            Width           =   390
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
            Left            =   480
            TabIndex        =   47
            Top             =   375
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
            Left            =   480
            TabIndex        =   46
            Top             =   750
            Width           =   510
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
            Height          =   195
            Index           =   2
            Left            =   6960
            TabIndex        =   45
            Top             =   380
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
            Height          =   195
            Index           =   21
            Left            =   6960
            TabIndex        =   44
            Top             =   1106
            Width           =   1755
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
            TabIndex        =   43
            Top             =   1125
            Width           =   825
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
            Index           =   24
            Left            =   480
            TabIndex        =   42
            Top             =   1860
            Width           =   660
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
            Height          =   195
            Index           =   25
            Left            =   6960
            TabIndex        =   41
            Top             =   743
            Width           =   1590
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmOrdenRentaVariable.frx":4459
         Height          =   4485
         Left            =   -74760
         OleObjectBlob   =   "frmOrdenRentaVariable.frx":4473
         TabIndex        =   64
         Top             =   3180
         Width           =   13905
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
         Index           =   30
         Left            =   600
         TabIndex        =   65
         Top             =   7440
         Width           =   1065
      End
   End
End
Attribute VB_Name = "frmOrdenRentaVariable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Ordenes de Acciones"
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
Dim strCodAgente            As String, strCodigosFile               As String
Dim strCodConcepto          As String, strNemonico                  As String
Dim strEstado               As String, strSQL                       As String

Dim strCodFile              As String, strCodAnalitica              As String
Dim strCodGrupo             As String, strCodCiiu                   As String
Dim strEstadoOrden          As String, strCodCategoria              As String
Dim strCodRiesgo            As String, strCodSubRiesgo              As String
Dim strCalcVcto             As String, strCodSector                 As String
Dim strCodTipoCostoBolsa    As String, strCodTipoCostoConasev       As String
Dim strCodTipoCostoFondo    As String, strCodTipoCavali             As String
Dim strCodTipoCostoAgente   As String

Dim strIndCuponCero         As String, strCodGarantia               As String
Dim strCodTipoCostoFondoG   As String, strCodMonedaPago             As String
Dim dblTipoCambio           As Double, dblComisionFondoG            As Double
Dim dblComisionBolsa        As Double, dblComisionConasev           As Double
Dim dblComisionFondo        As Double, dblComisionCavali            As Double
Dim dblComisionAgente       As Double, dblComisionAgenteMin         As Double
Public oExportacion As clsExportacion
Dim strTipoDocumento        As String, strNumDocumento              As String

Dim AplicarTipoCambioPactado As Boolean '/**/
Dim MonedaPago              As String '/**/
Dim adoExportacion          As ADODB.Recordset
Public indOk                As Boolean
Dim adoConsulta             As ADODB.Recordset
Dim indSortAsc              As Boolean, indSortDesc                 As Boolean

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
        
    If Trim(txtNroPoliza.Text) = Valor_Caracter Then
        MsgBox "Debe indicar el número de Póliza.", vbCritical, Me.Caption
        If txtDescripOrden.Enabled Then txtDescripOrden.SetFocus
        Exit Function
    End If
        
    If Trim(txtDescripOrden.Text) = Valor_Caracter Then
        MsgBox "Debe indicar la Descripción de la ORDEN.", vbCritical, Me.Caption
        If txtDescripOrden.Enabled Then txtDescripOrden.SetFocus
        Exit Function
    End If
                        
    If CCur(txtCantidad.Text) = 0 Then
        MsgBox "Debe indicar la cantidad de títulos.", vbCritical, Me.Caption
        If txtCantidad.Enabled Then txtCantidad.SetFocus
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
    
    If CDbl(txtTipoCambio.Text) = 0 Then
        MsgBox "Debe indicar el Tipo de Cambio.", vbCritical, Me.Caption
        If txtTipoCambio.Enabled Then txtTipoCambio.SetFocus
        Exit Function
    End If
    
    'CDbl(txtSubTotalMonedaPago.Text) / CDbl(lblSubTotal(Index).Caption)
    
    If CDbl(lblSubTotal(0).Caption) = 0 Then
        MsgBox "Debe indicar el monto en la moneda de posición.", vbCritical, Me.Caption
        If txtTipoCambio.Enabled Then txtTipoCambio.SetFocus
        Exit Function
    End If
    
    
    '*** Validación de STOCK ***
    If strCodTipoOrden = Codigo_Orden_Venta Then
        If CCur(txtCantidad.Text) > CCur(lblStockNominal.Caption) Then
            MsgBox "Stock insuficiente para Registrar la Orden de Venta.", vbCritical, Me.Caption
            If txtCantidad.Enabled Then txtCantidad.SetFocus
            Exit Function
        End If
    End If
            
    '*** Si todo paso OK ***
    TodoOK = True
  
End Function

Public Sub Cancelar()

    cmdOpcion.Visible = True
    With tabRentaVariable
        .TabEnabled(0) = True
        .Tab = 0
    End With
    Call Buscar
    
End Sub

Public Sub Adicionar()

'    If Not EsDiaUtil(gdatFechaActual) Then
'        MsgBox "No se puede negociar en un día no útil !", vbCritical, Me.Caption
'        Exit Sub
'    End If
    
    If cboTipoInstrumento.ListCount > 1 Then
        frmMainMdi.stbMdi.Panels(3).Text = "Adicionar orden..."
                    
        strEstado = Reg_Adicion
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabRentaVariable
            .TabEnabled(0) = False
            .TabEnabled(2) = False
            .TabEnabled(3) = False
            .Tab = 1
        End With
    Else
        MsgBox "Acceso a Negociación Denegada", vbCritical, Me.Caption
    End If
    
End Sub

Private Sub LlenarFormulario(strModo As String)

    Dim adoRecord   As ADODB.Recordset
    Dim strSQL      As String
    Dim intRegistro As Integer
    
    Select Case strModo
        Case Reg_Adicion
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
            
            cboAgente.ListIndex = -1
            If cboAgente.ListCount > 0 Then cboAgente.ListIndex = 0
                                                
            intRegistro = ObtenerItemLista(arrOrigen(), Codigo_Negociacion_Local)
            If intRegistro >= 0 Then cboOrigen.ListIndex = intRegistro
            
            dtpFechaOrden.Value = gdatFechaActual
            dtpFechaLiquidacion.Value = DateAdd("d", gintDiasInversionRV, dtpFechaOrden.Value)
            dtpFechaLiquidacion_Change
            lblFechaLiquidacion.Caption = CStr(dtpFechaLiquidacion.Value)
                                    
            txtDescripOrden.Text = Valor_Caracter
            txtObservacion.Text = Valor_Caracter
            txtNroPoliza.Text = Valor_Caracter
            
            txtPrecioUnitario(0).Text = "0"
            txtPrecioUnitario(1).Text = "0"
            lblValorNominal.Caption = "0"
            txtCantidad.Text = "0"
            
            
            '/**/
            txtTipoCambioConversion.Text = "0"
            '/**/


            lblAnalitica.Caption = "??? - ????????"
            lblStockNominal.Caption = "0"
            lblUltimoPrecio.Caption = "0"
            lblCantidadResumen.Caption = "0"
                        
            chkAplicar(0).Value = vbUnchecked
            chkAplicar(1).Value = vbUnchecked
            lblSubTotal(0).Caption = "0"
            lblSubTotal(1).Caption = "0"
            'txtSubTotalMonedaPago.Text = "0" '/**/
            
            Call IniciarComisiones
                        
            lblMontoTotal(0).Caption = "0"
            lblMontoTotal(1).Caption = "0"
            txtTasaMensual.Text = "0"
            lblTirBruta.Caption = "0"
            lblTirNeta.Caption = "0"
            lblTirBrutaResumen.Caption = "0"
            lblTirNetaResumen.Caption = "0"
            
            cboFondoOrden.SetFocus
    End Select
    
End Sub
Public Sub Grabar()

    Dim adoRegistro         As ADODB.Recordset
    Dim strFechaOrden       As String, strFechaLiquidacion      As String
    Dim strFechaEmision     As String, strFechaVencimiento      As String
    Dim strMensaje          As String, strIndTitulo             As String
    Dim strCodReportado     As String, strFechaConfirmacion     As String
    Dim intRegistro         As Integer, intAccion               As Integer
    Dim lngNumError         As Long, strDescripOrden            As String
    
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
                "Poliza Nro." & Space(4) & ">" & Space(2) & Trim(txtNroPoliza.Text) & Chr(vbKeyReturn) & _
                "Fecha de Operación" & Space(4) & ">" & Space(2) & CStr(dtpFechaOrden.Value) & Chr(vbKeyReturn) & _
                "Fecha de Liquidación" & Space(3) & ">" & Space(2) & CStr(dtpFechaLiquidacion.Value) & Chr(vbKeyReturn) & Chr(vbKeyReturn) & _
                "Nominal" & Space(24) & ">" & Space(2) & lblValorNominal.Caption & Chr(vbKeyReturn) & _
                "Cantidad" & Space(22) & ">" & Space(2) & txtCantidad.Text & Chr(vbKeyReturn) & _
                "Precio Unitario (%)" & Space(6) & ">" & Space(2) & txtPrecioUnitario(0).Text & Chr(vbKeyReturn) & Chr(vbKeyReturn) & _
                "Monto Total" & Space(17) & ">" & Space(2) & Trim(lblDescripMonedaResumen(0).Caption) & Space(1) & lblMontoTotal(0).Caption & Chr(vbKeyReturn) & Chr(vbKeyReturn) & _
                Chr(vbKeyReturn) & Chr(vbKeyReturn) & _
                "¿ Seguro de continuar ?"

            If MsgBox(strMensaje, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
               Me.Refresh: Exit Sub
            End If

        
            Me.MousePointer = vbHourglass
            
            strNumDocumento = Trim(txtNroPoliza.Text)
            strFechaOrden = Convertyyyymmdd(dtpFechaOrden.Value)
            strFechaEmision = Convertyyyymmdd(dtpFechaOrden.Value)
            strFechaLiquidacion = Convertyyyymmdd(dtpFechaLiquidacion.Value)
            strFechaVencimiento = strFechaLiquidacion
            strFechaConfirmacion = Convertyyyymmdd(Valor_Fecha)
                        
            strDescripOrden = UCase(Trim(cboTipoOrden.Text) & " " & Trim(cboTipoInstrumentoOrden.Text) & " " & Trim(Left(cboTitulo.Text, 15)) & " CANT: " & (txtCantidad.Text) & " PRECIO: " & Trim(lblDescripMoneda(0).Caption) & " " & (txtPrecioUnitario(0).Text))
                        
                                    
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
                    strCodMonedaPago = lblDescripMoneda(0).Tag
                    strFechaVencimiento = Convertyyyymmdd(Valor_Fecha)
                    strCodReportado = Valor_Caracter
                End If
                
'                .CommandText = "BEGIN TRAN ProcOrden"
'                adoConn.Execute .CommandText
'CDec (lblSubTotal(0).Caption)
'CDec(txtSubTotalMonedaPago.Text)

                 Dim dblTipoCambioArbitraje As Double
 
                 'dblTipoCambioArbitraje = CDbl(txtSubTotalMonedaPago.Text) / CDbl(lblSubTotal(0).Caption)
                                
                .CommandText = "{ call up_IVAdicInversionOrden('" & strCodFondoOrden & "','" & _
                    gstrCodAdministradora & "','','" & strFechaOrden & "','" & _
                    strCodTitulo & "','" & strNemonico & "','" & gstrPeriodoActual & "','" & gstrMesActual & "','" & _
                    "','" & strEstadoOrden & "','" & strCodAnalitica & "','" & strCodFile & "','" & _
                    strCodAnalitica & "','" & strCodClaseInstrumento & "','" & strCodSubClaseInstrumento & "','" & strCodTipoOrden & "','" & _
                    strCodOperacion & "','" & strCodNegociacion & "','" & strCodOrigen & "','" & strDescripOrden & "','" & strCodEmisor & "','" & _
                    strCodAgente & "','" & strCodGarantia & "','" & strFechaConfirmacion & "','" & strFechaVencimiento & "','" & strFechaLiquidacion & "','" & _
                    strFechaEmision & "','" & strCodMoneda & "'," & CDec(txtCantidad.Text) & "," & CDec(txtTipoCambio.Text) & "," & _
                    CDec(lblValorNominal.Caption) & ",100," & CDec(lblValorNominal.Caption) & "," & CDec(txtPrecioUnitario(0).Text) & "," & CDec(txtPrecioUnitario(0).Text) & "," & CDec(lblSubTotal(0).Caption) & "," & _
                    "0," & CDec(txtComisionAgente(0).Text) & "," & CDec(txtComisionCavali(0).Text) & "," & _
                    CDec(txtComisionConasev(0).Text) & "," & CDec(txtComisionBolsa(0).Text) & "," & CDec(txtComisionFondo(0).Text) & "," & CDec(txtComisionFondoG(0).Text) & "," & CDec(txtComisionGastoBancario(0).Text) & "," & CDec(txtComisionxAccion(0).Text) & "," & _
                    CDec(txtComisionIgv(0).Text) & "," & CDec(lblMontoTotal(0).Caption) & "," & CDec(txtPrecioUnitario(1).Text) & "," & CDec(txtPrecioUnitario(1).Text) & "," & _
                    CDec(lblSubTotal(1).Caption) & ",0," & CDec(txtComisionAgente(1).Text) & "," & _
                    CDec(txtComisionCavali(1).Text) & "," & CDec(txtComisionConasev(1).Text) & "," & CDec(txtComisionBolsa(1).Text) & "," & _
                    CDec(txtComisionFondo(1).Text) & "," & CDec(txtComisionFondoG(1).Text) & "," & CDec(txtComisionGastoBancario(1).Text) & "," & CDec(txtComisionxAccion(1).Text) & "," & CDec(txtComisionIgv(1).Text) & "," & CDec(lblMontoTotal(1).Caption) & "," & _
                    "0,0,'','','','" & strTipoDocumento & "','" & strNumDocumento & "','" & strCodAgente & "','','','','','',0,'','','" & strIndTitulo & "','" & _
                    strCodTipoTasa & "','" & strCodBaseAnual & "'," & CDec(txtTasaMensual.Text) & "," & CDec(lblTirBruta.Caption) & "," & CDec(lblTirBruta.Caption) & "," & CDec(lblTirNeta.Caption) & ",'" & _
                    strCodRiesgo & "','" & strCodSubRiesgo & "','" & Trim(txtObservacion.Text) & "') }"
                adoConn.Execute .CommandText
                
'                .CommandText = "INSERT INTO InversionOrdenCosto VALUES ('" & _
'                    strCodFondoOrden & "','" & gstrCodAdministradora & "','','" & strFechaOrden & "','" & _
'                    strCodTitulo & "','" & Codigo_Valor_RentaVariable & "','" & strCodConcepto & "','" & _
'                    Codigo_Operacion_Contado & "','" & Codigo_Costo_Bolsa & "'," & _
'                    CDec(lblPorcenBolsa(0).Caption) & "," & CDec(txtComisionBolsa(0).Text) & ")"
'                adoConn.Execute .CommandText
                
'                .CommandText = "COMMIT TRAN ProcOrden"
'                adoConn.Execute .CommandText
                                                                                                      
            End With
                                                                                    
            Me.MousePointer = vbDefault
        
            MsgBox Mensaje_Adicion_Exitosa, vbExclamation
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
            cmdOpcion.Visible = True
            With tabRentaVariable
                .TabEnabled(0) = True
                .Tab = 0
            End With
            Call Buscar
        End If
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
'    adoComm.CommandText = "ROLLBACK TRAN ProcOrden"
'    adoConn.Execute adoComm.CommandText
    
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
            
            tabRentaVariable.TabEnabled(0) = True
            tabRentaVariable.Tab = 0
            Call Buscar
            
            Exit Sub
        End If
    End If
    
End Sub
Private Function PosicionLimites() As Boolean

    PosicionLimites = False
'
'    If cboFondo.ListIndex < 0 Then
'        MsgBox "Debe seleccionar el Fondo.", vbCritical, Me.Caption
'        cboTitulo.ListIndex = -1
'        If cboFondo.Enabled Then cboFondo.SetFocus
'        Exit Function
'    End If
'
'    If strTipOrd = "C" Then ValidLimites strCodEmpr, Convertyyyymmdd(dtpFechaOrden.Text), CDbl(txtTipoCambio.Text), strCodFile, strCodFon
'
'    '*** Si todo pasó OK ***
    PosicionLimites = True
    
End Function
Public Sub Salir()

    Unload Me
    
End Sub

Private Sub cboAgente_Click()

    strCodAgente = Valor_Caracter
    If cboAgente.ListIndex < 0 Then Exit Sub
    
    strCodAgente = Trim(arrAgente(cboAgente.ListIndex))
    
    Call cboConceptoCosto_Click
    
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
    
    strCodTipoCostoBolsa = Valor_Caracter: strCodTipoCostoConasev = Valor_Caracter
    strCodTipoCavali = Valor_Caracter: strCodTipoCostoFondo = Valor_Caracter: strCodTipoCostoFondoG = Valor_Caracter
    dblComisionBolsa = 0: dblComisionConasev = 0: dblComisionAgente = 0: dblComisionAgenteMin = 0
    dblComisionCavali = 0: dblComisionFondo = 0: dblComisionFondoG = 0
        
    With adoComm
        Set adoRegistro = New ADODB.Recordset
                
        .CommandText = "SELECT CodCosto,TipoCosto,ValorCosto,ValorAlterno FROM CostoNegociacion WHERE " & _
        "TipoOperacion='" & strCodConcepto & "' AND TipoValor='" & Codigo_Valor_RentaVariable & "' AND " & _
        "(CodAgente = '" & Valor_Caracter & "' OR CodAgente = '" & strCodAgente & "') " & _
        "ORDER BY CodCosto"
        Set adoRegistro = .Execute

        Do Until adoRegistro.EOF
            Select Case Trim(adoRegistro("CodCosto"))
                
                Case Codigo_Costo_Agente
                    strCodTipoCostoAgente = Trim(adoRegistro("TipoCosto"))
                    dblComisionAgente = CDbl(adoRegistro("ValorCosto"))
                    dblComisionAgenteMin = CDbl(adoRegistro("ValorAlterno"))
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
                Case Codigo_Costo_FGarantia
                    strCodTipoCostoFondoG = Trim(adoRegistro("TipoCosto"))
                    dblComisionFondoG = CDbl(adoRegistro("ValorCosto"))
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
        "WHERE TipoValor='" & Codigo_Valor_RentaVariable & "' AND IndInstrumento='X' AND IndVigente='X' AND " & _
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
            '/**/
            'Dim strCodMonedaPago  As String
'            strCodMonedaPago = Trim(adoRegistro("CodMoneda"))
            'txtTipoCambioConversion.Text = CStr(ObtenerTipoCambio(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, dtpFechaOrden.Value, strCodMoneda))
            '/**/
            
            If CDbl(txtTipoCambio.Text) = 0 Then txtTipoCambio.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, DateAdd("d", -1, dtpFechaOrden.Value), strCodMoneda, Codigo_Moneda_Local))
            dblTipoCambio = CDbl(txtTipoCambio.Text)
'            txtTipoCambio.Text = CStr(dblTipoCambio)
                        
            gstrPeriodoActual = Format(Year(gdatFechaActual), "0000")
            gstrMesActual = Format(Month(gdatFechaActual), "00")
            
            frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)

            'ACTUALIZA PARAMETROS GLOBALES POR FONDO
            If Not CargarParametrosGlobales(strCodFondoOrden) Then Exit Sub
            
            Me.Refresh
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    '*** Tipo de Instrumento ***
    strSQL = "SELECT FIF.CodFile CODIGO,DescripFile DESCRIP " & _
        "FROM FondoInversionFile FIF JOIN InversionFile IVF ON(IVF.CodFile=FIF.CodFile) " & _
        "WHERE TipoValor='" & Codigo_Valor_RentaVariable & "' AND IndInstrumento='X' AND IndVigente='X' AND " & _
        "CodFondo='" & strCodFondoOrden & "' AND CodAdministradora='" & gstrCodAdministradora & "' ORDER BY DescripFile"
    CargarControlLista strSQL, cboTipoInstrumentoOrden, arrTipoInstrumentoOrden(), Sel_Defecto
            
End Sub




Private Sub cboNegociacion_Click()

    Dim adoConsulta As ADODB.Recordset
    
    strCodNegociacion = Valor_Caracter
    If cboNegociacion.ListIndex < 0 Then Exit Sub
    
    strCodNegociacion = Trim(arrNegociacion(cboNegociacion.ListIndex))
    
    cboConceptoCosto.ListIndex = -1
    If cboConceptoCosto.ListCount > 0 Then cboConceptoCosto.ListIndex = 0
    
    cboConceptoCosto.Enabled = False
    If strCodNegociacion = Codigo_Mecanismo_Rueda Then cboConceptoCosto.Enabled = True
    
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


Private Sub cboOperacion_Click()

    strCodOperacion = Valor_Caracter
    If cboOperacion.ListIndex < 0 Then Exit Sub
    
    strCodOperacion = Trim(arrOperacion(cboOperacion.ListIndex))
    
End Sub



Private Sub cboTipoInstrumentoOrden_Click()

    strCodTipoInstrumentoOrden = Valor_Caracter
    If cboTipoInstrumentoOrden.ListIndex < 0 Then Exit Sub
    
    strCodTipoInstrumentoOrden = Trim(arrTipoInstrumentoOrden(cboTipoInstrumentoOrden.ListIndex))

    lblAnalitica.Caption = strCodTipoInstrumentoOrden & " - ????????"
    strCodFile = strCodTipoInstrumentoOrden

    '*** Tipo de Orden ***
    strSQL = "SELECT IFTON.CodTipoOperacion CODIGO,DescripTipoOperacion DESCRIP " & _
        "FROM InversionFileTipoOperacionNegociacion IFTON JOIN TipoOperacionNegociacion TON ON(TON.CodTipoOperacion=IFTON.CodTipoOperacion)" & _
        "WHERE IFTON.CodFile='" & strCodTipoInstrumentoOrden & "' ORDER BY DescripTipoOperacion"
    CargarControlLista strSQL, cboTipoOrden, arrTipoOrden(), Valor_Caracter

    If cboTipoOrden.ListCount > 0 Then cboTipoOrden.ListIndex = 0
    
    '*** Clase de Instrumento ***
    ' /* strSQL = "SELECT CodDetalleFile CODIGO,DescripDetalleFile DESCRIP FROM InversionDetalleFile WHERE CodFile='" & strCodTipoInstrumentoOrden & "' AND IndVigente='X'  ORDER BY DescripDetalleFile" */
    strSQL = "SELECT CodDetalleFile CODIGO,DescripDetalleFile DESCRIP FROM InversionDetalleFile WHERE CodFile='" & strCodTipoInstrumentoOrden & "' AND IndVigente='X' AND SUBSTRING (LTRIM(RTRIM(DescripDetalleFile)),1,1) <> '6'  ORDER BY DescripDetalleFile"
    CargarControlLista strSQL, cboClaseInstrumento, arrClaseInstrumento(), Sel_Defecto
    
    If cboClaseInstrumento.ListCount > 0 Then
        cboClaseInstrumento.ListIndex = 0
        cboClaseInstrumento.Enabled = True
    End If
    
End Sub


Private Sub cboTipoOrden_Click()

    strCodTipoOrden = Valor_Caracter
    If cboTipoOrden.ListIndex < 0 Then Exit Sub

    strCodTipoOrden = Trim(arrTipoOrden(cboTipoOrden.ListIndex))

    Me.MousePointer = vbHourglass
    Select Case strCodTipoOrden
        Case Codigo_Orden_Compra
            strSQL = "SELECT CodTitulo CODIGO,(CONVERT(char(15),Nemotecnico) + ' ' + DescripTitulo)DESCRIP FROM InstrumentoInversion " & _
                "WHERE CodFile='" & strCodFile & "' AND CodDetalleFile='" & strCodClaseInstrumento & "' AND IndVigente='X' ORDER BY Nemotecnico"
            CargarControlLista strSQL, cboTitulo, arrTitulo(), Sel_Defecto
        
            If cboTitulo.ListCount > 0 Then cboTitulo.ListIndex = 0
                        
            txtTasaMensual.Enabled = False
            fraComisionMontoFL2.Visible = False
            
        Case Codigo_Orden_Venta
            strSQL = "SELECT II.CodTitulo CODIGO," & _
                "(RTRIM(II.Nemotecnico) + ' ' + RTRIM(II.DescripTitulo)) DESCRIP " & _
                "FROM InstrumentoInversion II JOIN InversionKardex IK ON(IK.CodTitulo=II.CodTitulo)" & _
                "WHERE SaldoFinal > 0 AND IndUltimoMovimiento='X' AND " & _
                "II.CodFile='" & strCodFile & "' AND II.CodDetalleFile='" & strCodClaseInstrumento & "' AND " & _
                "IK.CodFondo='" & strCodFondoOrden & "' AND IK.CodAdministradora='" & gstrCodAdministradora & "' " & _
                "ORDER BY II.Nemotecnico"
            CargarControlLista strSQL, cboTitulo, arrTitulo(), Sel_Defecto
        
            If cboTitulo.ListCount > 0 Then cboTitulo.ListIndex = 0
                                    
            txtTasaMensual.Enabled = False
            fraComisionMontoFL2.Visible = False
                            
    End Select
    Me.MousePointer = vbDefault
    
End Sub


Private Sub cboTitulo_Click()

    Dim adoRegistro     As ADODB.Recordset
    Dim intRegistro     As Integer
    
    strCodGarantia = Valor_Caracter: txtDescripOrden.Text = Valor_Caracter: strCodAnalitica = Valor_Caracter
    lblAnalitica = strCodTipoInstrumentoOrden & "-????????": lblValorNominal.Caption = "0"
    lblStockNominal.Caption = "0"
    strCodEmisor = Valor_Caracter: strCodGrupo = Valor_Caracter
    If cboTitulo.ListIndex < 0 Then Exit Sub

    strCodGarantia = Trim(arrTitulo(cboTitulo.ListIndex))
    
    If strCodGarantia <> Valor_Caracter Then tabRentaVariable.TabEnabled(2) = True: tabRentaVariable.TabEnabled(3) = True

    With adoComm
        Set adoRegistro = New ADODB.Recordset

'/**/ se agrego al select CodMoneda1

        .CommandText = "SELECT CodAnalitica,ValorNominal,CodMoneda,CodEmisor,CodGrupo,Nemotecnico,CodMoneda1 " & _
            "FROM InstrumentoInversion WHERE CodTitulo='" & strCodGarantia & "'"
        Set adoRegistro = .Execute

        If Not adoRegistro.EOF Then
            strCodAnalitica = Trim(adoRegistro("CodAnalitica"))
            lblAnalitica.Caption = strCodFile & "-" & strCodAnalitica
            lblValorNominal.Caption = CStr(adoRegistro("ValorNominal"))
            
            MonedaPago = CStr(adoRegistro("CodMoneda1")) '/**/
            
            lblMoneda.Caption = ObtenerDescripcionMoneda(adoRegistro("CodMoneda"))
            lblMoneda.Tag = adoRegistro("CodMoneda")
            
            'txtTipoCambio.Text = CStr(ObtenerTipoCambio(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, dtpFechaOrden.Value, adoRegistro("CodMoneda")))
            
            txtTipoCambio.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, dtpFechaOrden.Value, adoRegistro("CodMoneda"), Codigo_Moneda_Local))
                     
            If CDbl(txtTipoCambio.Text) = 0 Then txtTipoCambio.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, DateAdd("d", -1, dtpFechaOrden.Value), adoRegistro("CodMoneda"), Codigo_Moneda_Local))
            dblTipoCambio = CDbl(txtTipoCambio.Text)
            
            lblDescripMoneda(0) = "S/.": lblDescripMoneda(0).Tag = Codigo_Moneda_Local
            lblDescripMoneda(1) = "S/.": lblDescripMoneda(1).Tag = Codigo_Moneda_Local
            lblDescripMoneda(2) = "S/.": lblDescripMoneda(2).Tag = Codigo_Moneda_Local
            lblDescripMoneda(3) = "S/.": lblDescripMoneda(3).Tag = Codigo_Moneda_Local
                       
                       
            lblDescripMonedaResumen(0) = "S/.": lblDescripMonedaResumen(0).Tag = Codigo_Moneda_Local
            lblDescripMonedaResumen(1) = "S/.": lblDescripMonedaResumen(1).Tag = Codigo_Moneda_Local
                    
            '/**/
            'Se cambio el CodMoneda por CodMoneda1 ya que esta es el tipo de moneda a pagar
            lblDescripMoneda(0).Caption = ObtenerSignoMoneda(adoRegistro("CodMoneda1")) '/**/
            lblDescripMoneda(0).Tag = adoRegistro("CodMoneda1") '/**/
            lblDescripMoneda(1).Caption = ObtenerSignoMoneda(adoRegistro("CodMoneda"))
            lblDescripMoneda(1).Tag = adoRegistro("CodMoneda")
            
            lblDescripMoneda(2).Caption = lblDescripMoneda(0).Caption
            lblDescripMoneda(2).Tag = lblDescripMoneda(0).Tag
            
            lblDescripMoneda(3).Caption = lblDescripMoneda(1).Caption
            lblDescripMoneda(3).Tag = lblDescripMoneda(1).Tag
            
            lblDescripMoneda(4).Caption = lblDescripMoneda(1).Caption
            lblDescripMoneda(4).Tag = lblDescripMoneda(1).Tag
            
            lblDescripMoneda(5).Caption = lblDescripMoneda(0).Caption
            lblDescripMoneda(5).Tag = lblDescripMoneda(0).Tag
           
           
            lblDescripMonedaResumen(0).Caption = ObtenerSignoMoneda(adoRegistro("CodMoneda"))
            lblDescripMonedaResumen(0).Tag = adoRegistro("CodMoneda")
            lblDescripMonedaResumen(1).Caption = ObtenerSignoMoneda(adoRegistro("CodMoneda"))
            lblDescripMonedaResumen(1).Tag = adoRegistro("CodMoneda")
                        
      
            lblMonedaPago.Caption = ObtenerDescripcionMoneda(adoRegistro("CodMoneda1")) ' /**/
                        
            lblMonOrigenAMonPago.Caption = "(" + ObtenerCodSignoMoneda(Codigo_Moneda_Local) + "/" + ObtenerCodSignoMoneda(adoRegistro("CodMoneda")) + ")"        'Mid( ObtenerDescripcionDeMonedaOrigenAMonedaPago(strCodAnalitica), 1, 7)
            lblMonPagoAMonOrigen.Caption = "(" + ObtenerCodSignoMoneda(adoRegistro("CodMoneda1")) + "/" + ObtenerCodSignoMoneda(adoRegistro("CodMoneda")) + ")"  'Mid(ObtenerDescripcionDeMonedaOrigenAMonedaPago(strCodAnalitica), 8, 15)
                        
            strCodEmisor = Trim(adoRegistro("CodEmisor")): strCodGrupo = Trim(adoRegistro("CodGrupo"))
            strNemonico = Trim(adoRegistro("Nemotecnico"))
            
            Call IniciarComisiones
            
            'inicializar precios
            txtDescripOrden.Text = Valor_Caracter
            txtObservacion.Text = Valor_Caracter
            txtPrecioUnitario(0).Text = "0"
            txtPrecioUnitario(1).Text = "0"
            txtCantidad.Text = "0"
            'txtSubTotalMonedaPago.Text = "0.00"
            'lblMontoTotalMonedaPago.Caption = "0.00"
            lblSubTotal(0).Caption = "0.00"
            lblSubTotal(1).Caption = "0.00"
                        
            '/* SI moneda de pago = moneda de emision*/
            
            If MonedaPago = lblMoneda.Tag Then
            
                lblDescrip(73).Visible = False
                lblDescrip(75).Visible = False
                lblDescripMoneda(0).Visible = False
                'txtSubTotalMonedaPago.Visible = False
                chkAplicarMon(0).Visible = False
                lblMonPagoAMonOrigen.Visible = False
                lblComisionAgenteOrigen.ForeColor = lblComisionAgenteOrigen.BackColor
                lblComisionBolsaOrigen.ForeColor = lblComisionBolsaOrigen.BackColor
                lblComisionCavaliOrigen.ForeColor = lblComisionCavaliOrigen.BackColor
                lblComisionFondoOrigen.ForeColor = lblComisionFondoOrigen.BackColor
                lblComisionFondoGOrigen.ForeColor = lblComisionFondoGOrigen.BackColor
                lblComisionConasevOrigen.ForeColor = lblComisionConasevOrigen.BackColor
                lblComisionGastoBancarioOrigen.ForeColor = lblComisionGastoBancarioOrigen.BackColor
                lblComisionxAccionOrigen.ForeColor = lblComisionxAccionOrigen.BackColor
                lblComisionIgvOrigen.ForeColor = lblComisionIgvOrigen.BackColor
                
            Else

                lblDescrip(73).Visible = True
                lblDescrip(75).Visible = True
                lblDescripMoneda(0).Visible = True
                'txtSubTotalMonedaPago.Visible = True
                chkAplicarMon(0).Visible = True
                lblMonPagoAMonOrigen.Visible = True
                lblComisionAgenteOrigen.ForeColor = &H80000012
                lblComisionBolsaOrigen.ForeColor = &H80000012
                lblComisionCavaliOrigen.ForeColor = &H80000012
                lblComisionFondoOrigen.ForeColor = &H80000012
                lblComisionFondoGOrigen.ForeColor = &H80000012
                lblComisionConasevOrigen.ForeColor = &H80000012
                lblComisionGastoBancarioOrigen.ForeColor = &H80000012
                lblComisionxAccionOrigen.ForeColor = &H80000012
                lblComisionIgvOrigen.ForeColor = &H80000012
                                
            End If
            '/**/
            
        End If
        adoRegistro.Close

        '*** Validar Limites ***
        If Not PosicionLimites() Then Exit Sub

        .CommandText = "SELECT SaldoFinal FROM InversionKardex WHERE CodAnalitica='" & strCodAnalitica & "' AND " & _
            "CodFile='" & strCodFile & "' AND CodFondo='" & strCodFondoOrden & "' AND " & _
            "CodAdministradora='" & gstrCodAdministradora & "' AND IndUltimoMovimiento='X' AND SaldoFinal > 0"
        Set adoRegistro = .Execute

        If Not adoRegistro.EOF Then
            lblStockNominal.Caption = CStr(adoRegistro("SaldoFinal"))
        End If
        adoRegistro.Close
        
        .CommandText = "SELECT PrecioCierre FROM InstrumentoPrecioTir WHERE CodTitulo='" & strCodGarantia & "' AND " & _
            "IndUltimoPrecio='X'"
        Set adoRegistro = .Execute

        If Not adoRegistro.EOF Then
            lblUltimoPrecio.Caption = CStr(adoRegistro("PrecioCierre"))
        End If
        adoRegistro.Close: Set adoRegistro = Nothing

    End With
    
    txtDescripOrden.Text = UCase(Trim(cboTipoOrden.Text) & " " & Trim(cboTipoInstrumentoOrden.Text) & " - " & Left(cboTitulo.Text, 15))

    
End Sub





Private Sub IniciarComisiones()

    Dim intContador As Integer
    
    For intContador = 0 To 1
        txtComisionAgente(intContador).Text = "0"
        txtComisionBolsa(intContador).Text = "0"
        txtComisionCavali(intContador).Text = "0"
        txtComisionFondo(intContador).Text = "0"
        txtComisionFondoG(intContador).Text = "0"
        txtComisionConasev(intContador).Text = "0"
        txtComisionIgv(intContador).Text = "0"
        
        '/* para las cajas de texto agregadas*/
        txtComisionGastoBancario(intContador).Text = "0"
        txtComisionxAccion(intContador).Text = "0"
        '/**/
        
        txtPorcenAgente(intContador).Text = "0"
        lblPorcenBolsa(intContador).Caption = "0"
        lblPorcenCavali(intContador).Caption = "0"
        lblPorcenFondo(intContador).Caption = "0"
        lblPorcenFondoG(intContador).Caption = "0"
        lblPorcenConasev(intContador).Caption = "0"
        lblPorcenIgv(intContador).Caption = CStr(gdblTasaIgv)
        
        '/*para las cajas de texto agregadas*/
        lblPorcenGastoBancario(intContador).Caption = "0"
        lblPorcenComisionAccion(intContador).Caption = "0"
        '/**/
        
        lblPrecioResumen(intContador).Caption = "0"
        lblSubTotalResumen(intContador).Caption = "0"
        lblComisionesResumen(intContador).Caption = "0"
        lblTotalResumen(intContador).Caption = "0"
    Next
    
End Sub
Private Sub AplicarCostos(Index As Integer)
    
    
    lblPorcenIgv(Index).Caption = CStr(gdblTasaIgv)
    
    If strCodOrigen = Codigo_Mercado_Extranjero Then
        lblPorcenIgv(Index).Caption = "0"
    End If
    
    If strCodTipoCostoAgente = Codigo_Tipo_Costo_Monto Then
        txtComisionAgente(Index).Text = CStr(dblComisionAgente)
    Else
        If CDbl(lblSubTotal(Index).Caption) * (dblComisionAgente / 100) >= dblComisionAgenteMin Then
            AsignaComision strCodTipoCostoAgente, dblComisionAgente, txtComisionAgente(Index)
        Else
            txtComisionAgente(Index).Text = CStr(dblComisionAgenteMin)
        End If
    End If
    
    If strCodTipoCostoBolsa = Codigo_Tipo_Costo_Monto Then
        txtComisionBolsa(Index).Text = CStr(dblComisionBolsa)
    Else
        AsignaComision strCodTipoCostoBolsa, dblComisionBolsa, txtComisionBolsa(Index)
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
    
    If strCodTipoCostoFondoG = Codigo_Tipo_Costo_Monto Then
        txtComisionFondoG(Index).Text = CStr(dblComisionFondoG)
    Else
        AsignaComision strCodTipoCostoFondoG, dblComisionFondoG, txtComisionFondoG(Index)
    End If
    
    If strCodTipoCavali = Codigo_Tipo_Costo_Monto Then
        txtComisionCavali(Index).Text = CStr(dblComisionCavali)
    Else
        AsignaComision strCodTipoCavali, dblComisionCavali, txtComisionCavali(Index)
    End If
                         
    Call CalculoImpuesto(Index)

    Call CalculoTotal(Index)
    
End Sub

Private Sub AsignaComision(strTipoComision As String, dblValorComision As Double, ctrlValorComision As Control)
    
    If Not IsNumeric(lblSubTotal(ctrlValorComision.Index).Caption) Then Exit Sub
    
    If dblValorComision > 0 Then
'        If Val(txtSubTotalMonedaPago.Text) = 0 Then
            ctrlValorComision.Text = CStr(CCur(lblSubTotal(ctrlValorComision.Index)) * dblValorComision / 100)
'        Else
'            ctrlValorComision.Text = CStr(CCur(txtSubTotalMonedaPago.Text) * dblValorComision / 100)
'                If InStr(1, CStr(CCur(txtSubTotalMonedaPago.Text) * dblValorComision / 100), ".") > 0 Then
'                    ctrlValorComision.Text = CStr(Mid(CStr(CCur(txtSubTotalMonedaPago.Text) * dblValorComision / 100), 1, InStr(1, CStr(CCur(txtSubTotalMonedaPago.Text) * dblValorComision / 100), ".") + 2))
'                Else
'                    ctrlValorComision.Text = CStr(CCur(txtSubTotalMonedaPago.Text) * dblValorComision / 100)
'                End If
'        End If
    End If
            
End Sub

Private Sub chkAplicar_Click(Index As Integer)

    If chkAplicar(Index).Value Then
        Call AplicarCostos(Index)
    Else
        Call IniciarComisiones
        Call CalculoImpuesto(Index)
        Call CalculoTotal(Index)
    End If

End Sub

'/* */
Private Sub chkAplicarMon_Click(Index As Integer)

    If Me.chkAplicarMon(Index).Value Then
        AplicarTipoCambioPactado = True
        lblDescripMoneda(2).Caption = lblDescripMoneda(0).Caption
        lblDescripMoneda(2).Tag = lblDescripMoneda(0).Tag
    Else
        AplicarTipoCambioPactado = False
        lblDescripMoneda(2).Caption = lblDescripMoneda(1).Caption
        lblDescripMoneda(2).Tag = lblDescripMoneda(1).Tag
    End If

End Sub
'/* */


Private Sub cmdCalculo_Click()

    Dim dblFactor As Double

    '*** Tir Bruta ***
'    If CInt(txtDiasPlazo.Text) > 0 And CCur(lblSubTotal(0).Caption) > 0 Then
'        If CCur(lblSubTotal(1).Caption) = 0 Or CCur(lblSubTotal(0).Caption) = 0 Then
'            MsgBox "Por favor verificar que el SubTotal al Contado y a Plazo tengan valores.", vbExclamation, Me.Caption
'            Exit Sub
'        End If
'        dblFactor = (CCur(lblSubTotal(1).Caption) / CCur(lblSubTotal(0).Caption)) ^ (365 / CInt(txtDiasPlazo.Text))
'        lblTirBruta.Caption = CStr((dblFactor - 1) * 100)
'    End If
'
'    '*** Tir Neta ***
'    If CInt(txtDiasPlazo.Text) > 0 And CCur(lblMontoTotal(0).Caption) > 0 Then
'        dblFactor = (CCur(lblMontoTotal(1).Caption) / CCur(lblMontoTotal(0).Caption)) ^ (365 / CInt(txtDiasPlazo.Text))
'        lblTirNeta.Caption = CStr((dblFactor - 1) * 100)
'    End If
    
End Sub

Private Sub cmdEnviar_Click()

    Dim strFechaDesde       As String, strFechaHasta        As String
    Dim intRegistro         As Integer, intContador         As Integer
    Dim datFecha            As Date
    
    If adoConsulta.RecordCount = 0 Then Exit Sub
    
     If Not IsNull(dtpFechaOrdenDesde.Value) Or Not IsNull(dtpFechaOrdenHasta.Value) Then
        strFechaDesde = Convertyyyymmdd(dtpFechaOrdenDesde.Value)
        datFecha = DateAdd("d", 1, dtpFechaOrdenHasta.Value)
        strFechaHasta = Convertyyyymmdd(datFecha)
    End If
    
    If Not IsNull(dtpFechaLiquidacionDesde.Value) Or Not IsNull(dtpFechaLiquidacionHasta.Value) Then
        strFechaDesde = Convertyyyymmdd(dtpFechaLiquidacionDesde.Value)
        datFecha = DateAdd("d", 1, dtpFechaLiquidacionHasta.Value)
        strFechaHasta = Convertyyyymmdd(datFecha)
    End If
    

    
    intContador = tdgConsulta.SelBookmarks.Count - 1
    
    If intContador < 0 Then
        MsgBox "No se ha seleccionado ningún registro", vbCritical, Me.Caption
        Exit Sub
    End If
        
    For intRegistro = 0 To intContador
        'tdgConsulta.Row = tdgConsulta.SelBookmarks.Count - 1 'tdgConsulta.SelBookmarks(intRegistro) - 1
               
        adoConsulta.MoveFirst
        
        adoConsulta.Move CLng(tdgConsulta.SelBookmarks(intRegistro) - 1), 0
                        
        tdgConsulta.Refresh
               
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

Private Sub dtpFechaLiquidacion_Change()

    If dtpFechaLiquidacion.Value < dtpFechaOrden.Value Then
        dtpFechaLiquidacion.Value = dtpFechaOrden.Value
    Else
        If Not EsDiaUtil(dtpFechaLiquidacion.Value) Then
            MsgBox "La Fecha no es un día útil...se cambiará por una fecha correcta !", vbInformation, Me.Caption
            dtpFechaLiquidacion.Value = ProximoDiaUtil(dtpFechaLiquidacion.Value)
        End If
    End If
    lblFechaLiquidacion.Caption = CStr(dtpFechaLiquidacion.Value)
    
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
    
    '/**/
        AplicarTipoCambioPactado = False
    '/**/
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

Public Sub Imprimir()

End Sub

Public Sub SubImprimir(Index As Integer)

    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    Dim strFechaDesde           As String, strFechaHasta        As String
    Dim strSeleccionRegistro    As String

    If tabRentaVariable.Tab = 1 Then Exit Sub
    
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
Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Ordenes de Inversión"
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Text = "Papeleta de Inversión"
    
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
        
    '*** Agentes de Bolsa ***
    strSQL = "SELECT CodPersona CODIGO,DescripPersona DESCRIP FROM InstitucionPersona WHERE TipoPersona='" & Codigo_Tipo_Persona_Agente & "' ORDER BY DescripPersona"
    CargarControlLista strSQL, cboAgente, arrAgente(), Sel_Defecto
    
    If cboAgente.ListCount > 0 Then cboAgente.ListIndex = 0
    
    '*** Tipo Liquidación Operación ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPLIQ' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboOperacion, arrOperacion(), Valor_Caracter

    If cboOperacion.ListCount > 0 Then cboOperacion.ListIndex = 0
    
    '*** Mecanismos de Negociación ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='MECNEG' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboNegociacion, arrNegociacion(), Valor_Caracter

    If cboNegociacion.ListCount > 0 Then cboNegociacion.ListIndex = 0
    
    '*** Conceptos de Costos de Negociación ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPCCO' AND ValorParametro='RV' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboConceptoCosto, arrConceptoCosto(), Sel_Defecto
    
    '*** Mercado de Negociación ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='MDONEG' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboOrigen, arrOrigen(), Valor_Caracter
        
    cboTipoVaR.AddItem "Histórico"
    cboTipoVaR.AddItem "Montecarlo"
    
    cboPeriodo.AddItem "Días"
    cboPeriodo.AddItem "Meses"
    cboPeriodo.AddItem "Años"

End Sub

Private Sub InicializarValores()

    Dim adoRegistro As ADODB.Recordset
    
    strEstado = Reg_Defecto
    tabRentaVariable.Tab = 0

    'Por Defecto
    strTipoDocumento = Codigo_Tipo_Comprobante_Pago_Poliza

    dtpFechaOrdenDesde.Value = gdatFechaActual
    dtpFechaOrdenHasta.Value = dtpFechaOrdenDesde.Value
    dtpFechaLiquidacionDesde.Value = Null
    dtpFechaLiquidacionHasta.Value = Null
    
    Set adoRegistro = New ADODB.Recordset
    With adoComm
        .CommandText = "SELECT CodFile FROM InversionFile " & _
            "WHERE TipoValor='" & Codigo_Valor_RentaVariable & "' AND IndInstrumento='X' AND IndVigente='X' " & _
            "ORDER BY DescripFile"
        Set adoRegistro = .Execute
                
        strCodigosFile = Valor_Caracter
        Do While Not adoRegistro.EOF
            If strCodigosFile <> Valor_Caracter Then strCodigosFile = strCodigosFile & ",'"
            
            strCodigosFile = strCodigosFile & Trim(adoRegistro("CodFile")) & "'"
        
            adoRegistro.MoveNext
        Loop
        adoRegistro.Close: Set adoRegistro = Nothing
        
        strCodigosFile = "('" & strCodigosFile & ")"
    End With
    
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgConsulta.Columns(1).Width = tdgConsulta.Width * 0.01 * 9
    tdgConsulta.Columns(8).Width = tdgConsulta.Width * 0.01 * 10
    tdgConsulta.Columns(3).Width = tdgConsulta.Width * 0.01 * 60
    tdgConsulta.Columns(0).Width = tdgConsulta.Width * 0.01 * 6
    tdgConsulta.Columns(4).Width = tdgConsulta.Width * 0.01 * 14
    tdgConsulta.Columns(6).Width = tdgConsulta.Width * 0.01 * 10
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    
End Sub
Private Function ValidaDatosGenerales() As Boolean

'    ValidaDatosGenerales = False
'
'    If cboFondo.ListIndex < 0 Then
'        MsgBox "Debe seleccionar el Fondo.", vbCritical, Me.Caption
'        If cboFondo.Enabled Then cboFondo.SetFocus
'        Exit Function
'    End If
'
'    If cboAgente.ListIndex < 0 Then
'        MsgBox "Debe seleccionar la Sociedad Agente de Bolsa.", vbCritical, Me.Caption
'        If cboAgente.Enabled Then cboAgente.SetFocus
'        Exit Function
'    End If
'
'    If cboNegociacion.ListIndex < 0 Then
'        MsgBox "Debe seleccionar el Mecanismo de Negociación.", vbCritical, Me.Caption
'        If cboNegociacion.Enabled Then cboNegociacion.SetFocus
'        Exit Function
'    End If
'
'    If cboTipoInstrumento.ListIndex < 0 Then
'        MsgBox "Debe seleccionar el Tipo de Instrumento de Inversión.", vbCritical, Me.Caption
'        If cboTipoInstrumento.Enabled Then cboTipoInstrumento.SetFocus
'        Exit Function
'    End If
'
'    If cboTipoOrden.ListIndex < 0 Then
'        MsgBox "Debe seleccionar el Tipo de ORDEN.", vbCritical, Me.Caption
'        If cboTipoOrden.Enabled Then cboTipoOrden.SetFocus
'        Exit Function
'    End If
'
'    If cboTitulo.ListIndex < 0 Then
'        MsgBox "Debe seleccionar el Título.", vbCritical, Me.Caption
'        If cboTitulo.Enabled Then cboTitulo.SetFocus
'        Exit Function
'    End If
'
'    If Trim(txtDscOrd.Text) = "" Then
'        MsgBox "Debe indicar la Descripción de la ORDEN.", vbCritical, Me.Caption
'        If txtDscOrd.Enabled Then txtDscOrd.SetFocus
'        Exit Function
'    End If
'
'    If txtCantOrden.Text = 0 Then
'        MsgBox "Debe indicar la Cantidad de Acciones a Negociar.", vbCritical, Me.Caption
'        If txtCantOrden.Enabled Then txtCantOrden.SetFocus
'        Exit Function
'    End If
'
'    If txtPrecioUnitario.Text = 0 Then
'        MsgBox "Debe indicar el Precio por cada Acción Negociada.", vbCritical, Me.Caption
'        If txtPrecioUnitario.Enabled Then txtPrecioUnitario.SetFocus
'        Exit Function
'    End If
'
'    If txtTipoCambio.Text = 0 Then
'        MsgBox "Debe indicar el Precio por cada Acción Negociada.", vbCritical, Me.Caption
'        If txtTipoCambio.Enabled Then txtTipoCambio.SetFocus
'        Exit Function
'    End If
'
'    '*** Validación de STOCK DE ACCIONES ***
'    If strTipOrd = "V" Then
'        If CLng(txtCantOrden.Text) > CLng(Format(lblSaldo.Caption, "0")) Then
'            MsgBox "Stock insuficiente para Registrar la Orden de Venta.", vbCritical, Me.Caption
'            If txtCantOrden.Enabled Then txtCantOrden.SetFocus
'            Exit Function
'        End If
'    End If
'
'    '*** Si todo paso ok ***
'    ValidaDatosGenerales = True
  
End Function

Private Sub CalculoTotal(Index As Integer)

    Dim curComImp As Currency, curMonTotal As Currency
    Dim curComImpOrigen As Currency, curMonTotalOrigen As Currency

    If Not (IsNumeric(txtComisionAgente(Index).Text) And IsNumeric(txtComisionBolsa(Index).Text) And IsNumeric(txtComisionConasev(Index).Text) And IsNumeric(txtComisionCavali(Index).Text) And IsNumeric(txtComisionFondo(Index).Text) And IsNumeric(txtComisionFondoG(Index).Text) And IsNumeric(txtComisionxAccion(Index).Text) And IsNumeric(txtComisionIgv(Index).Text)) Then Exit Sub
    
    Call CalculoImpuesto(Index)
    
    curComImp = CCur(txtComisionAgente(Index).Text) + CCur(txtComisionBolsa(Index).Text) + CCur(txtComisionConasev(Index).Text) + CCur(txtComisionCavali(Index).Text) + CCur(txtComisionFondo(Index).Text) + CCur(txtComisionFondoG(Index).Text) + CCur(txtComisionxAccion(Index).Text) + CCur(txtComisionIgv(Index).Text)

    lblComisionesResumen(Index).Caption = CStr(curComImp)
    
    If strCodTipoOrden = Codigo_Orden_Compra Then  '*** Compra ***
        If Index = 0 Then
            curMonTotal = CCur(lblSubTotal(Index).Caption) + curComImp
            curMonTotalOrigen = curMonTotal 'CCur(txtSubTotalMonedaPago.Text) + curComImp
        Else
            curMonTotal = CCur(lblSubTotal(Index).Caption) - curComImp
            curMonTotalOrigen = curMonTotal 'CCur(txtSubTotalMonedaPago.Text) - curComImp
        End If
    ElseIf strCodTipoOrden = Codigo_Orden_Venta Then '*** Venta ***
        curMonTotal = CCur(lblSubTotal(Index).Caption) - curComImp
        curMonTotalOrigen = curMonTotal 'CCur(txtSubTotalMonedaPago.Text) - curComImp
    End If
        
    lblMontoTotal(Index).Caption = CStr(curMonTotal)
    'lblMontoTotalMonedaPago.Caption = CStr(curMonTotalOrigen)
    
End Sub
Private Sub ActualizaMontosOrigen(Index As Integer)

Dim dblFactor As Double

'If CDbl(lblSubTotal(Index).Caption) > 0 Then
'   dblFactor = CDbl(txtSubTotalMonedaPago.Text) / CDbl(lblSubTotal(Index).Caption)
'Else
   dblFactor = 1
'End If

lblComisionAgenteOrigen.Caption = Format(CCur(txtComisionAgente(Index).Text) / dblFactor, "0.00")
lblComisionBolsaOrigen.Caption = Format(CCur(txtComisionBolsa(Index).Text) / dblFactor, "0.00")
lblComisionCavaliOrigen.Caption = Format(CCur(txtComisionCavali(Index).Text) / dblFactor, "0.00")
lblComisionFondoOrigen.Caption = Format(CCur(txtComisionFondo(Index).Text) / dblFactor, "0.00")
lblComisionFondoGOrigen.Caption = Format(CCur(txtComisionFondoG(Index).Text) / dblFactor, "0.00")
lblComisionConasevOrigen.Caption = Format(CCur(txtComisionConasev(Index).Text) / dblFactor, "0.00")
lblComisionGastoBancarioOrigen.Caption = Format(CCur(txtComisionGastoBancario(Index).Text) / dblFactor, "0.00")
lblComisionxAccionOrigen.Caption = Format(CCur(txtComisionxAccion(Index).Text) / dblFactor, "0.00")
lblComisionIgvOrigen.Caption = Format(txtComisionIgv(Index).Text / dblFactor, "0.00")



End Sub


Private Sub Form_Unload(Cancel As Integer)

    Set frmOrdenRentaVariable = Nothing
    Call OcultarReportes
    frmMainMdi.stbMdi.Panels(3).Text = "Acción"
    
End Sub


Private Sub dtpFechaLiquidacion_LostFocus()

'    Dim intRes As Integer
'
'    If Not LEsDiaUtil(dtpFechaLiquidacion) Then
'       dtpFechaLiquidacion.Text = LProxDiaUtil(dtpFechaLiquidacion.Text)
'    End If
'
'    If CVDate(dtpFechaOrden.Text) > CVDate(dtpFechaLiquidacion.Text) Then
'       MsgBox "Fecha de Liquidación debe ser posterior a la Fecha de Operación", vbCritical
'       dtpFechaLiquidacion.Text = dtpFechaOrden.Text
'       dtpFechaLiquidacion.SetFocus
'    End If
    
End Sub




Private Sub lblCantidadResumen_Change()

    Call FormatoMillarEtiqueta(lblCantidadResumen, Decimales_Monto)
    
End Sub

Private Sub lblComisionAgente_Click()

End Sub

Private Sub lblComisionesResumen_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblComisionesResumen(Index), Decimales_Monto)
    
End Sub

'Private Sub lblComisionIgv_Change(Index As Integer)
'
'    Call FormatoMillarEtiqueta(lblComisionIgv(Index), Decimales_Monto)
'
'End Sub

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

'/* */

Private Sub lblPorcenComisionAccion_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblPorcenComisionAccion(Index), Decimales_Tasa)

End Sub
'/* */


Private Sub lblPorcenConasev_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblPorcenConasev(Index), Decimales_Tasa)
    
End Sub

Private Sub lblPorcenFondo_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblPorcenFondo(Index), Decimales_Tasa)
    
End Sub

Private Sub lblPorcenFondoG_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblPorcenFondoG(Index), Decimales_Tasa)
    
End Sub

'/**/
Private Sub lblPorcenGastoBancario_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblPorcenGastoBancario(Index), Decimales_Tasa)

End Sub
'/**/


Private Sub lblPorcenIgv_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblPorcenIgv(Index), Decimales_Monto)
    
End Sub

Private Sub lblPrecioResumen_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblPrecioResumen(Index), Decimales_Precio)
    
End Sub

Private Sub lblStockNominal_Change()

    Call FormatoMillarEtiqueta(lblStockNominal, Decimales_Monto)
    
End Sub

Private Sub lblSubTotal_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblSubTotal(Index), Decimales_Monto)
    
    lblSubTotalResumen(Index).Caption = CStr(lblSubTotal(Index).Caption)
    
    'Call txtTipoCambioConversion_Change
End Sub

Private Sub txtComisionAgente_LostFocus(Index As Integer)

    Call CalculoTotal(Index)

End Sub

Private Sub txtComisionIgv_Change(Index As Integer)

    Call FormatoCajaTexto(txtComisionIgv(Index), Decimales_Monto)
    
End Sub

Private Sub txtComisionIgv_KeyPress(Index As Integer, KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtComisionIgv(Index), Decimales_Monto)
    
    If KeyAscii = vbKeyReturn Then
        Call CalculoTotal(Index)
    End If

End Sub

'/**/
Private Sub txtSubTotalMonedaPago_Change() '(Index As Integer)

    'Call FormatoMillarEtiqueta(txtSubTotalMonedaPago, Decimales_Monto)
        'Call FormatoCajaTexto(txtSubTotalMonedaPago, Decimales_Monto) 'Decimales_TipoCambio)

    'lblSubTotalResumen(Index).Caption = CStr(lblSubTotal(Index).Caption)
    
     Call CalculoTotal(0)
    
End Sub
'/**/

Private Sub lblSubTotalResumen_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblSubTotalResumen(Index), Decimales_Monto)
    
End Sub

Private Sub lblTirBrutaResumen_Change()

    Call FormatoMillarEtiqueta(lblTirBrutaResumen, Decimales_Tasa)
    
End Sub

Private Sub lblTirNetaResumen_Change()

    Call FormatoMillarEtiqueta(lblTirNetaResumen, Decimales_Tasa)
    
End Sub

Private Sub lblTotalResumen_Click(Index As Integer)

    Call FormatoMillarEtiqueta(lblTotalResumen(Index), Decimales_Monto)
    
End Sub

Private Sub lblUltimoPrecio_Change()

    Call FormatoMillarEtiqueta(lblUltimoPrecio, Decimales_Precio)
    
End Sub

Private Sub lblValorNominal_Change()

    Call FormatoMillarEtiqueta(lblValorNominal, Decimales_Monto)
    
End Sub

Private Sub tabRentaVariable_Click(PreviousTab As Integer)
   
    Select Case tabRentaVariable.Tab
        Case 1, 2, 3
            'If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vQuery)
            If PreviousTab = 0 And strEstado = Reg_Consulta Then tabRentaVariable.Tab = 0
            If strEstado = Reg_Defecto Then tabRentaVariable.Tab = 0
            If tabRentaVariable.Tab = 2 Then
                fraDatosNegociacion.Caption = "Negociación" & Space(1) & "-" & Space(1) & _
                    Trim(cboTipoOrden.Text) & Space(1) & Trim(Left(cboTitulo.Text, 15))
            End If
    End Select
    
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
    
    Call FormatoCajaTexto(txtCantidad, Decimales_Monto)
    
    If Trim(txtCantidad.Text) = Valor_Caracter Then Exit Sub
    
    If CCur(txtCantidad.Text) > 0 And cboTitulo.ListIndex > 0 Then
        If IsNumeric(txtCantidad.Text) Then
           curCantidad = CCur(txtCantidad.Text)
        Else
           curCantidad = 0
        End If
        
        lblCantidadResumen.Caption = CStr(curCantidad)
                    
        If IsNumeric(txtPrecioUnitario(0).Text) Then dblPreUni = CDbl(txtPrecioUnitario(0).Text) '* 0.01
        
        curSubTotal = curCantidad * dblPreUni
        lblSubTotal(0).Caption = CStr(curSubTotal)
        
        'Si moneda de la inversion es igual a la moneda de pago
'        If MonedaPago = lblMoneda.Tag Then
'            txtSubTotalMonedaPago.Text = lblSubTotal(0).Caption
'        End If
                
        Call CalculoTotal(0)
        
        If strCodTipoOrden = Codigo_Orden_Pacto Then
            If IsNumeric(txtPrecioUnitario(1).Text) Then dblPreUni = CDbl(txtPrecioUnitario(1).Text) '* 0.01
        
            curSubTotal = curCantidad * dblPreUni
            lblSubTotal(1).Caption = CStr(curSubTotal)
        
            Call CalculoTotal(1)
        
        End If
    End If
        
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtCantidad, Decimales_Monto)
    
End Sub


Private Sub ActualizaPorcentaje(ctrlComision As Control, ctrlPorcentaje As Control)

    If Not IsNumeric(ctrlComision) Or Not IsNumeric(lblSubTotal(ctrlComision.Index).Caption) Then Exit Sub
                
    If CCur(lblSubTotal(ctrlComision.Index)) = 0 Then
        ctrlPorcentaje = "0"
    Else
        If CCur(ctrlComision) > 0 Then
            ctrlPorcentaje = CStr((CCur(ctrlComision) / CCur(lblSubTotal(ctrlComision.Index).Caption)) * 100)
        Else
            ctrlPorcentaje = "0"
        End If
    End If
                
End Sub
'
'
'
'
'
Private Sub ActualizaComision(ctrlPorcentaje As Control, ctrlComision As Control)

    If Not IsNumeric(lblSubTotal(ctrlPorcentaje.Index).Caption) Or Not IsNumeric(ctrlPorcentaje) Then Exit Sub
        
    If CDbl(ctrlPorcentaje) > 0 Then
    
        'If AplicarTipoCambioPactado = False Then '/* */
        '    ctrlComision = CStr(CCur(lblSubTotal(ctrlPorcentaje.Index).Caption) * CDbl(ctrlPorcentaje) / 100)
        'Else
        'If Val(txtSubTotalMonedaPago.Text) = 0 Then
        
            ctrlComision = CStr(CCur(lblSubTotal(ctrlPorcentaje.Index).Caption) * CDbl(ctrlPorcentaje) / 100)
        
        'Else
        
        '    ctrlComision = CStr(CCur(txtSubTotalMonedaPago.Text) * CDbl(ctrlPorcentaje) / 100)
            
        'End If
        'End If
    Else
        ctrlComision = "0"
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


Private Sub txtComisionFondoG_Change(Index As Integer)

    Call FormatoCajaTexto(txtComisionFondoG(Index), Decimales_Monto)

    If chkAplicar(Index).Value Then
        ActualizaPorcentaje txtComisionFondoG(Index), lblPorcenFondoG(Index)
    End If
    
    Call CalculoTotal(Index)
    
End Sub

Private Sub txtComisionFondoG_KeyPress(Index As Integer, KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtComisionFondoG(Index), Decimales_Monto)
    
    If KeyAscii = vbKeyReturn Then
        Call CalculoTotal(Index)
    End If
    
End Sub


' /* 09:51 a.m. 04/09/2008 */
' /* Se agrego esta linea para el txtPorcenGastoBancario */

Private Sub txtComisionxAccion_Change(Index As Integer)

    Call FormatoCajaTexto(txtComisionxAccion(Index), Decimales_Monto)
    
    If chkAplicar(Index).Value Then
        ActualizaPorcentaje txtComisionxAccion(Index), lblPorcenComisionAccion(Index)
    End If
    
    Call CalculoTotal(Index)
    
End Sub

Private Sub txtComisionGastoBancario_Change(Index As Integer)

    Call FormatoCajaTexto(txtComisionGastoBancario(Index), Decimales_Monto)
   
    If chkAplicar(Index).Value Then
        ActualizaPorcentaje txtComisionGastoBancario(Index), lblPorcenGastoBancario(Index)
    End If
    
    Call CalculoTotal(Index)
    
End Sub

Private Sub txtComisionxAccion_KeyPress(Index As Integer, KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtComisionxAccion(Index), Decimales_Monto)
    
    If KeyAscii = vbKeyReturn Then
        Call CalculoTotal(Index)
    End If
    
End Sub

Private Sub txtComisionGastoBancario_KeyPress(Index As Integer, KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtComisionGastoBancario(Index), Decimales_Monto)
    
    If KeyAscii = vbKeyReturn Then
        Call CalculoTotal(Index)
    End If
    
End Sub


' /* until here */


Private Sub txtPorcenAgente_Change(Index As Integer)

    Call FormatoCajaTexto(txtPorcenAgente(Index), Decimales_Tasa)
    
'    If chkAplicar(Index).Value Then
'        ActualizaComision txtPorcenAgente(Index), txtComisionAgente(Index)
'    End If
    
End Sub

Private Sub txtPorcenAgente_KeyPress(Index As Integer, KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtPorcenAgente(Index), Decimales_Tasa)
    
    If KeyAscii = vbKeyReturn Then
        Call CalculoTotal(Index)
    End If
    
End Sub

Private Sub txtPrecioUnitario_Change(Index As Integer)

    Call FormatoCajaTexto(txtPrecioUnitario(Index), Decimales_Precio)
    
    txtCantidad_Change
    
    lblPrecioResumen(Index).Caption = CStr(txtPrecioUnitario(Index).Text)
    
End Sub

Private Sub txtPrecioUnitario_KeyPress(Index As Integer, KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtPrecioUnitario(Index), Decimales_Precio)
    
End Sub


Private Sub txtPrecioUnitario_LostFocus(Index As Integer)

    lblPrecioResumen(Index).Caption = CStr(txtPrecioUnitario(Index).Text)
    
End Sub

Private Sub txtTasaMensual_Change()

    Call FormatoCajaTexto(txtTasaMensual, Decimales_Tasa)
    
            
End Sub

Private Sub txtTipoCambio_Change()

    Call FormatoCajaTexto(txtTipoCambio, Decimales_TipoCambio)
    
End Sub

Private Sub txtTipoCambio_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtTipoCambio, Decimales_TipoCambio)
    
End Sub

'/**/
Private Sub txtTipoCambioConversion_Change()

    Call FormatoCajaTexto(txtTipoCambioConversion, Decimales_TipoCambio)

End Sub

Private Sub txtTipoCambioConversion_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtTipoCambioConversion, Decimales_TipoCambio)
   
End Sub
'/**/

Private Sub CalculoImpuesto(Index As Integer)

    Dim curComImp As Currency, curMonTotal As Currency

    If Not (IsNumeric(txtComisionAgente(Index).Text) And IsNumeric(txtComisionBolsa(Index).Text) And IsNumeric(txtComisionConasev(Index).Text) And IsNumeric(txtComisionCavali(Index).Text) And IsNumeric(txtComisionFondo(Index).Text) And IsNumeric(lblPorcenIgv(Index).Caption)) Then Exit Sub
    
    curComImp = CCur(CCur(txtComisionAgente(Index).Text) + CCur(txtComisionBolsa(Index).Text) + CCur(txtComisionCavali(Index).Text) + CCur(txtComisionFondo(Index).Text) + CCur(txtComisionFondoG(Index).Text)) * CDbl(lblPorcenIgv(Index).Caption)
    
    txtComisionIgv(Index).Text = CStr(curComImp)
    
End Sub

Private Sub ExportarExcel()
    
    Dim adoRegistro As ADODB.Recordset
    Dim execSQL As String
    Dim rutaExportacion As String
    
    Dim datFechaSiguiente As Date
    Dim strFechaLiquidacionHasta As String
    
    Set frmFormulario = frmOrdenRentaVariable
    
    Set adoRegistro = New ADODB.Recordset
    
    'If TodoOK() Then
        
        Dim strNameProc As String
        
        gstrNameRepo = "OrdenRentaVariable"
        
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
