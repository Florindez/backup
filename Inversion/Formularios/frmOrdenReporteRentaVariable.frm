VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmOrdenReporteRentaVariable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ordenes - Reporte con Renta Variable"
   ClientHeight    =   8730
   ClientLeft      =   1125
   ClientTop       =   1725
   ClientWidth     =   14340
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
   Icon            =   "frmOrdenReporteRentaVariable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8730
   ScaleWidth      =   14340
   Begin TAMControls2.ucBotonEdicion2 cmdOpcion 
      Height          =   735
      Left            =   480
      TabIndex        =   183
      Top             =   7800
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
      Left            =   12480
      TabIndex        =   182
      Top             =   7800
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
      Caption0        =   "&Salir"
      Tag0            =   "9"
      ToolTipText0    =   "Salir"
      UserControlWidth=   1200
   End
   Begin TabDlg.SSTab tabReporte 
      Height          =   7650
      Left            =   120
      TabIndex        =   50
      Top             =   45
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   13494
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Lista"
      TabPicture(0)   =   "frmOrdenReporteRentaVariable.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tdgConsulta"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraCriterio"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Datos Orden Inversión"
      TabPicture(1)   =   "frmOrdenReporteRentaVariable.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdAccion"
      Tab(1).Control(1)=   "txtObservacion"
      Tab(1).Control(2)=   "fraResumen"
      Tab(1).Control(3)=   "fraDatosOrden"
      Tab(1).Control(4)=   "fraDatosBasicos"
      Tab(1).Control(5)=   "lblDescrip(50)"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Negociación"
      TabPicture(2)   =   "frmOrdenReporteRentaVariable.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraPosicion"
      Tab(2).Control(1)=   "fraDatosNegociacion"
      Tab(2).Control(2)=   "fraComisionMontoFL2"
      Tab(2).Control(3)=   "fraComisionMontoFL1"
      Tab(2).ControlCount=   4
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   -64200
         TabIndex        =   184
         Top             =   6795
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
         Height          =   525
         Left            =   -72840
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   25
         Top             =   6840
         Width           =   7920
      End
      Begin VB.Frame fraPosicion 
         Caption         =   "Datos Posición Garantía"
         Height          =   1940
         Left            =   -65880
         TabIndex        =   158
         Top             =   480
         Width           =   4695
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
            Left            =   2160
            TabIndex        =   166
            Tag             =   "0.00"
            ToolTipText     =   "Moneda del Título"
            Top             =   1440
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
            Left            =   2160
            TabIndex        =   165
            Tag             =   "0.00"
            Top             =   1080
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
            Left            =   2160
            TabIndex        =   164
            Tag             =   "0.00"
            Top             =   720
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
            Height          =   285
            Left            =   2160
            TabIndex        =   163
            Tag             =   "0.00"
            Top             =   360
            Width           =   2025
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
            Index           =   71
            Left            =   480
            TabIndex        =   162
            Top             =   1455
            Width           =   585
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
            Index           =   70
            Left            =   480
            TabIndex        =   161
            Top             =   1095
            Width           =   1035
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
            Index           =   68
            Left            =   480
            TabIndex        =   160
            Top             =   740
            Width           =   885
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
            Index           =   67
            Left            =   480
            TabIndex        =   159
            Top             =   380
            Width           =   1050
         End
      End
      Begin VB.Frame fraDatosNegociacion 
         Caption         =   "Negociación"
         Height          =   1940
         Left            =   -74760
         TabIndex        =   153
         Top             =   480
         Width           =   8775
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
            Left            =   2280
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   1395
            Width           =   1575
         End
         Begin VB.TextBox txtPrecioMercado 
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
            Left            =   2280
            MaxLength       =   45
            TabIndex        =   26
            Top             =   360
            Width           =   1580
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
            Left            =   2280
            MaxLength       =   45
            TabIndex        =   28
            Top             =   1050
            Width           =   1580
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
            Left            =   2280
            MaxLength       =   45
            TabIndex        =   27
            Top             =   705
            Width           =   1580
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Plazo (Días)"
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
            Index           =   73
            Left            =   4920
            TabIndex        =   172
            Top             =   1415
            Width           =   870
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
            Index           =   72
            Left            =   4920
            TabIndex        =   171
            Top             =   897
            Width           =   1365
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
            Index           =   69
            Left            =   4920
            TabIndex        =   170
            Top             =   380
            Width           =   1305
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
            Left            =   6900
            TabIndex        =   169
            Tag             =   "0.00"
            ToolTipText     =   "Fecha de Liquidación"
            Top             =   360
            Width           =   1455
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000015&
            X1              =   4320
            X2              =   4320
            Y1              =   240
            Y2              =   1800
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
            Left            =   6900
            TabIndex        =   168
            Tag             =   "0.00"
            ToolTipText     =   "Fecha de Vencimiento del Título de la Orden"
            Top             =   877
            Width           =   1455
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
            Left            =   6900
            TabIndex        =   167
            Tag             =   "0.00"
            ToolTipText     =   "Días de Plazo del Título de la Orden"
            Top             =   1395
            Width           =   1455
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
            Left            =   360
            TabIndex        =   157
            Top             =   1070
            Width           =   885
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Precio Mercado"
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
            Left            =   360
            TabIndex        =   156
            Top             =   380
            Width           =   1125
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
            Left            =   360
            TabIndex        =   155
            Top             =   1415
            Width           =   570
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
            Index           =   16
            Left            =   360
            TabIndex        =   154
            Top             =   720
            Width           =   1095
         End
      End
      Begin VB.Frame fraResumen 
         Caption         =   "Resumen Negociación"
         Height          =   2535
         Left            =   -74760
         TabIndex        =   124
         Top             =   4200
         Width           =   13575
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
            Index           =   74
            Left            =   10080
            TabIndex        =   174
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
            Left            =   11280
            TabIndex        =   173
            Tag             =   "0.00"
            Top             =   360
            Width           =   2025
         End
         Begin VB.Line Line3 
            BorderColor     =   &H80000015&
            X1              =   9720
            X2              =   9720
            Y1              =   360
            Y2              =   2340
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
            TabIndex        =   150
            Top             =   720
            Width           =   1845
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
            TabIndex        =   149
            Top             =   720
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
            Index           =   62
            Left            =   480
            TabIndex        =   148
            Top             =   375
            Width           =   1095
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
            TabIndex        =   147
            Tag             =   "0.00"
            Top             =   360
            Width           =   2025
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000015&
            X1              =   4800
            X2              =   4800
            Y1              =   360
            Y2              =   2340
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
            Left            =   11280
            TabIndex        =   146
            Tag             =   "0.00"
            Top             =   1320
            Width           =   2025
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
            Left            =   11280
            TabIndex        =   145
            Tag             =   "0.00"
            Top             =   960
            Width           =   2025
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
            Index           =   56
            Left            =   10080
            TabIndex        =   144
            Top             =   1320
            Width           =   570
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
            Index           =   53
            Left            =   10080
            TabIndex        =   143
            Top             =   960
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
            TabIndex        =   142
            Top             =   720
            Width           =   390
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
            TabIndex        =   141
            Top             =   720
            Width           =   600
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
            TabIndex        =   140
            Tag             =   "0.00"
            Top             =   2020
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
            TabIndex        =   139
            Tag             =   "0.00"
            Top             =   1700
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
            TabIndex        =   138
            Tag             =   "0.00"
            Top             =   1360
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
            TabIndex        =   137
            Tag             =   "0.00"
            Top             =   1035
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
            TabIndex        =   136
            Tag             =   "0.00"
            Top             =   2020
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
            TabIndex        =   135
            Tag             =   "0.00"
            Top             =   1700
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
            TabIndex        =   134
            Tag             =   "0.00"
            Top             =   1360
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
            Index           =   0
            Left            =   2400
            TabIndex        =   133
            Tag             =   "0.00"
            Top             =   1035
            Width           =   2025
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
            TabIndex        =   132
            Top             =   2040
            Width           =   855
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
            TabIndex        =   131
            Top             =   1720
            Width           =   795
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
            TabIndex        =   130
            Top             =   1380
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
            Height          =   195
            Index           =   60
            Left            =   480
            TabIndex        =   129
            Top             =   2040
            Width           =   855
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
            TabIndex        =   128
            Top             =   1720
            Width           =   795
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
            TabIndex        =   127
            Top             =   1380
            Width           =   645
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
            Index           =   52
            Left            =   480
            TabIndex        =   126
            Top             =   1055
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
            Index           =   51
            Left            =   5280
            TabIndex        =   125
            Top             =   1055
            Width           =   450
         End
      End
      Begin VB.Frame fraComisionMontoFL2 
         Caption         =   "Comisiones y Montos - Plazo (FL2)"
         Height          =   4335
         Left            =   -67920
         TabIndex        =   100
         Top             =   2760
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
            TabIndex        =   180
            Top             =   2610
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
            TabIndex        =   42
            Top             =   980
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
            TabIndex        =   47
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
            Index           =   1
            Left            =   4290
            MaxLength       =   45
            TabIndex        =   46
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
            Index           =   1
            Left            =   4290
            MaxLength       =   45
            TabIndex        =   45
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
            Index           =   1
            Left            =   4290
            MaxLength       =   45
            TabIndex        =   44
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
            Index           =   1
            Left            =   4290
            MaxLength       =   45
            TabIndex        =   43
            Top             =   980
            Width           =   2025
         End
         Begin VB.TextBox txtInteresCorrido 
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
            TabIndex        =   48
            Top             =   3100
            Width           =   2025
         End
         Begin VB.CommandButton cmdCalculo 
            Caption         =   "#"
            Height          =   285
            Left            =   480
            TabIndex        =   49
            ToolTipText     =   "Calcular TIRs de la orden"
            Top             =   3885
            Width           =   375
         End
         Begin VB.CheckBox chkAplicar 
            Caption         =   "Aplicar"
            Height          =   255
            Index           =   1
            Left            =   390
            TabIndex        =   40
            Top             =   540
            Width           =   975
         End
         Begin VB.TextBox txtSubTotal 
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
            TabIndex        =   41
            Top             =   520
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
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   14
            Left            =   1680
            TabIndex        =   123
            Top             =   260
            Width           =   450
         End
         Begin VB.Label lblPrecio 
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
            TabIndex        =   122
            Top             =   240
            Width           =   1335
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
            TabIndex        =   121
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
            Index           =   18
            Left            =   390
            TabIndex        =   120
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
            Index           =   30
            Left            =   390
            TabIndex        =   119
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
            Index           =   40
            Left            =   390
            TabIndex        =   118
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
            Index           =   41
            Left            =   390
            TabIndex        =   117
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
            Index           =   42
            Left            =   390
            TabIndex        =   116
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
            Index           =   47
            Left            =   2640
            TabIndex        =   115
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
            Index           =   39
            Left            =   2640
            TabIndex        =   114
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
            Index           =   38
            Left            =   2640
            TabIndex        =   113
            Top             =   3460
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
            Index           =   1
            Left            =   4365
            TabIndex        =   112
            Top             =   240
            Width           =   1845
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
            Left            =   2280
            TabIndex        =   111
            Tag             =   "0.00"
            Top             =   3915
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
            Left            =   4905
            TabIndex        =   110
            Tag             =   "0.00"
            Top             =   3915
            Width           =   1335
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   3
            X1              =   2580
            X2              =   6300
            Y1              =   880
            Y2              =   880
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
            Left            =   630
            TabIndex        =   109
            Tag             =   "0.00"
            ToolTipText     =   "Monto de Comisión IGV"
            Top             =   3450
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
            Index           =   1
            Left            =   2625
            TabIndex        =   108
            Tag             =   "0.00"
            ToolTipText     =   "Porcentaje de Comisión IGV"
            Top             =   2620
            Width           =   1335
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   4
            X1              =   360
            X2              =   6300
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
            Index           =   1
            Left            =   4290
            TabIndex        =   107
            Tag             =   "0.00"
            ToolTipText     =   "Monto Total de la Orden"
            Top             =   3440
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
            TabIndex        =   106
            Tag             =   "0"
            ToolTipText     =   "Porcentaje de Comisión BVL"
            Top             =   1308
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
            TabIndex        =   105
            Tag             =   "0"
            ToolTipText     =   "Porcentaje de Comisión Cavali"
            Top             =   1636
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
            TabIndex        =   104
            Tag             =   "0"
            ToolTipText     =   "Porcentaje de Comisión Fondo Liquidación"
            Top             =   1964
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
            TabIndex        =   103
            Tag             =   "0"
            ToolTipText     =   "Porcentaje de Comisión Conasev"
            Top             =   2292
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
            Left            =   1320
            TabIndex        =   102
            Top             =   3900
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
            Left            =   3840
            TabIndex        =   101
            Top             =   3900
            Width           =   660
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   5
            X1              =   360
            X2              =   6300
            Y1              =   3800
            Y2              =   3800
         End
      End
      Begin VB.Frame fraComisionMontoFL1 
         Caption         =   "Comisiones y Montos - Contado (FL1)"
         Height          =   4335
         Left            =   -74760
         TabIndex        =   79
         Top             =   2775
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
            Index           =   0
            Left            =   4290
            MaxLength       =   45
            TabIndex        =   179
            Top             =   2610
            Width           =   2025
         End
         Begin VB.CheckBox chkAplicar 
            Caption         =   "Aplicar"
            Height          =   255
            Index           =   0
            Left            =   390
            TabIndex        =   30
            Top             =   540
            Width           =   975
         End
         Begin VB.TextBox txtInteresCorrido 
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
            TabIndex        =   38
            Top             =   3100
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
            Left            =   4290
            MaxLength       =   45
            TabIndex        =   33
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
            Index           =   0
            Left            =   4290
            MaxLength       =   45
            TabIndex        =   34
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
            Index           =   0
            Left            =   4290
            MaxLength       =   45
            TabIndex        =   35
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
            Index           =   0
            Left            =   4290
            MaxLength       =   45
            TabIndex        =   36
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
            Index           =   0
            Left            =   4290
            MaxLength       =   45
            TabIndex        =   37
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
            Index           =   0
            Left            =   2625
            MaxLength       =   45
            TabIndex        =   32
            Top             =   980
            Width           =   1340
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
            Left            =   4290
            MaxLength       =   45
            TabIndex        =   39
            Top             =   3900
            Width           =   2025
         End
         Begin VB.TextBox txtSubTotal 
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
            TabIndex        =   31
            Top             =   520
            Width           =   2025
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   2
            X1              =   360
            X2              =   6300
            Y1              =   3800
            Y2              =   3800
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
            TabIndex        =   99
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
            Index           =   0
            Left            =   2625
            TabIndex        =   98
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
            Index           =   0
            Left            =   2625
            TabIndex        =   97
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
            Index           =   0
            Left            =   2625
            TabIndex        =   96
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
            Index           =   0
            Left            =   4290
            TabIndex        =   95
            Tag             =   "0.00"
            ToolTipText     =   "Monto Total de la Orden"
            Top             =   3440
            Width           =   2025
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   1
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
            Index           =   0
            Left            =   2625
            TabIndex        =   94
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
            Index           =   0
            Left            =   120
            TabIndex        =   93
            Tag             =   "0.00"
            ToolTipText     =   "Monto de Comisión IGV"
            Top             =   3930
            Width           =   2025
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   0
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
            Index           =   0
            Left            =   4365
            TabIndex        =   92
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
            Index           =   36
            Left            =   2640
            TabIndex        =   91
            Top             =   3460
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
            Index           =   35
            Left            =   2640
            TabIndex        =   90
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
            Index           =   25
            Left            =   2640
            TabIndex        =   89
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
            Index           =   34
            Left            =   390
            TabIndex        =   88
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
            Index           =   33
            Left            =   390
            TabIndex        =   87
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
            Index           =   32
            Left            =   390
            TabIndex        =   86
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
            Index           =   27
            Left            =   390
            TabIndex        =   85
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
            Index           =   26
            Left            =   390
            TabIndex        =   84
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
            Index           =   24
            Left            =   390
            TabIndex        =   83
            Top             =   1000
            Width           =   990
         End
         Begin VB.Label lblPrecio 
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
            Left            =   2640
            TabIndex        =   82
            Top             =   240
            Width           =   1335
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
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   9
            Left            =   1680
            TabIndex        =   81
            Top             =   260
            Width           =   450
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
            TabIndex        =   80
            Top             =   3960
            Width           =   1260
         End
      End
      Begin VB.Frame fraDatosOrden 
         Caption         =   "Datos de la Orden"
         Height          =   1380
         Left            =   -74760
         TabIndex        =   66
         Top             =   2820
         Width           =   13575
         Begin VB.TextBox txtNemonico 
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
            Left            =   8235
            MaxLength       =   15
            TabIndex        =   24
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox txtDiasPlazo 
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
            Left            =   8235
            TabIndex        =   21
            Top             =   195
            Width           =   1275
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
            Left            =   1755
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   600
            Width           =   4840
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
            Left            =   1755
            TabIndex        =   20
            Top             =   960
            Width           =   4840
         End
         Begin MSComCtl2.DTPicker dtpFechaOrden 
            Height          =   285
            Left            =   1755
            TabIndex        =   17
            Top             =   255
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
         Begin MSComCtl2.UpDown updDiasPlazo 
            Height          =   285
            Left            =   9510
            TabIndex        =   22
            Top             =   195
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtDiasPlazo"
            BuddyDispid     =   196665
            OrigLeft        =   2920
            OrigTop         =   560
            OrigRight       =   3175
            OrigBottom      =   845
            Max             =   360
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.DTPicker dtpFechaLiquidacion 
            Height          =   285
            Left            =   5040
            TabIndex        =   18
            Top             =   255
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
         Begin MSComCtl2.DTPicker dtpFechaVencimiento 
            Height          =   285
            Left            =   11640
            TabIndex        =   23
            Top             =   195
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
         Begin MSComCtl2.DTPicker dtpFechaPago 
            Height          =   285
            Left            =   8235
            TabIndex        =   178
            Top             =   960
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
            Caption         =   "Pago"
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
            Left            =   7080
            TabIndex        =   177
            Top             =   960
            Width           =   375
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Nemómico"
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
            Left            =   7080
            TabIndex        =   175
            Top             =   600
            Width           =   750
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
            Index           =   10
            Left            =   240
            TabIndex        =   77
            Top             =   600
            Width           =   585
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
            Left            =   240
            TabIndex        =   71
            Top             =   960
            Width           =   840
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Vencimiento"
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
            Left            =   10110
            TabIndex        =   70
            Top             =   225
            Width           =   870
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Plazo (Días)"
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
            Left            =   7080
            TabIndex        =   69
            Top             =   225
            Width           =   870
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
            Index           =   4
            Left            =   3750
            TabIndex        =   68
            Top             =   285
            Width           =   810
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Orden"
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
            TabIndex        =   67
            Top             =   285
            Width           =   435
         End
      End
      Begin VB.Frame fraDatosBasicos 
         Caption         =   "Datos Básicos"
         Height          =   2340
         Left            =   -74760
         TabIndex        =   62
         Top             =   480
         Width           =   13575
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
            Left            =   9075
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   1755
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
            Left            =   9075
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   1404
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
            Left            =   9075
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   708
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
            Left            =   9075
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   1056
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
            Top             =   1755
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
            Left            =   1740
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1404
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
            Top             =   1056
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
            Left            =   9075
            Style           =   2  'Dropdown List
            TabIndex        =   12
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
            TabIndex        =   8
            Top             =   708
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
            Left            =   1755
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   360
            Width           =   4185
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
            Index           =   66
            Left            =   6720
            TabIndex        =   152
            Top             =   1770
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
            Height          =   195
            Index           =   65
            Left            =   6720
            TabIndex        =   151
            Top             =   1425
            Width           =   1530
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
            Left            =   6720
            TabIndex        =   76
            Top             =   735
            Width           =   1590
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
            Left            =   6720
            TabIndex        =   75
            Top             =   1080
            Width           =   1755
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
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   15
            Left            =   360
            TabIndex        =   74
            Top             =   1775
            Width           =   660
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
            TabIndex        =   73
            Top             =   1076
            Width           =   825
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
            TabIndex        =   72
            Top             =   1424
            Width           =   390
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Título Garantía"
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
            Left            =   6720
            TabIndex        =   65
            Top             =   375
            Width           =   1095
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
            TabIndex        =   64
            Top             =   728
            Width           =   510
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
            TabIndex        =   63
            Top             =   380
            Width           =   450
         End
      End
      Begin VB.Frame fraCriterio 
         Caption         =   "Criterios de Búsqueda"
         Height          =   2055
         Left            =   240
         TabIndex        =   51
         Top             =   480
         Width           =   13575
         Begin VB.CommandButton cmdExportarExcel 
            Caption         =   "Excel"
            Height          =   735
            Left            =   10440
            Picture         =   "frmOrdenReporteRentaVariable.frx":0060
            Style           =   1  'Graphical
            TabIndex        =   181
            Top             =   1200
            Width           =   1200
         End
         Begin VB.CommandButton cmdEnviar 
            Caption         =   "En&viar"
            Height          =   735
            Left            =   11960
            Picture         =   "frmOrdenReporteRentaVariable.frx":0668
            Style           =   1  'Graphical
            TabIndex        =   176
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
            Width           =   4905
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
            Width           =   4905
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
            Width           =   4905
         End
         Begin MSComCtl2.DTPicker dtpFechaOrdenDesde 
            Height          =   285
            Left            =   9360
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
            Left            =   11715
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
            Left            =   9360
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
            Left            =   11715
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
            Index           =   46
            Left            =   11040
            TabIndex        =   60
            Top             =   795
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
            Index           =   45
            Left            =   8640
            TabIndex        =   59
            Top             =   795
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
            Index           =   44
            Left            =   6960
            TabIndex        =   58
            Top             =   795
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
            Index           =   43
            Left            =   6960
            TabIndex        =   57
            Top             =   375
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
            TabIndex        =   56
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
            Left            =   8640
            TabIndex        =   55
            Top             =   375
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
            Index           =   21
            Left            =   11040
            TabIndex        =   54
            Top             =   375
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
            Index           =   22
            Left            =   480
            TabIndex        =   53
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
            TabIndex        =   52
            Top             =   1220
            Width           =   495
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmOrdenReporteRentaVariable.frx":0BC3
         Height          =   4575
         Left            =   240
         OleObjectBlob   =   "frmOrdenReporteRentaVariable.frx":0BDD
         TabIndex        =   61
         Top             =   2700
         Width           =   13545
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
         TabIndex        =   78
         Top             =   6855
         Width           =   1065
      End
   End
End
Attribute VB_Name = "frmOrdenReporteRentaVariable"
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
Dim strCodConcepto          As String
Dim strEstado               As String, strCodGarantia               As String

Dim strCodFile              As String, strCodAnalitica              As String
Dim strCodGrupo             As String, strCodCiiu                   As String
Dim strEstadoOrden          As String, strCodCategoria              As String
Dim strCodRiesgo            As String, strCodSubRiesgo              As String
Dim strCalcVcto             As String, strCodSector                 As String
Dim strCodTipoCostoBolsa    As String, strCodTipoCostoConasev       As String
Dim strCodTipoCostoFondo    As String, strCodTipoCavali             As String
Dim strCodTipoCostoBolsaP   As String, strCodTipoCostoConasevP      As String
Dim strCodTipoCostoFondoP   As String, strCodTipoCavaliP            As String
Dim strIndCuponCero         As String, strIndPacto                  As String
Dim strIndNegociable        As String, strSQL                       As String
Dim strCodigosFile          As String
Dim dblTipoCambio           As Double
Dim dblComisionBolsa        As Double, dblComisionConasev           As Double
Dim dblComisionFondo        As Double, dblComisionCavali            As Double
Dim dblComisionBolsaP       As Double, dblComisionConasevP          As Double
Dim dblComisionFondoP       As Double, dblComisionCavaliP           As Double
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
        
    Dim adoRegistro         As ADODB.Recordset
    Dim strFechaDesde       As String, strFechaHasta        As String
    
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
    
    If CVDate(dtpFechaOrden.Value) > CVDate(dtpFechaVencimiento.Value) Then
        MsgBox "La Fecha de Vencimiento debe ser mayor a la Fecha de la Orden.", vbCritical, Me.Caption
        If dtpFechaVencimiento.Enabled Then dtpFechaVencimiento.SetFocus
        Exit Function
    End If
    
    If CVDate(dtpFechaLiquidacion.Value) > CVDate(dtpFechaVencimiento.Value) Then
        MsgBox "La Fecha de Vencimiento debe ser mayor a la Fecha de la Liquidacion.", vbCritical, Me.Caption
        If dtpFechaVencimiento.Enabled Then dtpFechaVencimiento.SetFocus
        Exit Function
    End If
    
    If CInt(txtDiasPlazo.Text) = 0 Then
        MsgBox "Debe indicar el número de días de plazo.", vbCritical, Me.Caption
        If txtDiasPlazo.Enabled Then txtDiasPlazo.SetFocus
        Exit Function
    End If
    
    If Trim(txtDescripOrden.Text) = Valor_Caracter Then
        MsgBox "Debe indicar la Descripción de la ORDEN.", vbCritical, Me.Caption
        If txtDescripOrden.Enabled Then txtDescripOrden.SetFocus
        Exit Function
    End If
    
    Set adoRegistro = New ADODB.Recordset
        
    '*** Buscar en Títulos ***
    adoComm.CommandText = "SELECT Nemotecnico FROM InstrumentoInversion " & _
        "WHERE CodFile='" & strCodFile & "' AND Nemotecnico='" & Trim(txtNemonico.Text) & "' AND IndVigente='X'"
    Set adoRegistro = adoComm.Execute
    
    If Not adoRegistro.EOF Then
        MsgBox "Nemónico YA EXISTE...por favor verificar.", vbCritical, Me.Caption
        If txtNemonico.Enabled Then txtNemonico.SetFocus
        adoRegistro.Close: Set adoRegistro = Nothing
        Exit Function
    End If
    adoRegistro.Close
    
    strFechaDesde = Convertyyyymmdd(dtpFechaOrden.Value)
    strFechaHasta = Convertyyyymmdd(DateAdd("d", 1, dtpFechaOrden.Value))
    
    '*** Buscar en Ordenes del día ***
    adoComm.CommandText = "SELECT Nemotecnico FROM InversionOrden " & _
        "WHERE (FechaOrden>='" & strFechaDesde & "' AND FechaOrden<'" & strFechaHasta & "') AND " & _
        "CodFile='" & strCodFile & "' AND Nemotecnico='" & Trim(txtNemonico.Text) & "' AND EstadoOrden<>'" & Estado_Orden_Anulada & "'"
    Set adoRegistro = adoComm.Execute
    
    If Not adoRegistro.EOF Then
        MsgBox "Nemónico YA EXISTE...por favor verificar.", vbCritical, Me.Caption
        If txtNemonico.Enabled Then txtNemonico.SetFocus
        adoRegistro.Close: Set adoRegistro = Nothing
        Exit Function
    End If
    adoRegistro.Close: Set adoRegistro = Nothing
        
    If CDbl(txtPrecioMercado.Text) = 0 Then
        MsgBox "Debe indicar el Precio de Mercado.", vbCritical, Me.Caption
        If txtPrecioMercado.Enabled Then txtPrecioMercado.SetFocus
        Exit Function
    End If
    
    If CDbl(txtTipoCambio.Text) = 0 Then
        MsgBox "Debe indicar el Tipo de Cambio.", vbCritical, Me.Caption
        If txtTipoCambio.Enabled Then txtTipoCambio.SetFocus
        Exit Function
    End If
    
    If CCur(txtCantidad.Text) = 0 Then
        MsgBox "Debe indicar la Cantidad.", vbCritical, Me.Caption
        If txtCantidad.Enabled Then txtCantidad.SetFocus
        Exit Function
    End If
    
    If CDbl(txtTasaMensual.Text) = 0 Then
        MsgBox "Debe indicar la Tasa Mensual", vbCritical, Me.Caption
        If txtTasaMensual.Enabled Then txtTasaMensual.SetFocus
        Exit Function
    End If
                                    
    If CCur(lblMontoTotal(0).Caption) = 0 Then
        MsgBox "El Monto al contado es Cero.", vbCritical, Me.Caption
        If txtSubTotal(0).Enabled Then txtSubTotal(0).SetFocus
        Exit Function
    End If
    
    If CCur(lblMontoTotal(1).Caption) = 0 Then
        MsgBox "El Monto a plazo es Cero.", vbCritical, Me.Caption
        If txtSubTotal(1).Enabled Then txtSubTotal(1).SetFocus
        Exit Function
    End If
    
    If CDbl(lblTirBruta.Caption) = 0 Then
        MsgBox "Debe calcular la Tir Bruta.", vbCritical, Me.Caption
        If cmdCalculo.Enabled Then cmdCalculo.SetFocus
        Exit Function
    End If

    If CDbl(lblTirNeta.Caption) = 0 Then
        MsgBox "Debe calcular la Tir Neta.", vbCritical, Me.Caption
        If cmdCalculo.Enabled Then cmdCalculo.SetFocus
        Exit Function
    End If
    
    '*** Si todo paso OK ***
    TodoOK = True
  
End Function

Private Sub ActualizaPorcentaje(ctrlComision As Control, ctrlPorcentaje As Control)

    If Not IsNumeric(ctrlComision) Or Not IsNumeric(txtSubTotal(ctrlComision.Index).Text) Then Exit Sub
                
    If CCur(txtSubTotal(ctrlComision.Index)) = 0 Then
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

    If Not EsDiaUtil(gdatFechaActual) Then
        MsgBox "No se puede negociar en un día no útil !", vbCritical, Me.Caption
        Exit Sub
    End If
    
    If cboTipoInstrumento.ListCount > 1 Then
        frmMainMdi.stbMdi.Panels(3).Text = "Adicionar Título Valor..."
                    
        strEstado = Reg_Adicion
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabReporte
            .TabEnabled(0) = False
            .Tab = 1
        End With
    Else
        MsgBox "Acceso a Negociación Denegada", vbCritical, Me.Caption
    End If
    
End Sub

Private Sub AplicarCostos(Index As Integer)
    
    If Index = 0 Then
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
        
        If strCodTipoCavali = Codigo_Tipo_Costo_Monto Then
            txtComisionCavali(Index).Text = CStr(dblComisionCavali)
        Else
            AsignaComision strCodTipoCavali, dblComisionCavali, txtComisionCavali(Index)
        End If
    Else
        If strCodTipoCostoBolsa = Codigo_Tipo_Costo_Monto Then
            txtComisionBolsa(Index).Text = CStr(dblComisionBolsaP)
        Else
            AsignaComision strCodTipoCostoBolsaP, dblComisionBolsaP, txtComisionBolsa(Index)
        End If
        
        If strCodTipoCostoConasev = Codigo_Tipo_Costo_Monto Then
            txtComisionConasev(Index).Text = CStr(dblComisionConasevP)
        Else
            AsignaComision strCodTipoCostoConasevP, dblComisionConasevP, txtComisionConasev(Index)
        End If
        
        If strCodTipoCostoFondo = Codigo_Tipo_Costo_Monto Then
            txtComisionFondo(Index).Text = CStr(dblComisionFondoP)
        Else
            AsignaComision strCodTipoCostoFondoP, dblComisionFondoP, txtComisionFondo(Index)
        End If
        
        If strCodTipoCavali = Codigo_Tipo_Costo_Monto Then
            txtComisionCavali(Index).Text = CStr(dblComisionCavaliP)
        Else
            AsignaComision strCodTipoCavaliP, dblComisionCavaliP, txtComisionCavali(Index)
        End If
    End If
    
    Call CalculoImpuesto(Index)
    
    Call CalculoTotal(Index)
    
End Sub

Private Sub AsignaComision(strTipoComision As String, dblValorComision As Double, ctrlValorComision As Control)
    
    If Not IsNumeric(txtSubTotal(ctrlValorComision.Index).Text) Then Exit Sub
    
    If dblValorComision > 0 Then
        ctrlValorComision.Text = CStr(Round(CCur(txtSubTotal(ctrlValorComision.Index)) * dblValorComision / 100, 2))
    End If
            
End Sub

Public Sub Buscar()

    Set adoConsulta = New ADODB.Recordset

    Dim strFechaOrdenDesde          As String, strFechaOrdenHasta        As String
    Dim strFechaLiquidacionDesde    As String, strFechaLiquidacionHasta  As String
    Dim datFechaSiguiente           As Date

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
        "(RTRIM(DescripTipoOperacion) + SPACE(1) + DescripOrden) DescripOrden,CantOrden,ValorNominal,PrecioUnitarioMFL1,MontoTotalMFL1, CodSigno DescripMoneda " & _
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

    If Not IsNumeric(txtComisionAgente(Index).Text) And Not IsNumeric(txtComisionBolsa(Index).Text) And Not IsNumeric(txtComisionConasev(Index).Text) And Not IsNumeric(txtComisionCavali(Index).Text) And Not IsNumeric(txtComisionFondo(Index).Text) And Not IsNumeric(txtComisionIgv(Index).Text) Then Exit Sub
    
    curComImp = CCur(txtComisionAgente(Index).Text) + CCur(txtComisionBolsa(Index).Text) + CCur(txtComisionConasev(Index).Text) + CCur(txtComisionCavali(Index).Text) + CCur(txtComisionFondo(Index).Text) + CCur(txtComisionIgv(Index).Text)

    lblComisionesResumen(Index).Caption = CStr(curComImp)
    
    If strCodTipoOrden = Codigo_Orden_Compra Then  '*** Compra ***
        If Index = 0 Then
            curMonTotal = CCur(txtSubTotal(Index).Text) + curComImp
        Else
            curMonTotal = CCur(txtSubTotal(Index).Text) - curComImp
        End If
    ElseIf strCodTipoOrden = Codigo_Orden_Venta Then '*** Venta ***
        curMonTotal = CCur(txtSubTotal(Index).Text) - curComImp
    End If
        
    curMonTotal = curMonTotal + CCur(txtInteresCorrido(Index).Text)
    lblMontoTotal(Index).Caption = CStr(curMonTotal)
    
End Sub
Private Sub CalculoImpuesto(Index As Integer)

    Dim curComImp As Currency, curMonTotal As Currency

    If Not IsNumeric(txtComisionAgente(Index).Text) And Not IsNumeric(txtComisionBolsa(Index).Text) And Not IsNumeric(txtComisionConasev(Index).Text) And Not IsNumeric(txtComisionCavali(Index).Text) And Not IsNumeric(txtComisionFondo(Index).Text) Then Exit Sub
    
    curComImp = CCur(CCur(txtComisionAgente(Index).Text) + CCur(txtComisionBolsa(Index).Text) + CCur(txtComisionConasev(Index).Text) + CCur(txtComisionCavali(Index).Text) + CCur(txtComisionFondo(Index).Text)) * CDbl(lblPorcenIgv(Index).Caption)
    txtComisionIgv(Index).Text = CStr(curComImp)
    
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

    If cboOperacion.ListCount > 0 Then cboOperacion.ListIndex = 0
    
    '*** Mecanismos de Negociación ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='MECNEG' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboNegociacion, arrNegociacion(), Valor_Caracter

    If cboNegociacion.ListCount > 0 Then cboNegociacion.ListIndex = 0
    
    '*** Conceptos de Costos de Negociación ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPCCO' AND ValorParametro='RV' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboConceptoCosto, arrConceptoCosto(), Sel_Defecto

    '*** Agente ***
    strSQL = "SELECT (CodPersona + CodGrupo + CodCiiu) CODIGO,DescripPersona DESCRIP FROM InstitucionPersona WHERE TipoPersona='" & Codigo_Tipo_Persona_Agente & "' AND IndVigente='X' ORDER BY DescripPersona"
    CargarControlLista strSQL, cboAgente, arrAgente(), Sel_Defecto

    '*** Mercado de Negociación ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='MDONEG' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboOrigen, arrOrigen(), Valor_Caracter
            
    '*** Moneda ***
    strSQL = "{ call up_ACSelDatos(2) }"
    CargarControlLista strSQL, cboMoneda, arrMoneda(), Sel_Defecto
    
    '*** Tipo de Cálculo ***
    cboCalculo.AddItem Sel_Defecto, 0
    cboCalculo.AddItem "Cantidad", 1
    cboCalculo.AddItem "SubTotal", 2
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
    Dim strFechaPago        As String
    Dim strMensaje          As String, strIndTitulo             As String
    Dim intRegistro         As Integer, intAccion               As Integer
    Dim lngNumError         As Long
    
    On Error GoTo CtrlError
    
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
                "Fecha de Liquidación" & Space(3) & ">" & Space(2) & CStr(dtpFechaLiquidacion.Value) & Chr(vbKeyReturn) & Chr(vbKeyReturn) & _
                "Fecha de Vencimiento" & Space(1) & ">" & Space(2) & CStr(dtpFechaVencimiento.Value) & Chr(vbKeyReturn) & Chr(vbKeyReturn) & _
                "Fecha de Pago" & Space(12) & ">" & Space(2) & CStr(dtpFechaPago.Value) & Chr(vbKeyReturn) & Chr(vbKeyReturn) & _
                "Cantidad" & Space(22) & ">" & Space(2) & txtCantidad.Text & Chr(vbKeyReturn) & _
                "Precio Unitario (%)" & Space(6) & ">" & Space(2) & lblPrecio(0).Caption & Chr(vbKeyReturn) & Chr(vbKeyReturn) & _
                "Monto Total" & Space(17) & ">" & Space(2) & Trim(lblDescripMoneda(0).Caption) & Space(1) & lblMontoTotal(0).Caption & Chr(vbKeyReturn) & Chr(vbKeyReturn) & _
                "Tir Neta" & Space(23) & ">" & Space(2) & lblTirNeta.Caption & Chr(vbKeyReturn) & _
                Chr(vbKeyReturn) & Chr(vbKeyReturn) & _
                "¿ Seguro de continuar ?"

            If MsgBox(strMensaje, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
               Me.Refresh: Exit Sub
            End If

        
            Me.MousePointer = vbHourglass
            
            strFechaOrden = Convertyyyymmdd(dtpFechaOrden.Value)
            strFechaEmision = Convertyyyymmdd(dtpFechaOrden.Value)
            strFechaLiquidacion = Convertyyyymmdd(dtpFechaLiquidacion.Value)
            strFechaVencimiento = Convertyyyymmdd(dtpFechaVencimiento.Value)
            strFechaPago = Convertyyyymmdd(dtpFechaPago.Value)
            
            Set adoRegistro = New ADODB.Recordset
            '*** Guardar Orden de Inversion ***
            With adoComm
                strIndTitulo = Valor_Caracter
                strCodAnalitica = NumAleatorio(8)
                strCodTitulo = NumAleatorio(15)
                strCodTipoTasa = Codigo_Tipo_Tasa_Efectiva
                strCodBaseAnual = Codigo_Base_Actual_365
                strCodRiesgo = "00" ' Sin Clasificacion
                
'                .CommandText = "BEGIN TRAN ProcOrden"
'                adoConn.Execute .CommandText
                                
                .CommandText = "{ call up_IVAdicInversionOrden('" & strCodFondoOrden & "','" & _
                    gstrCodAdministradora & "','','" & strFechaOrden & "','" & _
                    strCodTitulo & "','" & Trim(txtNemonico.Text) & "','" & gstrPeriodoActual & "','" & gstrMesActual & "','" & _
                    "','" & strEstadoOrden & "','" & strCodAnalitica & "','" & strCodFile & "','" & _
                    strCodAnalitica & "','" & strCodClaseInstrumento & "','" & strCodSubClaseInstrumento & "','" & strCodTipoOrden & "','" & _
                    strCodOperacion & "','" & strCodNegociacion & "','" & strCodOrigen & "','" & Trim(txtDescripOrden.Text) & "','" & strCodEmisor & "','" & _
                    strCodAgente & "','" & strCodGarantia & "','" & strFechaPago & "','" & strFechaVencimiento & "','" & strFechaLiquidacion & "','" & _
                    strFechaEmision & "','" & strCodMoneda & "','" & strCodMoneda & "','" & strCodMoneda & "'," & CDec(txtCantidad.Text) & "," & CDec(txtTipoCambio.Text) & ",0," & _
                    "1,0,0," & CDec(lblPrecio(0).Caption) & "," & CDec(txtSubTotal(0).Text) & "," & CDec(txtSubTotal(0).Text) & "," & _
                    CDec(txtInteresCorrido(0).Text) & "," & CDec(txtComisionAgente(0).Text) & "," & CDec(txtComisionCavali(0).Text) & "," & _
                    CDec(txtComisionConasev(0).Text) & "," & CDec(txtComisionBolsa(0).Text) & "," & CDec(txtComisionFondo(0).Text) & ",0,0,0," & _
                    CDec(txtComisionIgv(0).Text) & "," & CDec(lblMontoTotal(0).Caption) & "," & CDec(lblMontoTotal(0).Caption) & "," & CDec(lblPrecio(1).Caption) & "," & _
                    CDec(txtSubTotal(1).Text) & "," & CDec(txtInteresCorrido(1).Text) & "," & CDec(txtComisionAgente(1).Text) & "," & _
                    CDec(txtComisionCavali(1).Text) & "," & CDec(txtComisionConasev(1).Text) & "," & CDec(txtComisionBolsa(1).Text) & "," & _
                    CDec(txtComisionFondo(1).Text) & ",0,0,0," & CDec(txtComisionIgv(1).Text) & "," & CDec(lblMontoTotal(1).Caption) & "," & _
                    CDec(lblMontoTotal(1).Caption) & "," & CInt(txtDiasPlazo.Text) & ",'','','','','','" & strCodAgente & "','','','','','',0,'','','" & strIndTitulo & "','" & _
                    strCodTipoTasa & "','" & strCodBaseAnual & "'," & CDec(txtTasaMensual.Text) & "," & CDec(lblTirBruta.Caption) & "," & CDec(lblTirNeta.Caption) & ",'" & _
                    strCodRiesgo & "','" & strCodSubRiesgo & "','" & Trim(txtObservacion.Text) & "','" & gstrLogin & "') }"
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
    
    Set adoRegistro = New ADODB.Recordset
    With adoComm
        .CommandText = "SELECT CodFile FROM InversionFile  " & _
            "WHERE TipoValor='" & Codigo_Valor_RentaFija & "' AND TipoPlazo='" & Valor_Caracter & "' AND CodEstructura='03' AND IndInstrumento='X' AND IndVigente='X' " & _
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
        txtComisionIgv(intContador).Text = "0"
        
        txtPorcenAgente(intContador).Text = "0"
        lblPorcenBolsa(intContador).Caption = "0"
        lblPorcenCavali(intContador).Caption = "0"
        lblPorcenFondo(intContador).Caption = "0"
        lblPorcenConasev(intContador).Caption = "0"
        
        lblPrecioResumen(intContador).Caption = "0"
        lblSubTotalResumen(intContador).Caption = "0"
        lblComisionesResumen(intContador).Caption = "0"
        lblTotalResumen(intContador).Caption = "0"
    Next
    
End Sub

Private Sub LlenarFormulario(strModo As String)

    Dim adoRecord   As ADODB.Recordset
    Dim strSQL      As String
    Dim intRegistro As Integer
    
    Select Case strModo
        Case Reg_Adicion
            intRegistro = ObtenerItemLista(arrFondoOrden(), strCodFondo)
            If intRegistro >= 0 Then cboFondoOrden.ListIndex = intRegistro
        
            cboAgente.ListIndex = -1
            If cboAgente.ListCount > 0 Then cboAgente.ListIndex = 0
            
            cboTipoInstrumentoOrden.ListIndex = -1
            If cboTipoInstrumentoOrden.ListCount > 0 Then cboTipoInstrumentoOrden.ListIndex = 0
                                    
            cboTipoOrden.ListIndex = -1
            If cboTipoOrden.ListCount > 0 Then cboTipoOrden.ListIndex = 0
    
            cboOperacion.ListIndex = -1
            If cboOperacion.ListCount > 0 Then cboOperacion.ListIndex = 0
        
            cboNegociacion.ListIndex = -1
            If cboNegociacion.ListCount > 0 Then cboNegociacion.ListIndex = 0
            
            intRegistro = ObtenerItemLista(arrOrigen(), Codigo_Negociacion_Local)
            If intRegistro >= 0 Then cboOrigen.ListIndex = intRegistro
                                    
            dtpFechaOrden.Value = gdatFechaActual
            dtpFechaLiquidacion.Value = dtpFechaOrden.Value
            lblFechaLiquidacion.Caption = CStr(dtpFechaOrden.Value)
            
            intRegistro = ObtenerItemLista(arrMoneda(), strCodMoneda)
            If intRegistro >= 0 Then cboMoneda.ListIndex = intRegistro
            
            txtDescripOrden.Text = Valor_Caracter
            txtNemonico.Text = Valor_Caracter
            txtObservacion.Text = Valor_Caracter
            txtPrecioMercado.Text = "0"
            
            txtCantidad.Text = "1"

            'lblAnalitica.Caption = "??? - ????????"
            'lblSaldo.Caption = "0"
            'lblRiesgo.Caption = Valor_Caracter
            
            dtpFechaVencimiento.Value = dtpFechaLiquidacion.Value
            dtpFechaPago.Value = dtpFechaVencimiento.Value
            lblPrecio(0).Caption = "0"
            lblPrecio(1).Caption = "0"
            txtSubTotal(0).Text = "0"
            txtSubTotal(1).Text = "0"
            
            cboCalculo.ListIndex = -1
            If cboCalculo.ListCount > 0 Then cboCalculo.ListIndex = 0
            
            txtDiasPlazo.Text = "0"
            txtTasaMensual.Text = "0"
                                                
            chkAplicar(0).Value = vbUnchecked
            chkAplicar(1).Value = vbUnchecked
                        
            Call IniciarComisiones
            
            txtInteresCorrido(0).Text = "0"
            txtInteresCorrido(1).Text = "0"
            lblMontoTotal(0).Caption = "0"
            lblMontoTotal(1).Caption = "0"
            lblTirBruta.Caption = "0"
            lblTirNeta.Caption = "0"
            
            lblFechaCupon.Caption = Valor_Caracter
            lblClasificacion.Caption = Valor_Caracter
            lblStockNominal.Caption = "0"
            lblMoneda.Caption = Valor_Caracter
            lblCantidadResumen.Caption = "0"
                                                
            lblTirBrutaResumen.Caption = "0"
            lblTirNetaResumen.Caption = "0"
            
            cboFondoOrden.SetFocus
                        
        Case Reg_Edicion
    
    End Select
    
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


Private Sub cboCalculo_Click()

    If cboCalculo.ListIndex <= 0 Then Exit Sub
    
    If cboTitulo.ListIndex <= 0 Then
        MsgBox "Por favor seleccione la Garantía", vbCritical, Me.Caption
        cboTitulo.SetFocus
        Exit Sub
    End If
            
    Select Case UCase(Trim(cboCalculo.Text))
        Case "PRECIO"
            If CLng(txtCantidad.Text) <= 0 Then
                MsgBox "Por Favor ingrese la cantidad de acciones del reporte", vbExclamation
                cboCalculo.ListIndex = 0
                txtCantidad.SetFocus
                Exit Sub
            End If
            If CCur(txtSubTotal(0).Text) <= 0 Then
                MsgBox "Por Favor ingrese el subtotal al contado del reporte", vbExclamation
                cboCalculo.ListIndex = 0
                txtSubTotal(0).SetFocus
                Exit Sub
            End If

            '*** Calculando Precio de Mercado ***
            If strCodMoneda = strCodMonedaGarantia Then
                txtPrecioMercado.Text = CStr(CCur(txtSubTotal(0).Text) / CCur(txtCantidad.Text))
            ElseIf strCodMoneda = Codigo_Moneda_Local And strCodMonedaGarantia <> Codigo_Moneda_Local Then
                    txtPrecioMercado.Text = CStr(CCur(txtSubTotal(0).Text) / CCur(txtCantidad.Text) / CDbl(txtTipoCambio.Text))
                ElseIf strCodMoneda <> Codigo_Moneda_Local And strCodMonedaGarantia = Codigo_Moneda_Local Then
                        txtPrecioMercado.Text = CStr(CCur(txtSubTotal(0).Text) / CCur(txtCantidad.Text) * CDbl(txtTipoCambio.Text))
                    End If

        Case "CANTIDAD"
            If CCur(txtSubTotal(0).Text) <= 0 Then
                MsgBox "Por Favor ingrese el subtotal al contado del reporte", vbExclamation
                cboCalculo.ListIndex = 0
                txtSubTotal(0).SetFocus
                Exit Sub
            End If
            If CDbl(txtPrecioMercado.Text) <= 0 Then
                MsgBox "Por Favor ingrese el precio de mercado del reporte", vbExclamation
                cboCalculo.ListIndex = 0
                txtPrecioMercado.SetFocus
                Exit Sub
            End If

            '*** Calculando Cantidad de Acciones ***
            If strCodMoneda = strCodMonedaGarantia Then
                txtCantidad.Text = CStr(CCur(txtSubTotal(0).Text) / CDbl(txtPrecioMercado.Text))
            ElseIf strCodMoneda = Codigo_Moneda_Local And strCodMonedaGarantia <> Codigo_Moneda_Local Then
                    txtCantidad.Text = CStr(CCur(txtSubTotal(0).Text) / (CDbl(txtPrecioMercado.Text) * CDbl(txtTipoCambio.Text)))
                ElseIf strCodMoneda <> Codigo_Moneda_Local And strCodMonedaGarantia = Codigo_Moneda_Local Then
                        txtCantidad.Text = CStr(CCur(txtSubTotal(0).Text) / (CDbl(txtPrecioMercado.Text) / CDbl(txtTipoCambio.Text)))
                    End If
            
            txtSubTotal(0).Tag = CCur(txtSubTotal(0).Text)
            lblPrecio(0).Tag = CCur(txtSubTotal(0).Text) / CCur(txtCantidad.Text)
            lblPrecio(0).Caption = CStr(lblPrecio(0).Tag)

        Case "SUBTOTAL"
            If CCur(txtCantidad.Text) <= 0 Then
                MsgBox "Por Favor ingrese la cantidad de acciones del reporte", vbExclamation
                cboCalculo.ListIndex = 0
                txtCantidad.SetFocus
                Exit Sub
            End If
            If CDbl(txtPrecioMercado.Text) <= 0 Then
                MsgBox "Por Favor ingrese el precio de mercado del reporte", vbExclamation
                cboCalculo.ListIndex = 0
                txtPrecioMercado.SetFocus
                Exit Sub
            End If

            '*** Calculando Sub Total ***
            If strCodMoneda = strCodMonedaGarantia Then
                txtSubTotal(0).Tag = CDbl(txtPrecioMercado.Text) * CCur(txtCantidad.Text)
            ElseIf strCodMoneda = Codigo_Moneda_Local And strCodMonedaGarantia <> Codigo_Moneda_Local Then
                    txtSubTotal(0).Tag = CDbl(txtPrecioMercado.Text) * CDbl(txtTipoCambio.Text) * CCur(txtCantidad.Text)
                ElseIf strCodMoneda <> Codigo_Moneda_Local And strCodMonedaGarantia = Codigo_Moneda_Local Then
                        txtSubTotal(0).Tag = (CDbl(txtPrecioMercado.Text) / CDbl(txtTipoCambio.Text)) * CCur(txtCantidad.Text)
                    End If
            txtSubTotal(0).Text = CStr(txtSubTotal(0).Tag)

    End Select
    
End Sub

Private Sub cboClaseInstrumento_Click()

    strCodClaseInstrumento = Valor_Caracter
    If cboClaseInstrumento.ListIndex < 0 Then Exit Sub
    
    strCodClaseInstrumento = Trim(arrClaseInstrumento(cboClaseInstrumento.ListIndex))
            
    'If strCodClaseInstrumento = "001" Then strCodFile = "004"
            
    Call cboTipoOrden_Click
            
End Sub

Private Sub cboConceptoCosto_Click()

    Dim adoRegistro As ADODB.Recordset
    
    strCodConcepto = Valor_Caracter
    If cboConceptoCosto.ListIndex < 0 Then Exit Sub
    
    strCodConcepto = Trim(arrConceptoCosto(cboConceptoCosto.ListIndex))
    
    strCodTipoCostoBolsa = Valor_Caracter: strCodTipoCostoConasev = Valor_Caracter
    strCodTipoCavali = Valor_Caracter: strCodTipoCostoFondo = Valor_Caracter
    strCodTipoCostoBolsaP = Valor_Caracter: strCodTipoCostoConasevP = Valor_Caracter
    strCodTipoCavaliP = Valor_Caracter: strCodTipoCostoFondoP = Valor_Caracter
    dblComisionBolsa = 0: dblComisionConasev = 0
    dblComisionCavali = 0: dblComisionFondo = 0
    dblComisionBolsaP = 0: dblComisionConasevP = 0
    dblComisionCavaliP = 0: dblComisionFondoP = 0
        
    With adoComm
        Set adoRegistro = New ADODB.Recordset
                
        .CommandText = "SELECT CodCosto,TipoCosto,ValorCosto,TipoPlazo,SignoRestriccion,CantDias,ValorAlterno FROM CostoNegociacion WHERE TipoOperacion='" & strCodConcepto & "' AND TipoValor='" & Codigo_Valor_RentaVariable & "' ORDER BY CodCosto"
        Set adoRegistro = .Execute

        Do Until adoRegistro.EOF
            Select Case Trim(adoRegistro("CodCosto"))
                Case Codigo_Costo_Bolsa
                    If adoRegistro("TipoPlazo") = Codigo_Operacion_Contado Then
                        strCodTipoCostoBolsa = Trim(adoRegistro("TipoCosto"))
                        dblComisionBolsa = CDbl(adoRegistro("ValorCosto"))
                        
                        If CInt(txtDiasPlazo.Text) > 0 Then
                            Select Case adoRegistro("SignoRestriccion")
                                Case Codigo_Signo_Igual
                                    If CInt(txtDiasPlazo.Text) = adoRegistro("CantDias") Then
                                        dblComisionBolsa = CDbl(adoRegistro("ValorCosto"))
                                    Else
                                        dblComisionBolsa = (CDbl(adoRegistro("ValorCosto")) / CDbl(adoRegistro("CantDias"))) * CDbl(txtDiasPlazo.Text)
                                    End If
                                Case Codigo_Signo_Mayor
                                    If CInt(txtDiasPlazo.Text) > adoRegistro("CantDias") Then
                                        dblComisionBolsa = CDbl(adoRegistro("ValorCosto"))
                                    Else
                                        dblComisionBolsa = (CDbl(adoRegistro("ValorCosto")) / CDbl(adoRegistro("CantDias"))) * CDbl(txtDiasPlazo.Text)
                                    End If
                                Case Codigo_Signo_MayorIgual
                                    If CInt(txtDiasPlazo.Text) >= adoRegistro("CantDias") Then
                                        dblComisionBolsa = CDbl(adoRegistro("ValorCosto"))
                                    Else
                                        dblComisionBolsa = (CDbl(adoRegistro("ValorCosto")) / CDbl(adoRegistro("CantDias"))) * CDbl(txtDiasPlazo.Text)
                                    End If
                                Case Codigo_Signo_Menor
                                    If CInt(txtDiasPlazo.Text) < adoRegistro("CantDias") Then
                                        dblComisionBolsa = CDbl(adoRegistro("ValorCosto"))
                                    Else
                                        
                                        dblComisionBolsa = (CDbl(adoRegistro("ValorCosto")) / CDbl(adoRegistro("CantDias"))) * CDbl(txtDiasPlazo.Text)
                                    End If
                                Case Codigo_Signo_MenorIgual
                                    If CInt(txtDiasPlazo.Text) <= adoRegistro("CantDias") Then
                                        dblComisionBolsa = CDbl(adoRegistro("ValorCosto"))
                                    Else
                                        dblComisionBolsa = (CDbl(adoRegistro("ValorCosto")) / CDbl(adoRegistro("CantDias"))) * CDbl(txtDiasPlazo.Text)
                                    End If
                            End Select
                        End If
                    Else
                        strCodTipoCostoBolsaP = Trim(adoRegistro("TipoCosto"))
                        dblComisionBolsaP = CDbl(adoRegistro("ValorCosto"))
                        
                        If CInt(txtDiasPlazo.Text) > 0 Then
                            Select Case adoRegistro("SignoRestriccion")
                                Case Codigo_Signo_Igual
                                    If CInt(txtDiasPlazo.Text) = adoRegistro("CantDias") Then
                                        dblComisionBolsaP = CDbl(adoRegistro("ValorCosto"))
                                    Else
                                        dblComisionBolsaP = (CDbl(adoRegistro("ValorCosto")) / CDbl(adoRegistro("CantDias"))) * CDbl(txtDiasPlazo.Text)
                                    End If
                                Case Codigo_Signo_Mayor
                                    If CInt(txtDiasPlazo.Text) > adoRegistro("CantDias") Then
                                        dblComisionBolsaP = CDbl(adoRegistro("ValorCosto"))
                                    Else
                                        dblComisionBolsaP = (CDbl(adoRegistro("ValorCosto")) / CDbl(adoRegistro("CantDias"))) * CDbl(txtDiasPlazo.Text)
                                    End If
                                Case Codigo_Signo_MayorIgual
                                    If CInt(txtDiasPlazo.Text) >= adoRegistro("CantDias") Then
                                        dblComisionBolsaP = CDbl(adoRegistro("ValorCosto"))
                                    Else
                                        dblComisionBolsaP = (CDbl(adoRegistro("ValorCosto")) / CDbl(adoRegistro("CantDias"))) * CDbl(txtDiasPlazo.Text)
                                    End If
                                Case Codigo_Signo_Menor
                                    If CInt(txtDiasPlazo.Text) < adoRegistro("CantDias") Then
                                        dblComisionBolsaP = CDbl(adoRegistro("ValorCosto"))
                                    Else
                                        dblComisionBolsaP = (CDbl(adoRegistro("ValorCosto")) / CDbl(adoRegistro("CantDias"))) * CDbl(txtDiasPlazo.Text)
                                    End If
                                Case Codigo_Signo_MenorIgual
                                    If CInt(txtDiasPlazo.Text) <= adoRegistro("CantDias") Then
                                        dblComisionBolsaP = CDbl(adoRegistro("ValorCosto"))
                                    Else
                                        dblComisionBolsaP = (CDbl(adoRegistro("ValorCosto")) / CDbl(adoRegistro("CantDias"))) * CDbl(txtDiasPlazo.Text)
                                    End If
                            End Select
                        End If
                    End If
                    
                Case Codigo_Costo_Conasev
                    If adoRegistro("TipoPlazo") = Codigo_Operacion_Contado Then
                        strCodTipoCostoConasev = Trim(adoRegistro("TipoCosto"))
                        dblComisionConasev = CDbl(adoRegistro("ValorCosto"))
                        
                        If CInt(txtDiasPlazo.Text) > 0 Then
                            Select Case adoRegistro("SignoRestriccion")
                                Case Codigo_Signo_Igual
                                    If CInt(txtDiasPlazo.Text) = adoRegistro("CantDias") Then
                                        dblComisionConasev = CDbl(adoRegistro("ValorCosto"))
                                    Else
                                        dblComisionConasev = CDbl(adoRegistro("ValorAlterno"))
                                    End If
                                Case Codigo_Signo_Mayor
                                    If CInt(txtDiasPlazo.Text) > adoRegistro("CantDias") Then
                                        dblComisionConasev = CDbl(adoRegistro("ValorCosto"))
                                    Else
                                        dblComisionConasev = CDbl(adoRegistro("ValorAlterno"))
                                    End If
                                Case Codigo_Signo_MayorIgual
                                    If CInt(txtDiasPlazo.Text) >= adoRegistro("CantDias") Then
                                        dblComisionConasev = CDbl(adoRegistro("ValorCosto"))
                                    Else
                                        dblComisionConasev = CDbl(adoRegistro("ValorAlterno"))
                                    End If
                                Case Codigo_Signo_Menor
                                    If CInt(txtDiasPlazo.Text) < adoRegistro("CantDias") Then
                                        dblComisionConasev = CDbl(adoRegistro("ValorCosto"))
                                    Else
                                        dblComisionConasev = CDbl(adoRegistro("ValorAlterno"))
                                    End If
                                Case Codigo_Signo_MenorIgual
                                    If CInt(txtDiasPlazo.Text) <= adoRegistro("CantDias") Then
                                        dblComisionConasev = CDbl(adoRegistro("ValorCosto"))
                                    Else
                                        dblComisionConasev = CDbl(adoRegistro("ValorAlterno"))
                                    End If
                            End Select
                        End If
                    Else
                        strCodTipoCostoConasevP = Trim(adoRegistro("TipoCosto"))
                        dblComisionConasevP = CDbl(adoRegistro("ValorCosto"))
                        
                        If CInt(txtDiasPlazo.Text) > 0 Then
                            Select Case adoRegistro("SignoRestriccion")
                                Case Codigo_Signo_Igual
                                    If CInt(txtDiasPlazo.Text) = adoRegistro("CantDias") Then
                                        dblComisionConasevP = CDbl(adoRegistro("ValorCosto"))
                                    Else
                                        dblComisionConasevP = CDbl(adoRegistro("ValorAlterno"))
                                    End If
                                Case Codigo_Signo_Mayor
                                    If CInt(txtDiasPlazo.Text) > adoRegistro("CantDias") Then
                                        dblComisionConasevP = CDbl(adoRegistro("ValorCosto"))
                                    Else
                                        dblComisionConasevP = CDbl(adoRegistro("ValorAlterno"))
                                    End If
                                Case Codigo_Signo_MayorIgual
                                    If CInt(txtDiasPlazo.Text) >= adoRegistro("CantDias") Then
                                        dblComisionConasevP = CDbl(adoRegistro("ValorCosto"))
                                    Else
                                        dblComisionConasevP = CDbl(adoRegistro("ValorAlterno"))
                                    End If
                                Case Codigo_Signo_Menor
                                    If CInt(txtDiasPlazo.Text) < adoRegistro("CantDias") Then
                                        dblComisionConasevP = CDbl(adoRegistro("ValorCosto"))
                                    Else
                                        dblComisionConasevP = CDbl(adoRegistro("ValorAlterno"))
                                    End If
                                Case Codigo_Signo_MenorIgual
                                    If CInt(txtDiasPlazo.Text) <= adoRegistro("CantDias") Then
                                        dblComisionConasevP = CDbl(adoRegistro("ValorCosto"))
                                    Else
                                        dblComisionConasevP = CDbl(adoRegistro("ValorAlterno"))
                                    End If
                            End Select
                        End If
                    End If
                Case Codigo_Costo_Cavali
                    If adoRegistro("TipoPlazo") = Codigo_Operacion_Contado Then
                        strCodTipoCavali = Trim(adoRegistro("TipoCosto"))
                        dblComisionCavali = CDbl(adoRegistro("ValorCosto"))
                        
                        If CInt(txtDiasPlazo.Text) > 0 Then
                            Select Case adoRegistro("SignoRestriccion")
                                Case Codigo_Signo_Igual
                                    If CInt(txtDiasPlazo.Text) = adoRegistro("CantDias") Then
                                        dblComisionCavali = CDbl(adoRegistro("ValorCosto"))
                                    Else
                                        dblComisionCavali = (CDbl(adoRegistro("ValorCosto")) / CDbl(adoRegistro("CantDias"))) * CDbl(txtDiasPlazo.Text)
                                    End If
                                Case Codigo_Signo_Mayor
                                    If CInt(txtDiasPlazo.Text) > adoRegistro("CantDias") Then
                                        dblComisionCavali = CDbl(adoRegistro("ValorCosto"))
                                    Else
                                        dblComisionCavali = (CDbl(adoRegistro("ValorCosto")) / CDbl(adoRegistro("CantDias"))) * CDbl(txtDiasPlazo.Text)
                                    End If
                                Case Codigo_Signo_MayorIgual
                                    If CInt(txtDiasPlazo.Text) >= adoRegistro("CantDias") Then
                                        dblComisionCavali = CDbl(adoRegistro("ValorCosto"))
                                    Else
                                        dblComisionCavali = (CDbl(adoRegistro("ValorCosto")) / CDbl(adoRegistro("CantDias"))) * CDbl(txtDiasPlazo.Text)
                                    End If
                                Case Codigo_Signo_Menor
                                    If CInt(txtDiasPlazo.Text) < adoRegistro("CantDias") Then
                                        dblComisionCavali = CDbl(adoRegistro("ValorCosto"))
                                    Else
                                        dblComisionCavali = (CDbl(adoRegistro("ValorCosto")) / CDbl(adoRegistro("CantDias"))) * CDbl(txtDiasPlazo.Text)
                                    End If
                                Case Codigo_Signo_MenorIgual
                                    If CInt(txtDiasPlazo.Text) <= adoRegistro("CantDias") Then
                                        dblComisionCavali = CDbl(adoRegistro("ValorCosto"))
                                    Else
                                        dblComisionCavali = (CDbl(adoRegistro("ValorCosto")) / CDbl(adoRegistro("CantDias"))) * CDbl(txtDiasPlazo.Text)
                                    End If
                            End Select
                        End If
                    Else
                        strCodTipoCavaliP = Trim(adoRegistro("TipoCosto"))
                        dblComisionCavaliP = CDbl(adoRegistro("ValorCosto"))
                        
                        If CInt(txtDiasPlazo.Text) > 0 Then
                            Select Case adoRegistro("SignoRestriccion")
                                Case Codigo_Signo_Igual
                                    If CInt(txtDiasPlazo.Text) = adoRegistro("CantDias") Then
                                        dblComisionCavaliP = CDbl(adoRegistro("ValorCosto"))
                                    Else
                                        dblComisionCavaliP = (CDbl(adoRegistro("ValorCosto")) / CDbl(adoRegistro("CantDias"))) * CDbl(txtDiasPlazo.Text)
                                    End If
                                Case Codigo_Signo_Mayor
                                    If CInt(txtDiasPlazo.Text) > adoRegistro("CantDias") Then
                                        dblComisionCavaliP = CDbl(adoRegistro("ValorCosto"))
                                    Else
                                        dblComisionCavaliP = (CDbl(adoRegistro("ValorCosto")) / CDbl(adoRegistro("CantDias"))) * CDbl(txtDiasPlazo.Text)
                                    End If
                                Case Codigo_Signo_MayorIgual
                                    If CInt(txtDiasPlazo.Text) >= adoRegistro("CantDias") Then
                                        dblComisionCavaliP = CDbl(adoRegistro("ValorCosto"))
                                    Else
                                        dblComisionCavaliP = (CDbl(adoRegistro("ValorCosto")) / CDbl(adoRegistro("CantDias"))) * CDbl(txtDiasPlazo.Text)
                                    End If
                                Case Codigo_Signo_Menor
                                    If CInt(txtDiasPlazo.Text) < adoRegistro("CantDias") Then
                                        dblComisionCavaliP = CDbl(adoRegistro("ValorCosto"))
                                    Else
                                        dblComisionCavaliP = (CDbl(adoRegistro("ValorCosto")) / CDbl(adoRegistro("CantDias"))) * CDbl(txtDiasPlazo.Text)
                                    End If
                                Case Codigo_Signo_MenorIgual
                                    If CInt(txtDiasPlazo.Text) <= adoRegistro("CantDias") Then
                                        dblComisionCavaliP = CDbl(adoRegistro("ValorCosto"))
                                    Else
                                        dblComisionCavaliP = (CDbl(adoRegistro("ValorCosto")) / CDbl(adoRegistro("CantDias"))) * CDbl(txtDiasPlazo.Text)
                                    End If
                            End Select
                        End If
                    End If
                    
                Case Codigo_Costo_FLiquidacion
                    If adoRegistro("TipoPlazo") = Codigo_Operacion_Contado Then
                        strCodTipoCostoFondo = Trim(adoRegistro("TipoCosto"))
                        dblComisionFondo = CDbl(adoRegistro("ValorCosto"))
                    Else
                        strCodTipoCostoFondoP = Trim(adoRegistro("TipoCosto"))
                        dblComisionFondoP = CDbl(adoRegistro("ValorCosto"))
                    End If
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
        "WHERE TipoValor='" & Codigo_Valor_RentaFija & "' AND TipoPlazo='" & Valor_Caracter & "' AND IndInstrumento='X' AND IndVigente='X' AND IVF.CodEstructura='03' AND " & _
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
            dtpFechaVencimiento.Value = dtpFechaLiquidacion.Value
            strCodMoneda = Trim(adoRegistro("CodMoneda"))
            txtTipoCambio.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, dtpFechaOrden.Value, Codigo_Moneda_Local, strCodMoneda))
            If CDbl(txtTipoCambio.Text) = 0 Then txtTipoCambio.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, DateAdd("d", -1, dtpFechaOrden.Value), Codigo_Moneda_Local, strCodMoneda))
            dblTipoCambio = CDbl(txtTipoCambio.Text)
'            txtTipoCambio.Text = CStr(dblTipoCambio)
                        
            frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
            Me.Refresh
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    '*** Tipo de Instrumento ***
    strSQL = "SELECT FIF.CodFile CODIGO,DescripFile DESCRIP " & _
        "FROM FondoInversionFile FIF JOIN InversionFile IVF ON(IVF.CodFile=FIF.CodFile) " & _
        "WHERE TipoValor='" & Codigo_Valor_RentaFija & "' AND TipoPlazo='" & Valor_Caracter & "' AND IndInstrumento='X' AND IndVigente='X' AND IVF.CodEstructura='03' AND " & _
        "CodFondo='" & strCodFondoOrden & "' AND CodAdministradora='" & gstrCodAdministradora & "' ORDER BY DescripFile"
    CargarControlLista strSQL, cboTipoInstrumentoOrden, arrTipoInstrumentoOrden(), Sel_Defecto
            
End Sub

Private Sub cboMoneda_Click()
    
    lblDescripMoneda(0).Caption = "S/.": lblDescripMoneda(0).Tag = Codigo_Moneda_Local
    lblDescripMoneda(1).Caption = "S/.": lblDescripMoneda(1).Tag = Codigo_Moneda_Local
    lblDescripMonedaResumen(0) = "S/.": lblDescripMonedaResumen(0).Tag = Codigo_Moneda_Local
    lblDescripMonedaResumen(1) = "S/.": lblDescripMonedaResumen(1).Tag = Codigo_Moneda_Local
    
    If cboMoneda.ListIndex < 0 Then Exit Sub
    
    strCodMoneda = Trim(arrMoneda(cboMoneda.ListIndex))
    
    txtTipoCambio.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, dtpFechaOrden.Value, Codigo_Moneda_Local, strCodMoneda))
    If CDbl(txtTipoCambio.Text) = 0 Then txtTipoCambio.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, DateAdd("d", -1, dtpFechaOrden.Value), Codigo_Moneda_Local, strCodMoneda))
    dblTipoCambio = CDbl(txtTipoCambio)
        
    lblDescripMoneda(0).Caption = ObtenerSignoMoneda(strCodMoneda)
    lblDescripMoneda(0).Tag = strCodMoneda
    lblDescripMoneda(1).Caption = ObtenerSignoMoneda(strCodMoneda)
    lblDescripMoneda(1).Tag = strCodMoneda
    lblDescripMonedaResumen(0).Caption = ObtenerSignoMoneda(strCodMoneda)
    lblDescripMonedaResumen(0).Tag = strCodMoneda
    lblDescripMonedaResumen(1).Caption = ObtenerSignoMoneda(strCodMoneda)
    lblDescripMonedaResumen(1).Tag = strCodMoneda
    lblMoneda.Caption = ObtenerDescripcionMoneda(strCodMoneda)
    
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

    Dim adoRegistro As ADODB.Recordset
    
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
    Set adoRegistro = New ADODB.Recordset
    With adoComm
        .CommandText = "SELECT DescripFile,DescripInicial FROM InversionFile WHERE CodFile='" & strCodTipoInstrumentoOrden & "'"
        Set adoRegistro = .Execute
            
        If Not adoRegistro.EOF Then
            txtNemonico.Text = Trim(adoRegistro("DescripInicial")) & CStr(gdatFechaActual)
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
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
            strSQL = "SELECT CodTitulo CODIGO,(Nemotecnico + ' ' + DescripTitulo) DESCRIP FROM InstrumentoInversion " & _
                "WHERE CodFile='004' AND IndVigente='X' ORDER BY DescripTitulo"
                
            CargarControlLista strSQL, cboTitulo, arrTitulo(), Sel_Defecto
        
            If cboTitulo.ListCount > 0 Then cboTitulo.ListIndex = 0
            
        Case Codigo_Orden_Venta
            
            strSQL = "{call up_IVLstOperacionesVigentes ('" & strCodFondoOrden & "','" & gstrCodAdministradora & "','" & strCodFile & "','" & gstrFechaActual & "') }"
'            strSQL = "SELECT InstrumentoInversion.CodTitulo CODIGO," & _
'                "(RTRIM(InstrumentoInversion.CodTitulo) + ' ' + RTRIM(InstrumentoInversion.Nemotecnico) + ' ' + RTRIM(InstrumentoInversion.DescripTitulo)) DESCRIP FROM InstrumentoInversion,InversionKardex " & _
'                "WHERE SaldoFinal > 0 AND IndUltimoMovimiento='X' AND InstrumentoInversion.CodFile=InversionKardex.CodFile AND " & _
'                "InstrumentoInversion.CodAnalitica=InversionKardex.CodAnalitica AND InversionKardex.CodFile='" & strCodFile & "' AND InversionInversion.CodDetalleFile='" & strCodClaseInstrumento & "' AND " & _
'                "InstrumentoInversion.CodFondo='" & strCodFondoOrden & "' AND InversionKardex.CodFondo='" & strCodFondoOrden & "' " & _
'                "ORDER BY InstrumentoInversion.Nemotecnico"
                            
            CargarControlLista strSQL, cboTitulo, arrTitulo(), Sel_Defecto
        
            If cboTitulo.ListCount > 0 Then cboTitulo.ListIndex = 0
                            
    End Select
    Me.MousePointer = vbDefault
    
End Sub

Private Sub cboTitulo_Click()

    Dim adoConsulta     As ADODB.Recordset
    Dim intRegistro     As Integer
    
    strCodGarantia = Valor_Caracter: txtDescripOrden = Valor_Caracter: strCodAnalitica = Valor_Caracter
    lblAnalitica = strCodTipoInstrumentoOrden & "-????????"
    strCodEmisor = Valor_Caracter: strCodGrupo = Valor_Caracter
    If cboTitulo.ListIndex < 0 Then Exit Sub

    strCodGarantia = Trim(arrTitulo(cboTitulo.ListIndex))

    With adoComm
        Set adoConsulta = New ADODB.Recordset

        .CommandText = "SELECT CodAnalitica,ValorNominal,CodMoneda,CodEmisor,CodGrupo FROM InstrumentoInversion WHERE CodTitulo='" & strCodGarantia & "'"
        Set adoConsulta = .Execute

        If Not adoConsulta.EOF Then
            strCodAnalitica = Trim(adoConsulta("CodAnalitica"))
            strCodMonedaGarantia = Trim(adoConsulta("CodMoneda"))
            
            lblMoneda.Caption = ObtenerDescripcionMoneda(strCodMonedaGarantia)
                
            strCodEmisor = Trim(adoConsulta("CodEmisor")): strCodGrupo = Trim(adoConsulta("CodGrupo"))
        End If
        adoConsulta.Close: Set adoConsulta = Nothing

        If strCodGarantia <> Valor_Caracter Then
            '*** Validar Limites ***
            If Not PosicionLimites() Then Exit Sub
        End If
    End With

    txtDescripOrden = Trim(cboTipoInstrumentoOrden.Text) & " - " & Left(cboTitulo.Text, 15)
    
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
    If CInt(txtDiasPlazo.Text) > 0 And CCur(txtSubTotal(0).Text) > 0 Then
        If CCur(txtSubTotal(1).Text) = 0 Or CCur(txtSubTotal(0).Text) = 0 Then
            MsgBox "Por favor verificar que el SubTotal al Contado y a Plazo tengan valores.", vbExclamation, Me.Caption
            Exit Sub
        End If
        dblFactor = (CCur(txtSubTotal(1).Text) / CCur(txtSubTotal(0).Text)) ^ (360 / CInt(txtDiasPlazo.Text))
        lblTirBruta.Caption = CStr((dblFactor - 1) * 100)
        lblTirBrutaResumen.Caption = lblTirBruta.Caption
    End If
    
    '*** Tir Neta ***
    If CInt(txtDiasPlazo.Text) > 0 And CCur(lblMontoTotal(0).Caption) > 0 Then
        dblFactor = (CCur(lblMontoTotal(1).Caption) / CCur(lblMontoTotal(0).Caption)) ^ (360 / CInt(txtDiasPlazo.Text))
        lblTirNeta.Caption = CStr((dblFactor - 1) * 100)
        lblTirNetaResumen.Caption = lblTirNeta.Caption
    End If
    
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

Private Sub dtpFechaLiquidacion_Change()

    If dtpFechaLiquidacion.Value < dtpFechaOrden.Value Then
        dtpFechaLiquidacion.Value = dtpFechaOrden.Value
    End If
    
    If Not EsDiaUtil(dtpFechaLiquidacion.Value) Then
        MsgBox "La Fecha no es un día útil...se cambiará por una fecha correcta !", vbInformation, Me.Caption
        dtpFechaLiquidacion.Value = ProximoDiaUtil(dtpFechaLiquidacion.Value)
    End If
    dtpFechaVencimiento.Value = dtpFechaLiquidacion.Value
    lblFechaLiquidacion.Caption = CStr(dtpFechaLiquidacion.Value)
    
End Sub

Private Sub dtpFechaLiquidacion_LostFocus()

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

Private Sub dtpFechaPago_Change()

    If dtpFechaPago.Value < dtpFechaVencimiento.Value Then
        dtpFechaPago.Value = dtpFechaVencimiento.Value
    End If
    
    If Not EsDiaUtil(dtpFechaPago.Value) Then
        MsgBox "La Fecha de Pago no es un día útil...se cambiará por una fecha correcta !", vbInformation, Me.Caption
        dtpFechaPago.Value = ProximoDiaUtil(dtpFechaPago.Value)
    End If
    
End Sub

Private Sub dtpFechaVencimiento_Change()

    If dtpFechaVencimiento.Value < dtpFechaOrden.Value Then
        dtpFechaVencimiento.Value = dtpFechaOrden.Value
    End If
    
    If dtpFechaVencimiento.Value < dtpFechaLiquidacion.Value Then
        dtpFechaVencimiento.Value = dtpFechaLiquidacion.Value
    End If
    
'    If Not EsDiaUtil(dtpFechaVencimiento.Value) Then
'        MsgBox "La Fecha no es un día útil...se cambiará por una fecha correcta !", vbInformation, Me.Caption
'        dtpFechaVencimiento.Value = ProximoDiaUtil(dtpFechaVencimiento.Value)
'    End If
    
    txtDiasPlazo.Text = CStr(DateDiff("d", dtpFechaLiquidacion.Value, dtpFechaVencimiento.Value))
    Call CalculoTotal(0)
    dtpFechaPago.Value = dtpFechaVencimiento.Value
    lblFechaVencimiento.Caption = CStr(dtpFechaVencimiento.Value)
    lblFechaCupon.Caption = CStr(dtpFechaVencimiento.Value)
    
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

Private Sub lblPrecio_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblPrecio(Index), Decimales_Precio)
    
    lblPrecioResumen(Index).Caption = CStr(lblPrecio(Index))
    
End Sub


Private Sub lblPrecioResumen_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblPrecioResumen(Index), Decimales_Monto)
    
End Sub

Private Sub lblStockNominal_Change()

    Call FormatoMillarEtiqueta(lblStockNominal, Decimales_Monto)
    
End Sub

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

Private Sub tabReporte_Click(PreviousTab As Integer)

    Select Case tabReporte.Tab
        Case 1, 2
            If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vQuery)
            If strEstado = Reg_Defecto Then tabReporte.Tab = 0
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

    Call FormatoCajaTexto(txtCantidad, Decimales_Monto)
    lblCantidadResumen.Caption = CStr(txtCantidad)
    
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtCantidad, Decimales_Monto)
    
End Sub


Private Sub txtComisionAgente_Change(Index As Integer)

    Call FormatoCajaTexto(txtComisionAgente(Index), Decimales_Monto)
    
End Sub

Private Sub txtComisionAgente_KeyPress(Index As Integer, KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtComisionAgente(Index), Decimales_Monto)
    
    If KeyAscii = vbKeyReturn Then
        If chkAplicar(Index).Value Then
            ActualizaPorcentaje txtComisionAgente(Index), txtPorcenAgente(Index)
        End If
        Call CalculoImpuesto(Index)
        Call CalculoTotal(Index)
    End If
    
End Sub


Private Sub txtComisionBolsa_Change(Index As Integer)

    Call FormatoCajaTexto(txtComisionBolsa(Index), Decimales_Monto)
    
    If chkAplicar(Index).Value Then
        ActualizaPorcentaje txtComisionBolsa(Index), lblPorcenBolsa(Index)
    End If
    
End Sub


Private Sub txtComisionBolsa_KeyPress(Index As Integer, KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtComisionBolsa(Index), Decimales_Monto)
    
    If KeyAscii = vbKeyReturn Then
        Call CalculoImpuesto(Index)
        Call CalculoTotal(Index)
    End If
    
End Sub


Private Sub txtComisionCavali_Change(Index As Integer)

    Call FormatoCajaTexto(txtComisionCavali(Index), Decimales_Monto)
    
    If chkAplicar(Index).Value Then
        ActualizaPorcentaje txtComisionCavali(Index), lblPorcenCavali(Index)
    End If
    
End Sub

Private Sub txtComisionCavali_KeyPress(Index As Integer, KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtComisionCavali(Index), Decimales_Monto)
    
    If KeyAscii = vbKeyReturn Then
        Call CalculoImpuesto(Index)
        Call CalculoTotal(Index)
    End If
    
End Sub

Private Sub txtComisionConasev_Change(Index As Integer)

    Call FormatoCajaTexto(txtComisionConasev(Index), Decimales_Monto)
    
    If chkAplicar(Index).Value Then
        ActualizaPorcentaje txtComisionConasev(Index), lblPorcenConasev(Index)
    End If
    
End Sub

Private Sub txtComisionConasev_KeyPress(Index As Integer, KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtComisionConasev(Index), Decimales_Monto)
    
    If KeyAscii = vbKeyReturn Then
        Call CalculoImpuesto(Index)
        Call CalculoTotal(Index)
    End If
    
End Sub

Private Sub txtComisionFondo_Change(Index As Integer)

    Call FormatoCajaTexto(txtComisionFondo(Index), Decimales_Monto)
    
    If chkAplicar(Index).Value Then
        ActualizaPorcentaje txtComisionFondo(Index), lblPorcenFondo(Index)
    End If
    
End Sub

Private Sub txtComisionFondo_KeyPress(Index As Integer, KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtComisionFondo(Index), Decimales_Monto)
    
    If KeyAscii = vbKeyReturn Then
        Call CalculoImpuesto(Index)
        Call CalculoTotal(Index)
    End If
    
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

Private Sub txtDiasPlazo_Change()
   
    Call FormatoCajaTexto(txtDiasPlazo, 0)
    
    If IsNumeric(txtDiasPlazo.Text) Then
        dtpFechaVencimiento.Value = DateAdd("d", txtDiasPlazo.Text, CVDate(dtpFechaLiquidacion.Value))
    Else
        dtpFechaVencimiento.Value = dtpFechaLiquidacion.Value
    End If

    Call cboConceptoCosto_Click
    Call CalculoTotal(0)
    
    dtpFechaPago.Value = dtpFechaVencimiento.Value
    dtpFechaPago_Change
    lblDiasPlazo.Caption = CStr(txtDiasPlazo.Text)
    lblFechaVencimiento.Caption = CStr(dtpFechaVencimiento.Value)
    lblFechaCupon.Caption = CStr(dtpFechaVencimiento.Value)
    
End Sub

Private Sub txtInteresCorrido_Change(Index As Integer)

    Call FormatoCajaTexto(txtInteresCorrido(Index), Decimales_Monto)
    
End Sub

Private Sub txtInteresCorrido_KeyPress(Index As Integer, KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtInteresCorrido(Index), Decimales_Monto)
    
    If KeyAscii = vbKeyReturn Then
        Call CalculoTotal(Index)
    End If
    
End Sub

Private Sub txtNemonico_Change()

    txtDescripOrden.Text = Trim(cboTipoInstrumentoOrden.Text) & " - " & Trim(txtNemonico.Text)
    
End Sub

Private Sub txtNemonico_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
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

Private Sub txtPrecioMercado_Change()

    Call FormatoCajaTexto(txtPrecioMercado, Decimales_Precio)
    'txtValorNominal_Change
    
End Sub

Private Sub txtPrecioMercado_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtPrecioMercado, Decimales_Precio)
    
End Sub

Private Sub txtSubTotal_Change(Index As Integer)

    Call FormatoCajaTexto(txtSubTotal(Index), Decimales_Monto)
    
    If CLng(txtCantidad.Text) > 0 Then
        lblPrecio(Index).Caption = CStr(CCur(txtSubTotal(Index).Text) / CLng(txtCantidad.Text))
    End If
    lblSubTotalResumen(Index).Caption = CStr(CCur(txtSubTotal(Index)))
    
End Sub

Private Sub txtSubTotal_KeyPress(Index As Integer, KeyAscii As Integer)
    
    Call ValidaCajaTexto(KeyAscii, "M", txtSubTotal(Index), Decimales_Monto)
    
End Sub

Private Sub txtTasaMensual_Change()

    Call FormatoCajaTexto(txtTasaMensual, Decimales_Tasa)
    
    If CLng(txtDiasPlazo.Text) = 0 And tabReporte.Tab = 2 Then
        MsgBox "Por favor especifique el Plazo del Reporte", vbCritical, Me.Caption
        tabReporte.Tab = 1
        txtDiasPlazo.SetFocus
        Exit Sub
    End If
    
    If strCodMoneda = strCodMonedaGarantia Then
        txtSubTotal(1).Text = CStr(((CDbl(txtTasaMensual.Text) / 100 + 1) ^ (CLng(txtDiasPlazo.Text) / 30)) * CDbl(txtPrecioMercado.Text) * CLng(txtCantidad.Text))
    ElseIf strCodMoneda = Codigo_Moneda_Local And strCodMonedaGarantia <> Codigo_Moneda_Local Then
        txtSubTotal(1).Text = CStr(((CDbl(txtTasaMensual.Text) / 100 + 1) ^ (CLng(txtDiasPlazo.Text) / 30)) * (CDbl(txtPrecioMercado.Text) * CDbl(txtTipoCambio.Text)) * CLng(txtCantidad.Text))
    ElseIf strCodMoneda <> Codigo_Moneda_Local And strCodMonedaGarantia = Codigo_Moneda_Local Then
        txtSubTotal(1).Text = CStr(((CDbl(txtTasaMensual.Text) / 100 + 1) ^ (CLng(txtDiasPlazo.Text) / 30)) * (CDbl(txtPrecioMercado.Text) / CDbl(txtTipoCambio.Text)) * CLng(txtCantidad.Text))
    End If
    
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

Private Sub ExportarExcel()
    
    Dim adoRegistro As ADODB.Recordset
    Dim execSQL As String
    Dim rutaExportacion As String
    
    Dim datFechaSiguiente As Date
    Dim strFechaLiquidacionHasta As String
    
    Set frmFormulario = frmOrdenReporteRentaVariable
    
    Set adoRegistro = New ADODB.Recordset
    
    'If TodoOK() Then
        
        Dim strNameProc As String
        
        gstrNameRepo = "OrdenReporteRentaVariable"
        
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
