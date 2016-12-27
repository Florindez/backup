VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmFondoGastos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comisiones y Gastos del Fondo"
   ClientHeight    =   7035
   ClientLeft      =   855
   ClientTop       =   1050
   ClientWidth     =   11550
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7035
   ScaleWidth      =   11550
   Begin MSAdodcLib.Adodc adoConsulta 
      Height          =   330
      Left            =   8130
      Top             =   5130
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
   Begin TabDlg.SSTab tabGasto 
      Height          =   6075
      Left            =   30
      TabIndex        =   8
      Top             =   30
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   10716
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      ForeColor       =   8388608
      TabCaption(0)   =   "Lista"
      TabPicture(0)   =   "frmFondoGastos.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "tdgConsulta"
      Tab(0).Control(1)=   "fraGastos(0)"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "frmFondoGastos.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraGastos(1)"
      Tab(1).Control(1)=   "lblCodProveedor"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Condiciones"
      TabPicture(2)   =   "frmFondoGastos.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "lblDescrip(18)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame3"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame1"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      Begin VB.Frame Frame1 
         Caption         =   "Datos de Cálculo"
         Height          =   1245
         Left            =   180
         TabIndex        =   60
         Top             =   480
         Width           =   11145
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
            Left            =   8040
            Style           =   2  'Dropdown List
            TabIndex        =   81
            Top             =   720
            Visible         =   0   'False
            Width           =   2595
         End
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
            Height          =   315
            Left            =   8040
            Style           =   2  'Dropdown List
            TabIndex        =   79
            Top             =   330
            Width           =   2595
         End
         Begin VB.ComboBox cboModalidadCalculo 
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
            Left            =   2490
            Style           =   2  'Dropdown List
            TabIndex        =   77
            Top             =   330
            Width           =   3225
         End
         Begin VB.ComboBox cboPeriodoCalculo 
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
            Left            =   3630
            Style           =   2  'Dropdown List
            TabIndex        =   76
            Top             =   1680
            Visible         =   0   'False
            Width           =   2085
         End
         Begin VB.TextBox txtValorPeriodoCalculo 
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
            Left            =   2460
            TabIndex        =   75
            Text            =   "1"
            Top             =   1680
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.ComboBox cboTipoCalculo 
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
            Left            =   2460
            Style           =   2  'Dropdown List
            TabIndex        =   70
            Top             =   1290
            Visible         =   0   'False
            Width           =   3255
         End
         Begin VB.TextBox txtValorFrecuenciaCalculo 
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
            Height          =   315
            Left            =   2490
            TabIndex        =   65
            Text            =   "1"
            Top             =   720
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.ComboBox cboFrecuenciaCalculo 
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
            Left            =   3660
            Style           =   2  'Dropdown List
            TabIndex        =   64
            Top             =   720
            Visible         =   0   'False
            Width           =   2055
         End
         Begin MSComCtl2.DTPicker dtpFechaPrimerCalculo 
            Height          =   315
            Left            =   8040
            TabIndex        =   68
            Top             =   1320
            Visible         =   0   'False
            Width           =   1485
            _ExtentX        =   2619
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
            Format          =   175767553
            CurrentDate     =   38068
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Periodo:"
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   22
            Left            =   390
            TabIndex        =   86
            Top             =   1350
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Cada"
            ForeColor       =   &H00800000&
            Height          =   165
            Index           =   11
            Left            =   1770
            TabIndex        =   85
            Top             =   780
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.Line Line2 
            Visible         =   0   'False
            X1              =   360
            X2              =   10590
            Y1              =   1140
            Y2              =   1140
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Base de Cálculo"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   32
            Left            =   6120
            TabIndex        =   82
            Top             =   780
            Visible         =   0   'False
            Width           =   1380
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Forma de Cálculo"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   33
            Left            =   6120
            TabIndex        =   80
            Top             =   390
            Width           =   1470
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Modalidad de Cálculo"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   21
            Left            =   390
            TabIndex        =   78
            Top             =   390
            Width           =   1830
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Cada"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   20
            Left            =   1800
            TabIndex        =   74
            Top             =   1740
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "1era. Fecha Corte"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   42
            Left            =   6120
            TabIndex        =   69
            Top             =   1380
            Visible         =   0   'False
            Width           =   1545
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Frecuencia:"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   5
            Left            =   390
            TabIndex        =   66
            Top             =   780
            Visible         =   0   'False
            Width           =   1020
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Contables y Tributarias"
         Height          =   1515
         Left            =   180
         TabIndex        =   40
         Top             =   4140
         Visible         =   0   'False
         Width           =   11145
         Begin VB.CheckBox chkGeneraCreditoFiscal 
            Caption         =   "Genera Crédito Fiscal"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   2460
            TabIndex        =   73
            Top             =   1110
            Width           =   2295
         End
         Begin VB.ComboBox cboAfectacion 
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
            Left            =   2460
            Style           =   2  'Dropdown List
            TabIndex        =   72
            Top             =   300
            Width           =   2055
         End
         Begin VB.ComboBox cboCreditoFiscal 
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
            Left            =   2460
            Style           =   2  'Dropdown List
            TabIndex        =   62
            Top             =   690
            Width           =   4335
         End
         Begin VB.TextBox txtBaseCalculo 
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
            Height          =   315
            Left            =   2850
            TabIndex        =   51
            Top             =   4980
            Width           =   1845
         End
         Begin VB.CheckBox chkSinVencimiento 
            Caption         =   "Sin Vencimiento"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   5010
            TabIndex        =   49
            Top             =   2490
            Width           =   3315
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Impuesto"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   39
            Left            =   240
            TabIndex        =   71
            Top             =   360
            Width           =   780
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Crédito Fiscal"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   12
            Left            =   240
            TabIndex        =   63
            Top             =   720
            Width           =   1605
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto Cálculo"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   16
            Left            =   570
            TabIndex        =   53
            Top             =   5010
            Width           =   1215
         End
         Begin VB.Label lblMoneda 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   4740
            TabIndex        =   52
            Top             =   5010
            Width           =   360
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Importe"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   15
            Left            =   660
            TabIndex        =   50
            Top             =   4410
            Width           =   1080
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Fecha de Pago"
         Height          =   2235
         Left            =   180
         TabIndex        =   28
         Top             =   1800
         Width           =   11145
         Begin VB.ComboBox cboModalidadPago 
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
            Left            =   2460
            Style           =   2  'Dropdown List
            TabIndex        =   83
            Top             =   330
            Width           =   3285
         End
         Begin VB.ComboBox cboTipoPago 
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
            Left            =   2460
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   870
            Width           =   2655
         End
         Begin VB.TextBox txtValorPeriodoPago 
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
            Height          =   315
            Left            =   6750
            TabIndex        =   36
            Text            =   "1"
            Top             =   870
            Width           =   1125
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
            Height          =   315
            Left            =   7920
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   870
            Width           =   2595
         End
         Begin VB.ComboBox cboTipoDesplazamiento 
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
            Left            =   6750
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   1290
            Width           =   3765
         End
         Begin VB.CheckBox chkIndicadorPagoFinMes 
            Caption         =   "Fin de Periodo"
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   360
            TabIndex        =   29
            Top             =   1710
            Width           =   1815
         End
         Begin MSComCtl2.DTPicker dtpFechaPrimerPago 
            Height          =   315
            Left            =   2460
            TabIndex        =   30
            Top             =   1290
            Width           =   1485
            _ExtentX        =   2619
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
            Format          =   175767553
            CurrentDate     =   38068
         End
         Begin VB.Line Line4 
            X1              =   360
            X2              =   10590
            Y1              =   750
            Y2              =   750
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Modalidad de Pago"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   31
            Left            =   360
            TabIndex        =   84
            Top             =   420
            Width           =   1650
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   29
            Left            =   360
            TabIndex        =   38
            Top             =   930
            Width           =   390
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Cada"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   28
            Left            =   6150
            TabIndex        =   35
            Top             =   930
            Width           =   450
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Desplazamiento"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   24
            Left            =   4830
            TabIndex        =   33
            Top             =   1350
            Width           =   1785
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "1era. Fecha Pago"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   23
            Left            =   360
            TabIndex        =   31
            Top             =   1350
            Width           =   1530
         End
      End
      Begin VB.Frame fraGastos 
         Caption         =   "General"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5445
         Index           =   1
         Left            =   -74850
         TabIndex        =   12
         Top             =   480
         Width           =   11130
         Begin VB.ComboBox cboTipoProveedor 
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
            Left            =   360
            Style           =   2  'Dropdown List
            TabIndex        =   89
            Top             =   1710
            Width           =   1845
         End
         Begin VB.CheckBox chkNoIncluyeEnBalancePrecierre 
            Caption         =   "No incluye en Balance de Precierre"
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   2640
            TabIndex        =   88
            Top             =   4590
            Width           =   3975
         End
         Begin VB.CheckBox chkIndicadorGastoIterativo 
            Caption         =   "Gasto Iterativo"
            ForeColor       =   &H00800000&
            Height          =   150
            Left            =   2640
            TabIndex        =   87
            Top             =   4320
            Width           =   3315
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
            Left            =   2640
            Style           =   2  'Dropdown List
            TabIndex        =   58
            Top             =   4980
            Width           =   3285
         End
         Begin VB.ComboBox cboFormula 
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
            ItemData        =   "frmFondoGastos.frx":0054
            Left            =   6480
            List            =   "frmFondoGastos.frx":0061
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   3450
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.ComboBox cboTipoGasto 
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
            ItemData        =   "frmFondoGastos.frx":008A
            Left            =   2640
            List            =   "frmFondoGastos.frx":008C
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Top             =   3090
            Width           =   2655
         End
         Begin VB.TextBox txtMontoGasto 
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
            Height          =   315
            Left            =   2640
            TabIndex        =   6
            Top             =   3450
            Width           =   1755
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
            Left            =   2640
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   2700
            Width           =   2655
         End
         Begin VB.CommandButton cmdProveedor 
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
            Height          =   315
            Left            =   10200
            TabIndex        =   4
            ToolTipText     =   "Buscar Proveedor"
            Top             =   1680
            Width           =   375
         End
         Begin VB.ComboBox cboGasto 
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
            Left            =   2640
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   960
            Width           =   7965
         End
         Begin VB.TextBox txtDescripGasto 
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
            Left            =   2640
            MaxLength       =   60
            TabIndex        =   3
            Top             =   1335
            Width           =   7935
         End
         Begin MSComCtl2.DTPicker dtpFechaGasto 
            Height          =   315
            Left            =   2640
            TabIndex        =   1
            Top             =   330
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
            Format          =   175767553
            CurrentDate     =   38068
         End
         Begin MSComCtl2.DTPicker dtpFechaInicio 
            Height          =   315
            Left            =   2640
            TabIndex        =   54
            Top             =   3840
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
            Format          =   175767553
            CurrentDate     =   38068
         End
         Begin MSComCtl2.DTPicker dtpFechaFin 
            Height          =   315
            Left            =   4830
            TabIndex        =   55
            Top             =   3840
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
            Format          =   175767553
            CurrentDate     =   38068
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Gasto"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   19
            Left            =   360
            TabIndex        =   67
            Top             =   3150
            Width           =   945
         End
         Begin VB.Line Line1 
            X1              =   330
            X2              =   10740
            Y1              =   2520
            Y2              =   2520
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Del"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   9
            Left            =   2220
            TabIndex        =   61
            Top             =   3930
            Width           =   300
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Estado"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   17
            Left            =   360
            TabIndex        =   59
            Top             =   5010
            Width           =   600
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Al"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   37
            Left            =   4530
            TabIndex        =   57
            Top             =   3900
            Width           =   180
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Vigencia Gasto"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   36
            Left            =   360
            TabIndex        =   56
            Top             =   3930
            Width           =   1305
         End
         Begin VB.Label lblFormula 
            AutoSize        =   -1  'True
            Caption         =   "Formula"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   5700
            TabIndex        =   48
            Top             =   3510
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   10
            Left            =   360
            TabIndex        =   43
            Top             =   2760
            Width           =   690
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Base Imponible (V.Venta)"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   360
            TabIndex        =   42
            Top             =   3510
            Width           =   2160
         End
         Begin VB.Label lblMagnitud 
            AutoSize        =   -1  'True
            Caption         =   "Magnitud"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   4500
            TabIndex        =   41
            Top             =   3510
            Width           =   795
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "xxxx"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   26
            Left            =   120
            TabIndex        =   27
            Top             =   7500
            Width           =   375
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "xxxx"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   8
            Left            =   600
            TabIndex        =   26
            Top             =   7560
            Width           =   375
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "xxx"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   4
            Left            =   1050
            TabIndex        =   24
            Top             =   7500
            Width           =   285
         End
         Begin VB.Line Line3 
            Index           =   0
            X1              =   300
            X2              =   10590
            Y1              =   810
            Y2              =   810
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Documento ID"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   14
            Left            =   360
            TabIndex        =   23
            Top             =   2160
            Width           =   1335
         End
         Begin VB.Label lblNumDocID 
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
            Left            =   5220
            TabIndex        =   22
            Top             =   2100
            Width           =   2655
         End
         Begin VB.Label lblTipoDocID 
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
            Left            =   2640
            TabIndex        =   21
            Top             =   2100
            Width           =   2535
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   13
            Left            =   360
            TabIndex        =   19
            Top             =   1785
            Width           =   1275
         End
         Begin VB.Label lblProveedor 
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
            Left            =   2640
            TabIndex        =   18
            Top             =   1710
            Width           =   7500
         End
         Begin VB.Label lblSaldoProvision 
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
            Left            =   1620
            TabIndex        =   7
            Top             =   7710
            Width           =   1905
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Provisión"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   7
            Left            =   570
            TabIndex        =   17
            Top             =   7770
            Width           =   1335
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Concepto"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   6
            Left            =   360
            TabIndex        =   16
            Top             =   1020
            Width           =   825
         End
         Begin VB.Label lblAnalitica 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "099-00000000"
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
            Left            =   8460
            TabIndex        =   15
            Top             =   360
            Width           =   1395
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Detalle"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   14
            Top             =   1410
            Width           =   615
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   13
            Top             =   405
            Width           =   540
         End
      End
      Begin VB.Frame fraGastos 
         Caption         =   "Criterios de Búsqueda"
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
         Height          =   1425
         Index           =   0
         Left            =   -74850
         TabIndex        =   9
         Top             =   480
         Width           =   11145
         Begin VB.ComboBox cboFondoSerie 
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
            Left            =   1710
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   690
            Visible         =   0   'False
            Width           =   4185
         End
         Begin VB.CheckBox chkVerSoloGastosVigentes 
            Caption         =   "Ver sólo gastos vigentes"
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   360
            TabIndex        =   39
            Top             =   1080
            Width           =   2595
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
            Left            =   1710
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   300
            Width           =   6315
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Serie"
            ForeColor       =   &H00800000&
            Height          =   315
            Index           =   34
            Left            =   360
            TabIndex        =   45
            Top             =   750
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fondo"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   330
            TabIndex        =   11
            Top             =   390
            Width           =   540
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmFondoGastos.frx":008E
         Height          =   3915
         Left            =   -74850
         OleObjectBlob   =   "frmFondoGastos.frx":00A8
         TabIndex        =   10
         Top             =   1980
         Width           =   11145
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         Caption         =   "0"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   18
         Left            =   450
         TabIndex        =   25
         Top             =   5760
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label lblCodProveedor 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   -74700
         TabIndex        =   20
         Top             =   8820
         Visible         =   0   'False
         Width           =   645
      End
   End
   Begin TAMControls2.ucBotonEdicion2 cmdOpcion 
      Height          =   735
      Left            =   780
      TabIndex        =   91
      Top             =   6210
      Width           =   4200
      _ExtentX        =   7408
      _ExtentY        =   1296
      Buttons         =   3
      Caption0        =   "&Nuevo "
      Tag0            =   "0"
      ToolTipText0    =   "Nuevo"
      Caption1        =   "&Modificar"
      Tag1            =   "3"
      ToolTipText1    =   "Modificar"
      Caption2        =   "&Anular"
      Tag2            =   "4"
      ToolTipText2    =   "Anular"
      UserControlWidth=   4200
   End
   Begin TAMControls2.ucBotonEdicion2 cmdAccion 
      Height          =   735
      Left            =   5280
      TabIndex        =   92
      Top             =   6210
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
      Left            =   9510
      TabIndex        =   90
      Top             =   6210
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
      Caption0        =   "&Salir"
      Tag0            =   "9"
      ToolTipText0    =   "Salir"
      UserControlWidth=   1200
   End
End
Attribute VB_Name = "frmFondoGastos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrFondo()                  As String, arrTipoProvision()           As String
Dim arrMoneda()                 As String, arrFrecuenciaDevengo()       As String
Dim arrCuentaGasto()            As String, arrGasto()                   As String
Dim arrTipoPago()               As String, arrCreditoFiscal()           As String
Dim arrTipoValor()              As String, arrEstado()                  As String
Dim arrPeriodoPago()            As String, arrTipoDesplazamiento()      As String
Dim arrModalidadPago()          As String, arrFormaCalculo()             As String
Dim arrAplicacionDevengo()      As String, arrPeriodoTasa()             As String
Dim arrTipoTasa()               As String, arrTipoCalculo()               As String
Dim arrBaseCalculo()            As String, arrFondoSerie()              As String
Dim arrFrecuenciaGasto()        As String, arrFormula()                 As String
Dim arrTipoGasto()              As String, arrModalidadCalculo()        As String
Dim arrTipoProveedor()          As String
Dim arrPeriodoGasto()           As String, arrFrecuenciaCalculo()          As String
Dim arrDevengo()                As String, arrAfectacion()              As String
Dim arrPeriodoCalculo()         As String, arrUnidadesPeriodo()         As String

Dim strCodFondo                 As String, strCodTipoProvision          As String
Dim strCodMoneda                As String, strCodFrecuenciaDevengo      As String
Dim strCodCuenta                As String, strCodGasto                  As String
Dim strCodFile                  As String, strCodAnalitica              As String
Dim strCodTipoPago              As String, strCodDetalleGasto           As String
Dim strCodCreditoFiscal         As String, strCodTipoDesplazamiento     As String
Dim strEstado                   As String, strSQL                       As String
Dim strCodTipoValor             As String, strCodPeriodoPago            As String
Dim strCodModalidadPago         As String, strCodFormaCalculo            As String
Dim strCodAplicacionDevengo     As String, strCodPeriodoTasa            As String
Dim strCodTipoTasa              As String, strCodTipoCalculo              As String
Dim strCodBaseCalculo           As String, strEstadoGasto               As String
Dim strCodFondoSerie            As String, strCodFormula                As String
Dim strCodFrecuenciaGasto       As String, strIndGastoIterativo         As String
Dim strCodPeriodoGasto          As String, strCodFrecuenciaCalculo         As String
Dim strDevengo                  As String, strCodAfectacion             As String
Dim strIndGeneraCreditoFiscal   As String, strCodPeriodoCalculo         As String
Dim strCodUnidadesPeriodo       As String, strCodTipoProveedor          As String

Dim intNumPeriodo               As Integer, strFechaInicio              As String
Dim strFechaFin                 As String, strFechaPago                 As String
Dim intCantDias                 As Long, strIndVigente                     As String
Dim intSecuencialGasto          As Integer, intNumSecuencial            As Integer
Dim strCodTipoGasto             As String, strCodModalidadCalculo       As String

Dim fechaLimiteInferior As Date

Public Sub Buscar()

'    strSQL = "SELECT FG.CodGasto,FG.CodDetalleGasto,NumGasto,DCG.CodAnalitica,CG.DescripConcepto,DCG.DescripGasto,MontoGasto " & _
'        "FROM FondoGasto FG JOIN DetalleConceptoGasto DCG ON(DCG.CodDetalleGasto=FG.CodDetalleGasto AND DCG.CodGasto=FG.CodGasto) " & _
'        "JOIN ConceptoGasto CG ON(CG.CodGasto=DCG.CodGasto) " & _
'        "WHERE CodTipoCalculo='" & strCodTipoProvision & "' AND CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND FG.IndVigente='X'" & _
'        "ORDER BY NumGasto"
'    strSQL = "SELECT FG.CodCuenta,FG.NumGasto,FG.CodAnalitica,PCG.DescripCuenta,FG.DescripGasto,FG.MontoGasto,FG.CodFile,INP.DescripPersona as DescripProveedor " & _
'        "FROM FondoGasto FG JOIN FondoConceptoGasto FCG ON(FCG.CodCuenta=FG.CodCuenta AND FCG.CodAdministradora=FG.CodAdministradora AND FCG.CodFondo=FG.CodFondo) " & _
'        "JOIN PlanContable PCG ON(PCG.CodCuenta=FG.CodCuenta) " & _
'        "JOIN InstitucionPersona INP ON(INP.CodPersona=FG.CodProveedor AND INP.TipoPersona = '" & Codigo_Tipo_Persona_Proveedor & "') " & _
'        "WHERE CodFile='099' AND CodTipoCalculo='" & strCodTipoProvision & "' AND FG.CodFondo='" & strCodFondo & "' AND FG.CodAdministradora='" & gstrCodAdministradora & "' AND FG.IndVigente='X' " & _
'        " UNION " & _
'        "SELECT FG.CodCuenta,FG.NumGasto,FG.CodAnalitica,FG.DescripGasto DescripCuenta,DescripGasto,MontoGasto,CodFile,INP.DescripPersona as DescripProveedor " & _
'        "FROM FondoGasto FG " & _
'        "JOIN InstitucionPersona INP ON(INP.CodPersona=FG.CodProveedor AND INP.TipoPersona = '" & Codigo_Tipo_Persona_Proveedor & "') " & _
'        "WHERE CodFile='098' AND CodTipoCalculo='" & strCodTipoProvision & "' AND CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND FG.IndVigente='X'"
                        
    strSQL = "SELECT FG.CodCuenta,FG.NumGasto,FG.CodAnalitica,PCG.DescripCuenta,PRV.DescripParametro as TipoCalculo, FG.DescripGasto,FG.MontoGasto,FG.CodFile,INP.DescripPersona as DescripProveedor " & _
        "FROM FondoGasto FG JOIN FondoConceptoGasto FCG ON(FCG.CodCuenta=FG.CodCuenta AND FCG.CodAdministradora=FG.CodAdministradora AND FCG.CodFondo=FG.CodFondo) " & _
        "JOIN PlanContable PCG ON(PCG.CodCuenta=FG.CodCuenta) " & _
        "JOIN InstitucionPersona INP ON(INP.CodPersona=FG.CodProveedor AND INP.TipoPersona = '" & Codigo_Tipo_Persona_Proveedor & "') " & _
        "JOIN AuxiliarParametro PRV ON(PRV.CodParametro=FG.CodTipoGasto AND PRV.CodTipoParametro='TIPPAG') " & _
        "WHERE FG.CodFondo='" & gstrCodFondoContable & "' AND FG.CodAdministradora='" & gstrCodAdministradora & "'"
                        
    If chkVerSoloGastosVigentes.Value = vbChecked Then
        strSQL = strSQL & " AND (FG.IndVigente = 'X')"
    End If
    
    strSQL = strSQL & " ORDER BY FG.NumGasto"
                        
    strEstado = Reg_Defecto
    With adoConsulta
        .ConnectionString = gstrConnectConsulta
        .RecordSource = strSQL
        .Refresh
    End With
        
    tdgConsulta.Refresh
    
    If adoConsulta.Recordset.RecordCount > 0 Then strEstado = Reg_Consulta
            
End Sub

Private Sub CargarGastos()

    '*** Gastos del Fondo ***
'    strSQL = "SELECT (FCG.CodDetalleGasto + FCG.CodGasto + DCG.CodAnalitica) CODIGO,(RTRIM(CG.DescripConcepto) + '-' + RTRIM(DCG.DescripGasto)) DESCRIP " & _
'        "FROM FondoConceptoGasto FCG JOIN DetalleConceptoGasto DCG ON(DCG.CodDetalleGasto=FCG.CodDetalleGasto AND DCG.CodGasto=FCG.CodGasto) " & _
'        "JOIN ConceptoGasto CG ON (CG.CodGasto=DCG.CodGasto) " & _
'        "WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' " & _
'        "ORDER BY DCG.DescripGasto"
    strSQL = "SELECT FCG.CodCuenta CODIGO,(RTRIM(DescripCuenta)) DESCRIP " & _
        "FROM FondoConceptoGasto FCG JOIN PlanContable PCG ON(PCG.CodCuenta=FCG.CodCuenta AND PCG.CodAdministradora=FCG.CodAdministradora) " & _
        "WHERE CodFondo='" & strCodFondo & "' AND FCG.CodAdministradora='" & gstrCodAdministradora & "' " & _
        "ORDER BY DescripCuenta"
    CargarControlLista strSQL, cboGasto, arrGasto(), Sel_Defecto
    
End Sub

Public Sub Salir()

    Unload Me
    
End Sub



Private Sub cboAfectacion_Change()
    'JAFR 18/02/2011
    If cboAfectacion.ListIndex = 0 Then
        chkGeneraCreditoFiscal.Enabled = True
    Else
        chkGeneraCreditoFiscal.Enabled = False
        chkGeneraCreditoFiscal.Value = Not vbChecked
    End If
    'Fin JAFR
End Sub

Private Sub cboAfectacion_Click()
    strCodAfectacion = Valor_Caracter
    If cboAfectacion.ListIndex < 0 Then Exit Sub
    
    'JAFR 18/02/11
    If cboAfectacion.ListIndex = 1 Then
        chkGeneraCreditoFiscal.Enabled = True
    Else
        chkGeneraCreditoFiscal.Enabled = False
        chkGeneraCreditoFiscal.Value = 0
    End If
    'Fin JAFR
    strCodAfectacion = arrAfectacion(cboAfectacion.ListIndex)
End Sub

'Private Sub cboAplicacionDevengo_Click()
'
'    strCodAplicacionDevengo = Valor_Caracter
'    If cboAplicacionDevengo.ListIndex < 0 Then Exit Sub
'
'    strCodAplicacionDevengo = Trim(arrAplicacionDevengo(cboAplicacionDevengo.ListIndex))
'
'    If strCodAplicacionDevengo = Codigo_Aplica_Devengo_Inmediata Then
'        lblDescrip(22).Visible = False
'        cboFrecuenciaDevengo.Visible = False
'        cboFrecuenciaDevengo.ListIndex = -1
'    Else
'        lblDescrip(22).Visible = True
'        cboFrecuenciaDevengo.Visible = True
'        cboFrecuenciaDevengo.ListIndex = -1
'    End If
'
'End Sub

Private Sub cboBaseCalculo_Click()

    strCodBaseCalculo = Valor_Caracter
    If cboBaseCalculo.ListIndex < 0 Then Exit Sub
    
    strCodBaseCalculo = Trim(arrBaseCalculo(cboBaseCalculo.ListIndex))
    
End Sub

Private Sub cboEstado_Click()

    strEstadoGasto = Valor_Caracter
    If cboEstado.ListIndex < 0 Then Exit Sub
    
    strEstadoGasto = Trim(arrEstado(cboEstado.ListIndex))

End Sub

Private Sub cboFondoSerie_Click()
    strCodFondoSerie = ""
    If cboFondoSerie.ListIndex < 0 Then Exit Sub
    
    strCodFondoSerie = Trim(arrFondoSerie(cboFondoSerie.ListIndex))
    
    Call Buscar
End Sub



Private Sub cboFormula_Click()
    strCodFormula = Valor_Caracter
    If cboFormula.ListIndex < 0 Then Exit Sub
    
    strCodFormula = Trim(arrFormula(cboFormula.ListIndex))
End Sub

'Private Sub cboFrecuenciaGasto_Click()
'
'    strCodFrecuenciaGasto = Valor_Caracter
'    If cboFrecuenciaGasto.ListIndex < 0 Then Exit Sub
'
'    strCodFrecuenciaGasto = Trim(arrFrecuenciaGasto(cboFrecuenciaGasto.ListIndex))
'
'
'End Sub

Private Sub cboModalidadCalculo_Click()
'JCB R01
Dim intRegistro As Integer

    strCodModalidadCalculo = Valor_Caracter
    If cboModalidadCalculo.ListIndex < 0 Then Exit Sub
    
    strCodModalidadCalculo = Trim(arrModalidadCalculo(cboModalidadCalculo.ListIndex))

    If strCodModalidadCalculo = Codigo_Modalidad_Devengo_Inmediata Then
    
        intRegistro = ObtenerItemLista(arrTipoCalculo(), Codigo_Tipo_Gasto_Unico)
        If intRegistro >= 0 Then cboTipoCalculo.ListIndex = intRegistro
        cboTipoCalculo_Click
        cboTipoCalculo.Enabled = False
        
        intRegistro = ObtenerItemLista(arrFormaCalculo(), Codigo_Tipo_Devengo_Valor_Total)
        If intRegistro >= 0 Then cboFormaCalculo.ListIndex = intRegistro
        cboFormaCalculo_Click
        cboFormaCalculo.Enabled = False
        
        intRegistro = ObtenerItemLista(arrModalidadPago(), Codigo_Modalidad_Pago_Vencimiento)
        If intRegistro >= 0 Then cboModalidadPago.ListIndex = intRegistro
        cboModalidadPago.Enabled = False
     
        intRegistro = ObtenerItemLista(arrTipoPago(), Codigo_Tipo_Pago_Unico)
        If intRegistro >= 0 Then cboTipoPago.ListIndex = intRegistro
        cboTipoPago.Enabled = False

        txtValorFrecuenciaCalculo.Enabled = False
        cboFrecuenciaCalculo.Enabled = False

        Frame2.Enabled = False
    Else
        cboTipoCalculo.Enabled = True
        cboFormaCalculo.Enabled = True
        txtValorFrecuenciaCalculo.Enabled = True
        cboFrecuenciaCalculo.Enabled = True
        cboTipoPago.Enabled = True
         Frame2.Enabled = True
    End If
    
    If strCodModalidadCalculo = Codigo_Modalidad_Devengo_Ganancia_Diferida Then
        intRegistro = ObtenerItemLista(arrModalidadPago(), Codigo_Modalidad_Pago_Adelantado)
        If intRegistro >= 0 Then cboModalidadPago.ListIndex = intRegistro
        cboModalidadPago.Enabled = False
    Else
        intRegistro = ObtenerItemLista(arrModalidadPago(), Codigo_Modalidad_Pago_Vencimiento)
        If intRegistro >= 0 Then cboModalidadPago.ListIndex = intRegistro
        cboModalidadPago.Enabled = False
    End If
    
End Sub

Private Sub cboModalidadPago_Click()

    Dim intRegistro As Integer

    strCodModalidadPago = Valor_Caracter
    If cboModalidadPago.ListIndex < 0 Then Exit Sub
    
    strCodModalidadPago = Trim(arrModalidadPago(cboModalidadPago.ListIndex))

    If strCodModalidadPago = Codigo_Modalidad_Pago_Adelantado Then
        intRegistro = ObtenerItemLista(arrTipoPago(), Codigo_Tipo_Pago_Unico)
        If intRegistro >= 0 Then cboTipoPago.ListIndex = intRegistro
        cboTipoPago.Enabled = False
    Else
        cboTipoPago.ListIndex = 1
        cboPeriodoPago.ListIndex = 6
        cboTipoPago.Enabled = True
    End If



End Sub

Private Sub cboCreditoFiscal_Click()

    strCodCreditoFiscal = Valor_Caracter
    If cboCreditoFiscal.ListIndex < 0 Then Exit Sub
    
    strCodCreditoFiscal = arrCreditoFiscal(cboCreditoFiscal.ListIndex)
    
End Sub

Private Sub cboFondo_Click()

    Dim adoRegistro     As ADODB.Recordset
    
    strCodFondo = Valor_Caracter
    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(arrFondo(cboFondo.ListIndex))
    
    'Cargamos las series del fondo
'    strSQL = "{ call up_ACSelDatosParametro(50,'" & gstrCodAdministradora & "','" & strCodFondo & "') }"
'    CargarControlLista strSQL, cboFondoSerie, arrFondoSerie(), Valor_Caracter
'
'    If cboFondoSerie.ListCount > 0 Then cboFondoSerie.ListIndex = 0
'
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Fecha Vigente, Valor Cuota, Valor Cuota T-1, Moneda y Cantidad Inicial de Cuotas del Fondo ***
        .CommandText = "{ call up_ACSelDatosParametro(23,'" & strCodFondo & "','" & gstrCodAdministradora & "','000') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            dtpFechaGasto.Value = adoRegistro("FechaCuota")
            strCodMoneda = adoRegistro("CodMoneda")
            
            gdatFechaActual = adoRegistro("FechaCuota")
            frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    Call Buscar
    
End Sub


'Private Sub cboFrecuenciaDevengo_Click()
'
'    strCodFrecuenciaDevengo = Valor_Caracter
'    If cboFrecuenciaDevengo.ListIndex < 0 Then Exit Sub
'
'    strCodFrecuenciaDevengo = Trim(arrFrecuenciaDevengo(cboFrecuenciaDevengo.ListIndex))
'
'End Sub


Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Listado de Gastos del Fondo Vigentes"
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Text = "Listado de Gastos del Fondo No Vigentes"
'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo3").Visible = True
'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo3").Text = "Resumen de Gastos del Fondo"
    
End Sub

Private Sub cboFrecuenciaCalculo_Click()

    strCodFrecuenciaCalculo = Valor_Caracter
    If cboFrecuenciaCalculo.ListIndex < 0 Then Exit Sub
    
    strCodFrecuenciaCalculo = Trim(arrFrecuenciaCalculo(cboFrecuenciaCalculo.ListIndex))

End Sub


Private Sub cboPeriodoCalculo_Click()
    strCodPeriodoCalculo = Valor_Caracter
    If cboPeriodoCalculo.ListIndex < 0 Then Exit Sub
    strCodPeriodoCalculo = Trim(arrPeriodoCalculo(cboPeriodoCalculo.ListIndex))
End Sub

Private Sub cboPeriodoPago_Click()

    strCodPeriodoPago = Valor_Caracter
    If cboPeriodoPago.ListIndex < 0 Then Exit Sub
    
    strCodPeriodoPago = Trim(arrPeriodoPago(cboPeriodoPago.ListIndex))
    
End Sub



Private Sub cboGasto_Click()
    
    Dim adoAuxiliar As ADODB.Recordset

    strCodGasto = Valor_Caracter ': strCodAnalitica = Valor_Caracter
    If cboGasto.ListIndex <= 0 Then Exit Sub
    
    strCodGasto = Trim(arrGasto(cboGasto.ListIndex))
    
'    With adoComm
'        .CommandText = "SELECT * FROM DinamicaContable " & _
'            "WHERE CodFile = '099' AND CodAdministradora = '" & gstrCodAdministradora & "' AND CodMoneda = '" & strCodMoneda & _
'            "' and TipoOperacion = '19' and TipoCuentaInversion in ('" & Codigo_CtaProvGasto & "','" & Codigo_CtaComision & "') AND CodCuenta = '" & strCodGasto & "'"
'
'        Set adoAuxiliar = .Execute
'
'        If adoAuxiliar.EOF Then
'            MsgBox "No existe dinámica contable para el gasto seleccionado", vbExclamation
'           ' cboGasto.ListIndex = 0
'        End If
'
'    End With

    Dim adoRegistro     As ADODB.Recordset

    Set adoRegistro = New ADODB.Recordset

'    With adoComm
'        .CommandText = "SELECT CodDetalleFile FROM InversionDetalleFile " & _
'            "WHERE CodFile='" & strCodFile & "' AND DescripDetalleFile='" & strCodGasto & "'"
'        Set adoRegistro = .Execute
'
'        If Not adoRegistro.EOF Then
'            strCodAnalitica = Format(adoRegistro("CodDetalleFile"), "00000000")
'        End If
'        adoRegistro.Close: Set adoRegistro = Nothing
'    End With

    If strEstado = Reg_Adicion Then
        lblAnalitica.Caption = "099-" & strCodAnalitica
    Else
        lblAnalitica.Caption = Trim(tdgConsulta.Columns(8).Value) & "-" & strCodAnalitica
    End If
    
End Sub


Private Sub cboMoneda_Click()

    strCodMoneda = Valor_Caracter
    If cboMoneda.ListIndex < 0 Then Exit Sub
    
    strCodMoneda = Trim(arrMoneda(cboMoneda.ListIndex))
    
    lblMoneda.Caption = ObtenerCodSignoMoneda(strCodMoneda)
    
    lblMagnitud.Caption = lblMoneda.Caption

    
End Sub




Private Sub cboTipoGasto_Click()

    strCodTipoGasto = Valor_Caracter
    
    If cboTipoGasto.ListIndex < 0 Then Exit Sub
    
    strCodTipoGasto = Trim(arrTipoGasto(cboTipoGasto.ListIndex))
    
    If strCodTipoGasto = Tipo_Calculo_Variable Then
        lblFormula.Visible = True
        cboFormula.Visible = True
        txtMontoGasto.Enabled = False
        txtMontoGasto.Text = "0"
    Else
        lblFormula.Visible = False
        cboFormula.Visible = False
        cboFormula.ListIndex = -1
        txtMontoGasto.Enabled = True
    End If

End Sub

Private Sub cboTipoDesplazamiento_Click()

    strCodTipoDesplazamiento = Valor_Caracter
    If cboTipoDesplazamiento.ListIndex < 0 Then Exit Sub
    
    strCodTipoDesplazamiento = Trim(arrTipoDesplazamiento(cboTipoDesplazamiento.ListIndex))

End Sub

Private Sub cboFormaCalculo_Click()

    strCodFormaCalculo = Valor_Caracter
    If cboFormaCalculo.ListIndex < 0 Then Exit Sub
    
    strCodFormaCalculo = Trim(arrFormaCalculo(cboFormaCalculo.ListIndex))

    If strCodFormaCalculo = Codigo_Tipo_Devengo_Alicuota_Lineal Or Codigo_Tipo_Devengo_Alicuota_Incremental Then
        txtValorFrecuenciaCalculo.Text = "1"
        txtValorFrecuenciaCalculo.Enabled = True
        cboFrecuenciaCalculo.Enabled = True
        cboBaseCalculo.Enabled = True
    Else
        txtValorFrecuenciaCalculo.Text = "0"
        txtValorFrecuenciaCalculo.Enabled = False
        cboFrecuenciaCalculo.Enabled = False
        cboFrecuenciaCalculo.ListIndex = -1
        cboBaseCalculo.Enabled = False
        cboBaseCalculo.ListIndex = -1
    End If
  
    
End Sub

Private Sub cboTipoCalculo_Click()
    Dim intRegistro As Integer
    
    strCodTipoCalculo = Valor_Caracter
    If cboTipoCalculo.ListIndex < 0 Then Exit Sub
    
    strCodTipoCalculo = Trim(arrTipoCalculo(cboTipoCalculo.ListIndex))

    If strCodTipoCalculo = Codigo_Tipo_Pago_Unico Then
       ' lblDescrip(5).Visible = False
        txtValorPeriodoCalculo.Enabled = False
        cboPeriodoCalculo.Enabled = False
    Else
        lblDescrip(5).Visible = True
        txtValorPeriodoCalculo.Enabled = True
        cboPeriodoCalculo.Enabled = True
    End If

End Sub

Private Sub cboTipoPago_Click()

    Dim intRegistro As Integer
    
    strCodTipoPago = Valor_Caracter
    If cboTipoPago.ListIndex < 0 Then Exit Sub
    
    strCodTipoPago = Trim(arrTipoPago(cboTipoPago.ListIndex))
    
    If strCodTipoPago = Codigo_Tipo_Pago_Unico Then
        lblDescrip(28).Visible = False
        txtValorPeriodoPago.Text = 0
        txtValorPeriodoPago.Visible = False
        cboPeriodoPago.ListIndex = 6
        cboPeriodoPago.Visible = False
        
        intRegistro = ObtenerItemLista(arrTipoDesplazamiento(), Tipo_Desplazamiento_Ningun_Desplazamiento)
        If intRegistro >= 0 Then cboTipoDesplazamiento.ListIndex = intRegistro
        cboTipoDesplazamiento.Enabled = False

    Else
        lblDescrip(28).Visible = True
        txtValorPeriodoPago.Visible = True
        cboPeriodoPago.Visible = True
        cboTipoDesplazamiento.ListIndex = 0
        cboTipoDesplazamiento.Enabled = True
    End If
    
    
End Sub

Private Sub cboTipoProveedor_Click()
    strCodTipoProveedor = Trim(arrTipoProveedor(cboTipoProveedor.ListIndex))
    lblCodProveedor.Caption = Valor_Caracter
    lblProveedor.Caption = Valor_Caracter
End Sub

Private Sub chkIndicadorPagoFinMes_Click()
    If chkIndicadorPagoFinMes.Value = vbChecked Then
        dtpFechaPrimerPago.Enabled = False
    Else
        dtpFechaPrimerPago.Enabled = True
    End If
End Sub

Private Sub chkVerSoloGastosVigentes_Click()

     Call Buscar

End Sub

Private Sub cmdProveedor_Click()

    Dim sSql As String
   
    Screen.MousePointer = vbHourglass
   
    Dim frmBus As frmBuscar
    
    Set frmBus = New frmBuscar
    
    With frmBus.TBuscarRegistro1
           
        .ADOConexion = adoConn
        .ADOConexion.CommandTimeout = 0
        'If Index <> 2 Then
        '    .iTipoGrilla = 1
        'Else
        '    .iTipoGrilla = 2
        .iTipoGrilla = 2
        
        frmBus.Caption = " Relación de Proveedores"
        If strCodTipoProveedor = Codigo_Tipo_Persona_Proveedor Then
            .sSql = "{ call up_ACSelDatos(26) }"
        ElseIf strCodTipoProveedor = Codigo_Tipo_Persona_Comisionista Then
            .sSql = "{ call up_ACSelDatos(60) }"
        ElseIf strCodTipoProveedor = Codigo_Tipo_Persona_Emisor Then
            .sSql = "{ call up_ACSelDatos(22) }"
        End If
        
        
        .OutputColumns = "1,2,3,4,5,6"
        .HiddenColumns = "1,2,6"
        
        .BuscarTabla
        
        Screen.MousePointer = vbNormal
        frmBus.Show 1
       
        If .iParams.Count = 0 Then Exit Sub
        
        If .iParams(1).Valor <> "" Then
            lblProveedor.Caption = .iParams(5).Valor
            lblTipoDocID.Caption = .iParams(3).Valor
            lblNumDocID.Caption = .iParams(4).Valor
            'lblDireccion.Caption = .iParams(6).Valor
            lblCodProveedor.Caption = .iParams(1).Valor
            strCodAnalitica = lblCodProveedor.Caption
            lblAnalitica.Caption = "099-" & lblCodProveedor.Caption
        
        Else
            strCodAnalitica = ""
            lblAnalitica.Caption = "099-????????"
        End If
            
       
    End With
    
    Set frmBus = Nothing


End Sub

Private Sub dtpFechaFin_Change()

    If dtpFechaFin.Value < dtpFechaInicio.Value Then
        dtpFechaFin.Value = dtpFechaInicio.Value
    End If
                        
End Sub


Private Sub dtpFechaInicio_Change()

    If dtpFechaInicio.Value > dtpFechaFin.Value Then
        dtpFechaInicio.Value = dtpFechaFin.Value
    End If
    
    
End Sub
'JAFR 07/03/2011
Private Sub dtpFechaPrimerPago_Change()
    If dtpFechaPrimerPago.Value < fechaLimiteInferior Then
        dtpFechaPrimerPago.Value = fechaLimiteInferior
    End If
End Sub
'FIN JAFR

Private Sub Form_Activate()

    Call CargarReportes
    
End Sub

Private Sub Form_Deactivate()

    Call OcultarReportes
    
End Sub


Private Sub Form_Load()


    Call ValidarPermisoUsoControl(Trim(gstrLogin), Me, Trim(App.Title) + Separador_Codigo_Objeto + _
    gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)
    
    Call InicializarValores
    Call CargarListas
    Call Buscar
      
    CentrarForm Me
        
    Call ValidaExisteTipoCambio(Codigo_TipoCambio_SBS, gstrFechaActual)
 

End Sub

Private Sub CargarListas()
        
    Dim intRegistro         As Integer
    
    '*** Fondos ***
    '    strSQL = "{ call up_ACSelDatosParametro(29,'" & gstrCodAdministradora & "') }"
'    CargarControlLista strSQL, cboFondo, arrFondo(), Valor_Caracter
    strSQL = "SELECT CodFondo CODIGO,DescripFondo DESCRIP FROM Fondo WHERE CodAdministradora='" & gstrCodAdministradora & "' AND Estado='01' and CodFondo = '" & gstrCodFondoContable & "' ORDER BY DescripFondo"
    CargarControlLista strSQL, cboFondo, arrFondo(), Valor_Caracter
        

    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
   
'    '*** Tipo de Valor ***
'    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='VALCOM' ORDER BY DescripParametro"
'    CargarControlLista strSQL, cboTipoValor, arrTipoValor(), Valor_Caracter
    
    '*** Tipo de Gasto ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPPAG' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboTipoCalculo, arrTipoCalculo(), Valor_Caracter
    
    '*** Tipo de Desplazamiento ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPDES' ORDER BY CodParametro"
    CargarControlLista strSQL, cboTipoDesplazamiento, arrTipoDesplazamiento(), Valor_Caracter
    
    '*** Tipo de Desplazamiento ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPCAC' ORDER BY CodParametro"
    CargarControlLista strSQL, cboTipoGasto, arrTipoGasto(), Valor_Caracter
    
    '*** Moneda ***
    strSQL = "{ call up_ACSelDatos(2) }"
    CargarControlLista strSQL, cboMoneda, arrMoneda(), Valor_Caracter
            
'    '*** Tipos de Frecuencias ***
'    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPFRE' ORDER BY CodParametro"
'    CargarControlLista strSQL, cboFrecuenciaDevengo, arrFrecuenciaDevengo(), Valor_Caracter
'
'    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPFRE' ORDER BY CodParametro"
'    CargarControlLista strSQL, cboFrecuenciaGasto, arrFrecuenciaGasto(), Valor_Caracter
    
'    '*** Tipos de Frecuencias ***
'    strSQL = "{ call up_ACSelDatos(17) }"
'    CargarControlLista strSQL, cboPeriodoTasa, arrPeriodoTasa(), Valor_Caracter
   
    '*** Base de Cálculo ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='BASANU' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboBaseCalculo, arrBaseCalculo(), Valor_Caracter
    
'    '*** Tipo de Periodos ***
'    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='NATTAS' ORDER BY CodParametro"
'    CargarControlLista strSQL, cboTipoTasa, arrTipoTasa(), Valor_Caracter
      
    '*** Tipo de Periodos ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPPRD' ORDER BY CodParametro"
    CargarControlLista strSQL, cboPeriodoPago, arrPeriodoPago(), Valor_Caracter
    
    '*** Tipo de Proveedor ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPPER' and CodParametro in ('" & Codigo_Tipo_Persona_Emisor & "','" & Codigo_Tipo_Persona_Proveedor & "','" & Codigo_Tipo_Persona_Comisionista & "') ORDER BY CodParametro"
    CargarControlLista strSQL, cboTipoProveedor, arrTipoProveedor(), Valor_Caracter
    cboTipoProveedor.ListIndex = 1
    
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPPRD' ORDER BY CodParametro"
    CargarControlLista strSQL, cboFrecuenciaCalculo, arrFrecuenciaCalculo(), Valor_Caracter
        
    strSQL = "{ call up_ACSelDatos(34) }"
    CargarControlLista strSQL, cboPeriodoCalculo, arrPeriodoCalculo(), Valor_Caracter
  
    '*** Tipo de Pago ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPPAG' ORDER BY CodParametro"
    CargarControlLista strSQL, cboTipoPago, arrTipoPago(), Valor_Caracter
        
    '*** Modalidad de Pago ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='MODPAG' ORDER BY CodParametro"
    CargarControlLista strSQL, cboModalidadPago, arrModalidadPago(), Valor_Caracter
    
    '*** Afectacion ***
    strSQL = "SELECT CodParametro CODIGO, DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro ='AFEIMP' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboAfectacion, arrAfectacion(), Sel_Defecto
    
    '*** Tipo Crédito Fiscal ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP From AuxiliarParametro WHERE CodTipoParametro='CREFIS' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboCreditoFiscal, arrCreditoFiscal(), Valor_Caracter
        
    '*** Estados ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP From AuxiliarParametro WHERE CodTipoParametro='INDREG' AND CodParametro<>'03' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboEstado, arrEstado(), Valor_Caracter
        
    '*** Tipo de Devengo ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP From AuxiliarParametro WHERE CodTipoParametro='MODDEV' and Estado = '01' ORDER BY CodParametro"
    CargarControlLista strSQL, cboFormaCalculo, arrFormaCalculo(), Valor_Caracter
        
'    '*** Modalidad de Devengo ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP From AuxiliarParametro WHERE CodTipoParametro='TIPDEV' ORDER BY CodParametro"
    CargarControlLista strSQL, cboModalidadCalculo, arrModalidadCalculo(), Valor_Caracter
        
    '*** Tipo de desplazamiento
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP From AuxiliarParametro WHERE CodTipoParametro='TIPDES' ORDER BY CodParametro"
    CargarControlLista strSQL, cboTipoDesplazamiento, arrTipoDesplazamiento(), Valor_Caracter
    
    '*** Formulas
    strSQL = "SELECT CodFormula CODIGO,DescripFormula DESCRIP From Formula ORDER BY DescripFormula"
    CargarControlLista strSQL, cboFormula, arrFormula(), Valor_Caracter
    
    
        
End Sub
Private Sub InicializarValores()
                        
    '*** Valores Iniciales ***
    tabGasto.Tab = 0
    strCodFile = "099"
    
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgConsulta.Columns(2).Width = tdgConsulta.Width * 0.01 * 6
    tdgConsulta.Columns(3).Width = tdgConsulta.Width * 0.01 * 10
    tdgConsulta.Columns(4).Width = tdgConsulta.Width * 0.01 * 15
    tdgConsulta.Columns(5).Width = tdgConsulta.Width * 0.01 * 30
    
    chkVerSoloGastosVigentes.Value = vbChecked
    
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    Set cmdOpcion.FormularioActivo = Me
    
End Sub
Private Sub Form_Unload(Cancel As Integer)

    Call OcultarReportes
    Set frmFondoGastos = Nothing
    
End Sub

Public Sub Accion(nAccion As ButtonAction)
    
    Select Case nAccion
        
        Case vNew
            Call Adicionar
        Case vQuery
            Call Modificar
        Case vDelete
            Call Eliminar
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

Public Sub Cancelar()

    cmdOpcion.Visible = True
    With tabGasto
        .TabEnabled(0) = True
        .Tab = 0
    End With
    strEstado = Reg_Consulta
    
End Sub

Public Sub Grabar()

    Dim adoRegistro                     As ADODB.Recordset, adoRec      As ADODB.Recordset
    Dim adoConsulta                     As ADODB.Recordset
    Dim intAccion                       As Integer, lngNumError         As Long
    Dim dblTipCambio                    As Double
    Dim strFechaAnterior                As String, strFechaSiguiente    As String
    Dim datFechaFinPeriodo              As Date
    Dim numMontoGasto                   As Double
    Dim numPorcenGasto                  As Double
    Dim numPorcenGastoAnual             As Double
    Dim strIndNoIncluyeBalancePreCierre As String
    Dim strIndFinPeriodo                    As String
    Dim strIndSinVencimiento            As String
    Dim intDiasProvision                As String
    Dim intDiasBaseAnual                As Integer
    Dim intNumPeriodoAnualTasa          As Integer
    
    
    Dim montoOrdenPago As Double
    Dim adoAuxConsulta As ADODB.Recordset
    Dim strFechaPago As String
    Dim strFechaVencimiento As String
            
    cboAfectacion.ListIndex = 1
    cboCreditoFiscal.ListIndex = 1
            
    If strEstado = Reg_Consulta Then Exit Sub
    If Not TodoOK() Then Exit Sub
    
    'On Error GoTo CtrlError
    
'    If strCodTipoValor = Codigo_Tipo_Costo_Porcentaje Then
'        numPorcenGasto = CDbl(txtMontoGasto.Text)
'        numMontoGasto = CalculoInteres(numPorcenGasto, strCodTipoTasa, strCodPeriodoTasa, strCodBaseCalculo, CDbl(txtBaseCalculo.Text), dtpFechaInicio.Value, dtpFechaFin.Value)
'    Else
    numPorcenGasto = 0
    numMontoGasto = CDbl(txtMontoGasto.Text)
'    End If
    
    If chkNoIncluyeEnBalancePrecierre.Value = vbChecked Then
        strIndNoIncluyeBalancePreCierre = Valor_Indicador
    Else
        strIndNoIncluyeBalancePreCierre = Valor_Caracter
    End If
    
    If chkIndicadorPagoFinMes.Value = vbChecked Then
        strIndFinPeriodo = Valor_Indicador
    Else
        strIndFinPeriodo = Valor_Caracter
    End If
    
    If chkIndicadorGastoIterativo.Value = vbChecked Then
        strIndGastoIterativo = "X"
    Else
        strIndGastoIterativo = ""
    End If
    
    'JAFR 18/02/2011
    If chkGeneraCreditoFiscal.Value = vbChecked Then
        strIndGeneraCreditoFiscal = Valor_Indicador
    Else
        strIndGeneraCreditoFiscal = Valor_Caracter
    End If
    'fin JAFR
    
'    If chkSinVencimiento.Value = vbChecked Then
'        strIndSinVencimiento = Valor_Indicador
'    Else
'        strIndSinVencimiento = Valor_Caracter
'    End If
    
    'dtpFechaPrimerPago.Value = dtpFechaFin.Value
    
    If strEstado = Reg_Adicion Then
        Set adoRegistro = New ADODB.Recordset
                                        
        Me.MousePointer = vbHourglass
        
        intSecuencialGasto = 0
        
        '*** Guardar ***
        With adoComm
        
            If txtValorFrecuenciaCalculo.Text = Valor_Caracter Then
                txtValorFrecuenciaCalculo.Text = "0"
            End If
            If txtValorPeriodoPago.Text = Valor_Caracter Then
                txtValorPeriodoPago.Text = "0"
            End If
        
            .CommandText = "{ call up_GNManFondoGasto('" & strCodFondo & "','" & _
                gstrCodAdministradora & "'," & intSecuencialGasto & ",'" & _
                strCodFondoSerie & "','" & Convertyyyymmdd(dtpFechaGasto.Value) & "','" & strCodGasto & "','" & _
                strCodFile & "','" & strCodAnalitica & "','" & strCodTipoProveedor & "','" & _
                Trim(lblCodProveedor.Caption) & "','" & Trim(txtDescripGasto.Text) & "','" & _
                Convertyyyymmdd(CVDate(Valor_Fecha)) & "','','" & strEstadoGasto & "'," & _
                gdblTipoCambio & ",'" & strCodMoneda & "','" & strCodFormula & "'," & _
                numMontoGasto & ",'" & strCodTipoGasto & "','" & _
                Convertyyyymmdd(dtpFechaInicio.Value) & "','" & Convertyyyymmdd(dtpFechaFin.Value) & "','" & _
                strIndGastoIterativo & "','" & strCodModalidadCalculo & "','" & strCodFormaCalculo & "','" & _
                strCodBaseCalculo & "','" & strCodTipoCalculo & "'," & CLng(txtValorPeriodoCalculo.Text) & ",'" & _
                strCodPeriodoCalculo & "'," & CLng(txtValorFrecuenciaCalculo.Text) & ",'" & strCodFrecuenciaCalculo & "','" & _
                Convertyyyymmdd(dtpFechaPrimerCalculo.Value) & "','" & strCodModalidadPago & "','" & strCodTipoPago & "'," & _
                CLng(txtValorPeriodoPago.Text) & ",'" & strCodPeriodoPago & "','" & Convertyyyymmdd(dtpFechaPrimerPago.Value) & "','" & _
                strCodTipoDesplazamiento & "','" & strIndFinPeriodo & "','" & strCodAfectacion & "','" & _
                strCodCreditoFiscal & "','" & strIndGeneraCreditoFiscal & "','" & strIndNoIncluyeBalancePreCierre & "','I') }"
           
            adoConn.Execute .CommandText
            
            
            Call GenerarPeriodos

        End With
                                                                                                            
        Me.MousePointer = vbDefault
                    
        MsgBox Mensaje_Adicion_Exitosa, vbExclamation
        
        If strCodTipoProvision = Codigo_Gasto_MismoDia Then
            MsgBox "Para que este gasto genere la orden debe ser ingresado al Registro de Compras !", vbInformation, Me.Caption
        End If
        
        frmMainMdi.stbMdi.Panels(3).Text = "Acción"
        
        cmdOpcion.Visible = True
        With tabGasto
            .TabEnabled(0) = True
            .Tab = 0
        End With
        
        'adoRegistro.Close: Set adoRegistro = Nothing
    End If
    
    If strEstado = Reg_Edicion Then
        Me.MousePointer = vbHourglass
        
        'intSecuencialGasto = CInt(tdgConsulta.Columns(2).Value)
        
        '*** Actualizar ***
        With adoComm
            
            
            .CommandText = "{ call up_GNManFondoGasto('" & strCodFondo & "','" & _
                gstrCodAdministradora & "'," & intSecuencialGasto & ",'" & _
                strCodFondoSerie & "','" & Convertyyyymmdd(dtpFechaGasto.Value) & "','" & strCodGasto & "','" & _
                strCodFile & "','" & strCodAnalitica & "','" & strCodTipoProveedor & "','" & _
                Trim(lblCodProveedor.Caption) & "','" & Trim(txtDescripGasto.Text) & "','" & _
                Convertyyyymmdd(CVDate(Valor_Fecha)) & "','','" & strEstadoGasto & "'," & _
                gdblTipoCambio & ",'" & strCodMoneda & "','" & strCodFormula & "'," & _
                numMontoGasto & ",'" & strCodTipoGasto & "','" & _
                Convertyyyymmdd(dtpFechaInicio.Value) & "','" & Convertyyyymmdd(dtpFechaFin.Value) & "','" & _
                strIndGastoIterativo & "','" & strCodModalidadCalculo & "','" & strCodFormaCalculo & "','" & _
                strCodBaseCalculo & "','" & strCodTipoCalculo & "'," & CLng(txtValorPeriodoCalculo.Text) & ",'" & _
                strCodPeriodoCalculo & "'," & CLng(txtValorFrecuenciaCalculo.Text) & ",'" & strCodFrecuenciaCalculo & "','" & _
                Convertyyyymmdd(dtpFechaPrimerCalculo.Value) & "','" & strCodModalidadPago & "','" & strCodTipoPago & "'," & _
                CLng(txtValorPeriodoPago.Text) & ",'" & strCodPeriodoPago & "','" & Convertyyyymmdd(dtpFechaPrimerPago.Value) & "','" & _
                strCodTipoDesplazamiento & "','" & strIndFinPeriodo & "','" & strCodAfectacion & "','" & _
                strCodCreditoFiscal & "','" & strIndGeneraCreditoFiscal & "','" & strIndNoIncluyeBalancePreCierre & "','U') }"
                
            adoConn.Execute .CommandText
            
            Call GenerarPeriodos
            
        End With

        Me.MousePointer = vbDefault
                    
        MsgBox Mensaje_Edicion_Exitosa, vbExclamation
        
        frmMainMdi.stbMdi.Panels(3).Text = "Acción"
        
        cmdOpcion.Visible = True
        With tabGasto
            .TabEnabled(0) = True
            .Tab = 0
        End With
        
    End If
    
    Call cboFondo_Click
    Call Buscar
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

Public Sub SubImprimir(Index As Integer)

    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    
    gstrNameRepo = "FondoGasto"
                        
    Set frmReporte = New frmVisorReporte

    ReDim aReportParamS(2)
    ReDim aReportParamFn(4)
    ReDim aReportParamF(4)

    aReportParamFn(0) = "Usuario"
    aReportParamFn(1) = "Hora"
    aReportParamFn(2) = "NombreEmpresa"
    aReportParamFn(3) = "Fondo"
    aReportParamFn(4) = "Titulo"
    
    aReportParamF(0) = gstrLogin
    aReportParamF(1) = Format(Time(), "hh:mm:ss")
    aReportParamF(2) = gstrNombreEmpresa & Space(1)
    aReportParamF(3) = Trim(cboFondo.Text)
                
    aReportParamS(0) = strCodFondo
    aReportParamS(1) = gstrCodAdministradora
    
    Select Case Index
        Case 1
            aReportParamF(4) = "LISTADO DE GASTOS VIGENTES"
            aReportParamS(2) = Valor_Indicador
        Case 2
            aReportParamF(4) = "LISTADO DE GASTOS NO VIGENTES"
            aReportParamS(2) = Valor_Caracter
        Case 3
            aReportParamF(4) = "RESUMEN DE GASTOS"
            aReportParamS(2) = Valor_Indicador
            gstrNameRepo = "FondoGastoXPeriodo"
            
    End Select

    gstrSelFrml = Valor_Caracter
    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"

    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())

    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
    frmReporte.Show vbModal

    Set frmReporte = Nothing

    Screen.MousePointer = vbNormal
    
End Sub
Public Sub Imprimir()

    Call SubImprimir(1)
    
End Sub
Public Sub Eliminar()
    'If strEstado = Reg_Consulta Or strEstado = Reg_Edicion Then
        'If MsgBox("Se procederá a eliminar el Gasto del Fondo." & vbNewLine & vbNewLine & "Seguro de continuar ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
    
           ' adoComm.CommandText = "UPDATE FondoGasto SET IndVigente='" & Valor_Caracter & "' " & _
               ' "WHERE NumGasto=" & CInt(tdgConsulta.Columns(2)) & " AND CodGasto='" & tdgConsulta.Columns(1) & "' AND CodTipoCalculo='" & strCodTipoProvision & "' AND " & _
                '"CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
            'adoConn.Execute adoComm.CommandText
            
            'tabGasto.TabEnabled(0) = True
            'tabGasto.Tab = 0
            'Call Buscar
            
            'Exit Sub
        'End If
    'End If
    If strEstado = Reg_Consulta Or strEstado = Reg_Edicion Then
        If MsgBox("Se procederá a eliminar el Gasto del Fondo." & vbNewLine & vbNewLine & "Seguro de continuar ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            adoComm.CommandText = "UPDATE FondoGasto SET IndVigente='" & Valor_Caracter & "' " & _
            "WHERE NumGasto=" & CInt(tdgConsulta.Columns(2)) & " AND CodCuenta='" & tdgConsulta.Columns(1) & "' AND CodTipoCalculo='" & strCodTipoProvision & "' AND " & _
            "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
            adoConn.Execute adoComm.CommandText
            tabGasto.TabEnabled(0) = True
            tabGasto.Tab = 0
            Call Buscar
            Exit Sub
        End If
    End If
End Sub
Public Sub Modificar()

    If strEstado = Reg_Consulta Then
        strEstado = Reg_Edicion
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabGasto
            .TabEnabled(0) = False
            .Tab = 1
        End With
    End If
        
End Sub

Private Sub LlenarFormulario(strModo As String)

    Dim adoRegistro   As ADODB.Recordset
    Dim intRegistro As Integer
    
    Select Case strModo
        
        Case Reg_Adicion
            
            fraGastos(1).Caption = "Fondo : " & Trim(cboFondo.Text) & Space(1) & "-" & Space(1) & "Serie : " & Trim(cboFondoSerie.Text) '& "Tipo Gasto : " & Trim(cboTipoProvision.Text)
            lblAnalitica.Caption = "099-????????"
            txtDescripGasto.Text = Valor_Caracter
            txtMontoGasto.Text = "0"
            lblSaldoProvision.Caption = "0"
            txtBaseCalculo.Text = "0"
            txtValorPeriodoPago.Text = "0"
            dtpFechaGasto.Value = gdatFechaActual
            dtpFechaGasto.Enabled = False
            dtpFechaInicio.Value = gdatFechaActual
            dtpFechaFin.Value = gdatFechaActual
            dtpFechaPrimerCalculo.Value = gdatFechaActual
            'JAFR 07/03/2011
            dtpFechaPrimerPago.Value = gdatFechaActual
            'fin JAFR
            dtpFechaInicio.Enabled = True
            dtpFechaFin.Enabled = True
            chkSinVencimiento.Value = vbUnchecked
                        
            lblProveedor.Caption = Valor_Caracter
            lblCodProveedor.Caption = Valor_Caracter
            lblTipoDocID.Caption = Valor_Caracter
            lblNumDocID.Caption = Valor_Caracter
            
            Call CargarGastos
            
            cboGasto.ListIndex = -1
            If cboGasto.ListCount > 0 Then cboGasto.ListIndex = 0
            
                                  
            'Gastos
            intRegistro = ObtenerItemLista(arrMoneda(), strCodMoneda)
            If intRegistro >= 0 Then cboMoneda.ListIndex = intRegistro

            intRegistro = ObtenerItemLista(arrEstado(), Valor_Indicador)
            If intRegistro >= 0 Then cboEstado.ListIndex = intRegistro
                                  
            
            'Devengo
            cboFormaCalculo.ListIndex = -1
            txtValorFrecuenciaCalculo.Text = "0"
            cboFrecuenciaCalculo.ListIndex = -1
            cboPeriodoCalculo.ListIndex = -1
            cboModalidadCalculo.ListIndex = -1
            cboBaseCalculo.ListIndex = -1
                        
            intRegistro = ObtenerItemLista(arrFormaCalculo(), Codigo_Tipo_Devengo_Alicuota_Lineal)
            If intRegistro >= 0 Then cboFormaCalculo.ListIndex = intRegistro
                        
            txtValorFrecuenciaCalculo.Text = "1"
            
            intRegistro = ObtenerItemLista(arrFrecuenciaCalculo(), Codigo_Frecuencia_Diaria)
            If intRegistro >= 0 Then cboFrecuenciaCalculo.ListIndex = intRegistro
            
            intRegistro = ObtenerItemLista(arrPeriodoCalculo(), Codigo_Frecuencia_Diaria)
            If intRegistro >= 0 Then cboPeriodoCalculo.ListIndex = intRegistro
         
                        
            'Pago
            cboModalidadPago.ListIndex = -1
            cboTipoPago.ListIndex = -1
            cboPeriodoPago.ListIndex = -1
            cboTipoDesplazamiento.ListIndex = -1
            
            'Condiciones Contables y Tributarias
            cboAfectacion.ListIndex = -1
            cboCreditoFiscal.ListIndex = -1
            chkNoIncluyeEnBalancePrecierre.Value = vbUnchecked

            cboAfectacion.ListIndex = -1
                        
            intRegistro = ObtenerItemLista(arrModalidadCalculo(), Codigo_Modalidad_Devengo_Inmediata)
            If intRegistro >= 0 Then cboModalidadCalculo.ListIndex = intRegistro
                        
                        
'            intRegistro = ObtenerItemLista(arrAplicacionDevengo(), Codigo_Aplica_Devengo_Inmediata)
'            If intRegistro >= 0 Then cboAplicacionDevengo.ListIndex = intRegistro
                        
            
            intRegistro = ObtenerItemLista(arrTipoCalculo(), Codigo_Tipo_Gasto_Unico)
            If intRegistro >= 0 Then cboTipoCalculo.ListIndex = intRegistro
            
            intRegistro = ObtenerItemLista(arrModalidadPago(), Codigo_Modalidad_Pago_Vencimiento)
            If intRegistro >= 0 Then cboModalidadPago.ListIndex = intRegistro
            
            intRegistro = ObtenerItemLista(arrTipoGasto(), Tipo_Calculo_Fijo)
            If intRegistro >= 0 Then cboTipoGasto.ListIndex = intRegistro
           
            intRegistro = ObtenerItemLista(arrTipoPago(), Codigo_Tipo_Pago_Unico)
            If intRegistro >= 0 Then cboTipoPago.ListIndex = intRegistro
            
            intRegistro = ObtenerItemLista(arrTipoDesplazamiento(), Tipo_Desplazamiento_Ningun_Desplazamiento)
            If intRegistro >= 0 Then cboTipoDesplazamiento.ListIndex = intRegistro
           
'            intRegistro = ObtenerItemLista(arrTipoValor(), Codigo_Tipo_Costo_Monto)
'            If intRegistro >= 0 Then cboTipoValor.ListIndex = intRegistro
           
            intRegistro = ObtenerItemLista(arrBaseCalculo(), Codigo_Base_30_360)
            If intRegistro >= 0 Then cboBaseCalculo.ListIndex = intRegistro
                      
            lblMagnitud.Caption = lblMoneda.Caption
            
            cboGasto.SetFocus
                        
        Case Reg_Edicion
        
            Call CargarGastos
            
            Set adoRegistro = New ADODB.Recordset
            
            If tdgConsulta.AllowRowSelect = True Then
                adoComm.CommandText = "SELECT FG.CodAfectacion, FG.CodModalidadCalculo, FG.CodBaseCalculo,FG.NumGasto,FG.CodProveedor,FG.IndNoIncluyeEnBalancePreCierre, FG.IndVigente,FG.CodTipoCalculo,FG.CodCuenta,FG.CodFile,FG.CodAnalitica,FG.DescripGasto,FG.CodCreditoFiscal,FG.CodMoneda,FG.NumFrecuenciaCalculo,FG.CodFrecuenciaCalculo,FG.CodPeriodoPago,FG.CodTipoPago," & _
                    "FG.MontoGasto,FG.FechaDefinicion,FG.FechaInicial,FG.FechaFinal,FG.FechaPrimerPago,FG.NumPeriodoPago,FG.CodTipoDesplazamiento,FG.IndFinPeriodo,FG.CodFormaCalculo,FG.CodModalidadPago,FG.FechaPrimerPago," & _
                    "AP.DescripParametro TipoIdentidad,INP.DescripPersona,INP.NumIdentidad,FG.CodTipoGasto, FG.CodFormula, FG.IndGastoIterativo, FG.IndGeneraCreditoFiscal " & _
                    "FROM FondoGasto FG " & _
                    "JOIN InstitucionPersona INP ON(INP.CodPersona=FG.CodProveedor AND INP.TipoPersona = '" & Codigo_Tipo_Persona_Proveedor & "') " & _
                    "JOIN AuxiliarParametro AP ON (AP.CodParametro = INP.TipoIdentidad AND AP.CodTipoParametro = 'TIPIDE') " & _
                    "WHERE NumGasto=" & CInt(tdgConsulta.Columns("NumGasto")) & " AND CodCuenta='" & tdgConsulta.Columns("CodCuenta") & "' AND " & _
                    "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
                Set adoRegistro = adoComm.Execute
            Else
                MsgBox "Por lo menos debe haber un gasto registrado", vbExclamation, gstrNombreSistema
                Exit Sub
            End If
            
            If Not adoRegistro.EOF Then
                fraGastos(1).Caption = "Fondo : " & Trim(cboFondo.Text) & Space(1) & "-" & Space(1) & "Serie : " & Trim(cboFondoSerie.Text) ' & "Tipo Gasto : " & Trim(cboTipoProvision.Text)
                
                intSecuencialGasto = adoRegistro("NumGasto")
                
                strCodFile = Trim(adoRegistro("CodFile"))
                strCodAnalitica = Trim(adoRegistro("CodAnalitica"))
                lblAnalitica.Caption = Trim(adoRegistro("CodFile")) & "-" & Trim(adoRegistro("CodAnalitica"))
                txtDescripGasto.Text = Trim(adoRegistro("DescripGasto"))
                txtValorPeriodoPago.Text = Trim(adoRegistro("NumPeriodoPago"))
                txtValorFrecuenciaCalculo.Text = Trim(adoRegistro("NumFrecuenciaCalculo"))
            
                intRegistro = ObtenerItemLista(arrGasto(), adoRegistro("CodCuenta"))
                If intRegistro >= 0 Then cboGasto.ListIndex = intRegistro
            
                strCodTipoCalculo = CStr(adoRegistro("CodTipoCalculo"))
                
                lblProveedor.Caption = Trim(adoRegistro("DescripPersona"))
                lblTipoDocID.Caption = adoRegistro("TipoIdentidad")
                lblNumDocID.Caption = adoRegistro("NumIdentidad")
                lblCodProveedor.Caption = adoRegistro("CodProveedor")
               
                dtpFechaGasto.Value = adoRegistro("FechaDefinicion")
                dtpFechaInicio.Value = adoRegistro("FechaInicial")
                dtpFechaFin.Value = adoRegistro("FechaFinal")
                                                         
               
                intRegistro = ObtenerItemLista(arrTipoCalculo(), adoRegistro("CodTipoCalculo"))
                If intRegistro >= 0 Then cboTipoCalculo.ListIndex = intRegistro
                                   
                intRegistro = ObtenerItemLista(arrEstado(), adoRegistro("IndVigente"))
                If intRegistro >= 0 Then cboEstado.ListIndex = intRegistro
                
                lblSaldoProvision.Caption = 0 'CStr(adoRegistro("MontoDevengo"))
                                   
                'Montos
                intRegistro = ObtenerItemLista(arrMoneda(), adoRegistro("CodMoneda"))
                If intRegistro >= 0 Then cboMoneda.ListIndex = intRegistro
                                   
                txtMontoGasto.Text = CStr(adoRegistro("MontoGasto"))
                                   
                intRegistro = ObtenerItemLista(arrBaseCalculo(), adoRegistro("CodBaseCalculo"))
                If intRegistro >= 0 Then cboBaseCalculo.ListIndex = intRegistro
                                   
                                   
                intRegistro = ObtenerItemLista(arrModalidadCalculo(), adoRegistro("CodModalidadCalculo"))
                If intRegistro >= 0 Then cboModalidadCalculo.ListIndex = intRegistro
                                  
                'Fechas
                intRegistro = ObtenerItemLista(arrModalidadPago(), adoRegistro("CodModalidadPago"))
                If intRegistro >= 0 Then cboModalidadPago.ListIndex = intRegistro
                                   
                intRegistro = ObtenerItemLista(arrTipoPago(), adoRegistro("CodTipoPago"))
                If intRegistro >= 0 Then cboTipoPago.ListIndex = intRegistro
                                   
                intRegistro = ObtenerItemLista(arrPeriodoPago(), adoRegistro("CodPeriodoPago"))
                If intRegistro >= 0 Then cboPeriodoPago.ListIndex = intRegistro
                                   
                intRegistro = ObtenerItemLista(arrTipoDesplazamiento(), adoRegistro("CodTipoDesplazamiento"))
                If intRegistro >= 0 Then cboTipoDesplazamiento.ListIndex = intRegistro
                                   
                dtpFechaPrimerPago.Value = adoRegistro("FechaPrimerPago")
                
                intRegistro = ObtenerItemLista(arrTipoGasto(), adoRegistro("CodTipoGasto"))
                If intRegistro >= 0 Then cboTipoGasto.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrFormula(), "" & adoRegistro("CodFormula"))
                If intRegistro >= 0 Then cboFormula.ListIndex = intRegistro
                                   
                'Condiciones Contables y Tributarias
                intRegistro = ObtenerItemLista(arrFormaCalculo(), adoRegistro("CodFormaCalculo"))
                If intRegistro >= 0 Then cboFormaCalculo.ListIndex = intRegistro
                                                                                      
                intRegistro = ObtenerItemLista(arrFrecuenciaCalculo(), adoRegistro("CodFrecuenciaCalculo"))
                If intRegistro >= 0 Then cboFrecuenciaCalculo.ListIndex = intRegistro
                            
                intRegistro = ObtenerItemLista(arrAfectacion(), adoRegistro("CodAfectacion"))
                If intRegistro >= 0 Then cboAfectacion.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrCreditoFiscal(), adoRegistro("CodCreditoFiscal"))
                If intRegistro >= 0 Then cboCreditoFiscal.ListIndex = intRegistro
                
                If adoRegistro("IndNoIncluyeEnBalancePreCierre") = Valor_Indicador Then
                    chkNoIncluyeEnBalancePrecierre.Value = vbChecked
                Else
                    chkNoIncluyeEnBalancePrecierre.Value = vbUnchecked
                End If
                
                If adoRegistro("IndGastoIterativo") = Valor_Indicador Then
                    chkIndicadorGastoIterativo.Value = vbChecked
                Else
                    chkIndicadorGastoIterativo.Value = vbUnchecked
                End If
                                
                If adoRegistro("IndFinPeriodo") = Valor_Indicador Then
                    chkIndicadorPagoFinMes.Value = vbChecked
                Else
                    chkIndicadorPagoFinMes.Value = vbUnchecked
                End If
                
                If adoRegistro("IndGeneraCreditoFiscal") = Valor_Indicador Then
                    chkGeneraCreditoFiscal.Value = vbChecked
                Else
                    chkGeneraCreditoFiscal.Value = vbUnchecked
                End If
                
                                             
            End If
            adoRegistro.Close: Set adoRegistro = Nothing
    
    End Select
    
End Sub
Public Sub Adicionar()
                
    If strCodFondo = Valor_Caracter Then
        MsgBox "No existen fondos definidos...", vbCritical, Me.Caption
        Exit Sub
    End If
    
    frmMainMdi.stbMdi.Panels(3).Text = "Adicionar Gastos del Fondo..."
                
    strEstado = Reg_Adicion
    LlenarFormulario strEstado
    cmdOpcion.Visible = False
    With tabGasto
        .TabEnabled(0) = False
        .Tab = 1
    End With
      
End Sub
Private Function TodoOK()

    TodoOK = False
    
    If Trim(strCodGasto) = Valor_Caracter Then
        MsgBox "Debe Seleccionar el Gasto.", vbCritical
        cboGasto.SetFocus
        Exit Function
    End If
      
'    If Trim(strCodCreditoFiscal) = Valor_Caracter Then
'        MsgBox "Debe Seleccionar el Tipo de Crédito Fiscal.", vbCritical
'        cboCreditoFiscal.SetFocus
'        Exit Function
'    End If
    
    If Trim(txtDescripGasto.Text) = Valor_Caracter Then
        MsgBox "Debe Ingresar la Descripción del Gasto.", vbCritical
        txtDescripGasto.SetFocus
        Exit Function
    End If
    
    If cboMoneda.ListIndex = -1 Then
        MsgBox "Debe Ingresar la Moneda del Gasto.", vbCritical
        cboMoneda.SetFocus
        Exit Function
    End If
                
    If CDec(txtMontoGasto.Text) < 0 Then
        MsgBox "El Valor del Gasto no Puede Ser Menor que 0.", vbCritical
        txtMontoGasto.SetFocus
        Exit Function
    End If
    
    If Trim(lblCodProveedor.Caption) = "" Then
        MsgBox "Debe Indicar el Proveedor del Gasto.", vbCritical
        txtMontoGasto.SetFocus
        Exit Function
    End If
    
'    If cboAplicacionDevengo.ListIndex = -1 Then
'        MsgBox "Debe Seleccionar el Tipo de Aplicación de Devengo del Gasto.", vbCritical
'        cboAplicacionDevengo.SetFocus
'        Exit Function
'    End If
    
'    If strCodAplicacionDevengo = Codigo_Aplica_Devengo_Periodica And cboFrecuenciaDevengo.ListIndex = -1 Then
'        MsgBox "Debe Seleccionar la Frecuencia de Aplicación de Devengo del Gasto.", vbCritical
'        cboFrecuenciaDevengo.SetFocus
'        Exit Function
'    End If

    If cboAfectacion.ListIndex <= 0 Then
        MsgBox "Debe indicar si el Gasto está afecto a Impuesto.", vbCritical
        cboAfectacion.SetFocus
        Exit Function
    End If
    
    '*** Si todo paso OK ***
    TodoOK = True
  
End Function


Private Sub lblSaldoProvision_Change()

    Call FormatoMillarEtiqueta(lblSaldoProvision, Decimales_Monto)
    
End Sub


Private Sub tabGasto_Click(PreviousTab As Integer)

    Select Case tabGasto.Tab
        Case 1, 2
            cmdAccion.Visible = True
            If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vQuery)
            If PreviousTab = 1 Then
                If (cboGasto.ListIndex = 0) Or (lblProveedor.Caption = "") Or (txtDescripGasto.Text = "") Then
                    MsgBox "Faltan Datos", vbExclamation
                    tabGasto.Tab = 1
                End If
                If cboTipoGasto.ListIndex = 0 And CDec(txtMontoGasto.Text) = 0 Then
                    MsgBox "El importe del gasto no puede ser cero", vbExclamation
                    tabGasto.Tab = 1
                End If
            End If
            If strEstado = Reg_Defecto Then tabGasto.Tab = 0
        Case 0
            'If PreviousTab > 0 Then
            cmdAccion.Visible = False
            'End If
    End Select
    
End Sub

Private Sub tdgConsulta_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)

    If ColIndex = 7 Then
        Call DarFormatoValor(Value, Decimales_Monto)
    End If
    
End Sub


Private Sub txtBaseCalculo_Change()

    Call FormatoCajaTexto(txtBaseCalculo, Decimales_Monto)

End Sub

Private Sub txtBaseCalculo_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtBaseCalculo, Decimales_Monto)

End Sub

Private Sub txtMontoGasto_Change()

    Call FormatoCajaTexto(txtMontoGasto, Decimales_Monto)
    
End Sub


Private Sub txtMontoGasto_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtMontoGasto, Decimales_Monto)
    
End Sub

Private Sub GenerarPeriodos()

    Dim adoConsulta             As ADODB.Recordset
    Dim intNumMesesPeriodo      As Integer
    Dim datFechaIniValor        As Date, datFechaIniCupon       As Date
    Dim datFechaFinCupon        As Date, datFechaPago           As Date
    Dim blnCupon                As Boolean, blnIndProceso       As Boolean
    Dim intNumDiasPeriodo       As Integer, intNumDias          As Integer
    Dim intNumPeriodosAnual     As Integer, intUltimoDiaMes     As Integer
        
    datFechaIniValor = dtpFechaInicio.Value  '*** Por defecto a partir de la fecha de inicio ***
   
    datFechaIniCupon = datFechaIniValor '*** Fecha de inicio del cupón ***

    intNumSecuencial = 0: blnCupon = True: blnIndProceso = False
    
    If strEstado = Reg_Adicion Then
        With adoComm
            .CommandText = "SELECT MAX(NumGasto) AS NumGasto  FROM FondoGasto " & _
                "WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora = '" & gstrCodAdministradora & "' "
            Set adoConsulta = .Execute
            intSecuencialGasto = adoConsulta("NumGasto")
        End With
    End If

    With adoComm
        .CommandText = "SELECT NumPeriodo,FechaVencimiento FROM FondoGastoPeriodo " & _
            "WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora = '" & gstrCodAdministradora & "' AND NumGasto = " & intSecuencialGasto & " AND FechaInicio <= '" & Convertyyyymmdd(gdatFechaActual) & "' ORDER BY NumPeriodo DESC"

        Set adoConsulta = .Execute

        If Not adoConsulta.EOF Then
            'intNumSecuencial = (adoConsulta("NumPeriodo"))
            
            If adoConsulta("FechaVencimiento") < gdatFechaActual Then
                datFechaIniCupon = DateAdd("d", 1, adoConsulta("FechaVencimiento"))
                intNumSecuencial = (adoConsulta("NumPeriodo"))
                .CommandText = "DELETE FondoGastoPeriodo WHERE " & _
                " CodFondo='" & strCodFondo & "' AND CodAdministradora = '" & gstrCodAdministradora & "' AND NumGasto = " & intSecuencialGasto & " AND NumPeriodo > " & adoConsulta("NumPeriodo")
            Else
                .CommandText = "DELETE FondoGastoPeriodo WHERE " & _
                " CodFondo='" & strCodFondo & "' AND CodAdministradora = '" & gstrCodAdministradora & "' AND NumGasto = " & intSecuencialGasto & " AND NumPeriodo >= " & adoConsulta("NumPeriodo")
            End If
            adoConn.Execute .CommandText
        Else
            blnCupon = True 'False
        End If
        adoConsulta.Close: Set adoConsulta = Nothing
    End With
    
    If strCodModalidadCalculo = Codigo_Modalidad_Devengo_Inmediata Then Exit Sub
    
    With adoComm
        Set adoConsulta = New ADODB.Recordset

        '*** Obtener el número de días del periodo de pago ***
        
        .CommandText = "SELECT ValorParametro FROM AuxiliarParametro WHERE CodTipoParametro='TIPFRE' AND CodParametro='" & strCodPeriodoPago & "'"
        Set adoConsulta = .Execute

        If Not adoConsulta.EOF Then
            intNumDiasPeriodo = CInt(adoConsulta("ValorParametro")) '*** Días del periodo  ***
            intNumMesesPeriodo = CInt(intNumDiasPeriodo / 30)       '*** Meses del periodo ***
            If intNumMesesPeriodo > 0 Then
                intNumPeriodosAnual = CInt(12 / intNumMesesPeriodo)     '*** Periodos al año   ***
            Else
                intNumPeriodosAnual = 0
            End If
        End If
        adoConsulta.Close: Set adoConsulta = Nothing
    
    End With
    
    Do While datFechaIniCupon <= dtpFechaFin.Value
        intNumSecuencial = intNumSecuencial + 1
       
        '*** Fecha de corte del primer cupón ***
        '*** Mes calendario ***
        If intNumSecuencial = 1 And chkIndicadorPagoFinMes.Value = vbChecked And intNumMesesPeriodo > 0 Then
            datFechaFinCupon = datFechaIniCupon
            intUltimoDiaMes = UltimoDiaMes(Month(datFechaFinCupon), Year(datFechaFinCupon))
            If intUltimoDiaMes <> Day(datFechaFinCupon) Then
                datFechaFinCupon = DateSerial(Year(datFechaFinCupon), Month(datFechaFinCupon), intUltimoDiaMes)
            End If
        ElseIf intNumSecuencial > 1 Or chkIndicadorPagoFinMes.Value = vbChecked Then
            If intNumMesesPeriodo > 0 Then
                datFechaFinCupon = DateAdd("m", intNumMesesPeriodo, datFechaIniCupon) - 1
            Else
                datFechaFinCupon = DateAdd("d", intNumDiasPeriodo, datFechaIniCupon) - 1
            End If

            'datFechaFinCupon = CalculaFechaSiguienteCalendario(datFechaIniCupon, cboBaseCalculo.ListIndex, 7, cboPeriodoPago.ListIndex, CInt(txtValorPeriodoPago.Text))
            'datFechaFinCupon = DateAdd("d", -1, datFechaFinCupon)
            If chkIndicadorPagoFinMes.Value = vbChecked And intNumMesesPeriodo > 0 Then
                intUltimoDiaMes = UltimoDiaMes(Month(datFechaFinCupon), Year(datFechaFinCupon))
                If intUltimoDiaMes <> Day(datFechaFinCupon) Then
                    datFechaFinCupon = DateSerial(Year(datFechaFinCupon), Month(datFechaFinCupon), intUltimoDiaMes)
                End If
            End If
        Else
            datFechaFinCupon = dtpFechaPrimerPago.Value
        End If
        
        '*** SI EL NRO DE DIAS QUE FALTA PARA EL VCTO. YA NO CUBRE PARA GENERAR OTRO ***
        '*** CUPON, CONTROLAR QUE EL ULTIMO TENGA MAS DE UN DIA                      ***
        intNumDias = DateDiff("d", datFechaFinCupon, dtpFechaFin.Value) - 1

        If intNumDias <= 0 Then
            blnIndProceso = True
        Else
            blnIndProceso = False
            If dtpFechaFin.Value = datFechaFinCupon Then
                blnIndProceso = True
            Else
                blnIndProceso = False
            End If
        End If
       
        '*** Inicio para el último cupón
        If blnIndProceso = True Then
            datFechaFinCupon = dtpFechaFin.Value
            strFechaInicio = Convertyyyymmdd(datFechaIniCupon)
            strFechaFin = Convertyyyymmdd(datFechaFinCupon)
            intCantDias = DateDiff("d", datFechaIniCupon, datFechaFinCupon) + 1
            If strCodTipoDesplazamiento <> "" Then
                strFechaPago = Convertyyyymmdd(DesplazamientoDiaUtil(datFechaFinCupon, strCodTipoDesplazamiento))
            Else
                strFechaPago = strFechaFin
            End If

            '*** Grabar en temporal ***
            Call GrabarFechaCorteTmp
            
            blnCupon = False: Exit Do
        End If
        
        'JAFR 09/03/11 Caso de pago unico:
        If strCodTipoPago = "02" Then
            datFechaFinCupon = dtpFechaFin.Value
        End If
        'FIN JAFR
        
        If DateDiff("d", datFechaFinCupon, gdatFechaActual) >= 0 Then
            strIndVigente = ""
        Else
            If intNumPeriodo = 0 Then
                strIndVigente = "X": intNumPeriodo = 1
            Else
                strIndVigente = " "
            End If
        End If
   
        intNumPeriodo = intNumSecuencial
        strFechaInicio = Convertyyyymmdd(datFechaIniCupon)
        strFechaFin = Convertyyyymmdd(datFechaFinCupon)
        
        If strCodModalidadPago = "02" Then
            'caso de pago adelantado:
            strFechaPago = strFechaInicio
        Else
            'caso de pago al vencimiento
            If strCodTipoDesplazamiento <> "" Then
                strFechaPago = Convertyyyymmdd(DesplazamientoDiaUtil(datFechaFinCupon, strCodTipoDesplazamiento))
            Else
                strFechaPago = strFechaFin
            End If
        End If
        
        intCantDias = DateDiff("d", datFechaIniCupon, datFechaFinCupon) + 1
       
        '*** Grabar en temporal ***
        Call GrabarFechaCorteTmp
       
        '*** Para siguiente entrada ***
        datFechaIniCupon = DateAdd("d", 1, datFechaFinCupon)
    Loop
   
    
End Sub

Private Sub GrabarFechaCorteTmp()

    Dim intRegistro As Integer
    
    With adoComm
        
        .CommandText = "{ call up_GNManFondoGastoPeriodo('" & strCodFondo & "','" & _
            gstrCodAdministradora & "'," & intSecuencialGasto & "," & intNumSecuencial & ",'" & _
            strCodFile & "','" & strCodAnalitica & "','" & _
            strFechaInicio & "','" & _
            strFechaFin & "','" & strFechaPago & "'," & _
            intCantDias & ",'" & strIndVigente & "','','" & IIf(strEstado = Reg_Adicion, "I", "U") & "') }"
        
        adoConn.Execute .CommandText
       
    End With
    
End Sub

