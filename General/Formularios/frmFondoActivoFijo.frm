VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{830D5A06-9B70-4F8C-98B6-7A19C4E7760B}#1.0#0"; "TAMControls.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmFondoActivoFijo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Activos Fijos"
   ClientHeight    =   8145
   ClientLeft      =   855
   ClientTop       =   1050
   ClientWidth     =   11565
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
   ScaleHeight     =   8145
   ScaleWidth      =   11565
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   735
      Left            =   8640
      Picture         =   "frmFondoActivoFijo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   7320
      Width           =   1200
   End
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   10080
      TabIndex        =   70
      Top             =   7320
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
      Left            =   5160
      TabIndex        =   69
      Top             =   7320
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
      Left            =   240
      TabIndex        =   68
      Top             =   7320
      Width           =   4200
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
      Caption2        =   "&Anular"
      Tag2            =   "4"
      Visible2        =   0   'False
      ToolTipText2    =   "Anular"
      UserControlWidth=   4200
   End
   Begin TabDlg.SSTab tabGasto 
      Height          =   7035
      Left            =   60
      TabIndex        =   5
      Top             =   90
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   12409
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      ForeColor       =   8388608
      TabCaption(0)   =   "Lista"
      TabPicture(0)   =   "frmFondoActivoFijo.frx":05EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tdgConsulta"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraGastos(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "frmFondoActivoFijo.frx":0608
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblCodProveedor"
      Tab(1).Control(1)=   "fraGastos(1)"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Condiciones"
      TabPicture(2)   =   "frmFondoActivoFijo.frx":0624
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblDescrip(18)"
      Tab(2).Control(1)=   "lblDescrip(20)"
      Tab(2).Control(2)=   "Frame2"
      Tab(2).Control(3)=   "Frame3"
      Tab(2).ControlCount=   4
      Begin VB.Frame Frame3 
         Caption         =   "Contables"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2835
         Left            =   -74760
         TabIndex        =   49
         Top             =   630
         Width           =   10695
         Begin VB.ComboBox cboFrecuenciaDevengo 
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
            Left            =   7200
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   1770
            Width           =   2625
         End
         Begin VB.ComboBox cboAplicacionDevengo 
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
            TabIndex        =   52
            Top             =   1770
            Width           =   2595
         End
         Begin VB.ComboBox cboTipoDevengo 
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
            TabIndex        =   51
            Top             =   1320
            Width           =   2595
         End
         Begin VB.CheckBox chkNoIncluyeEnBalancePrecierre 
            Caption         =   "No incluye en Balance de Precierre"
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
            Left            =   360
            TabIndex        =   50
            Top             =   2250
            Width           =   3315
         End
         Begin MSComCtl2.DTPicker dtpFechaInicio 
            Height          =   315
            Left            =   3180
            TabIndex        =   54
            Top             =   420
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
            Format          =   175898625
            CurrentDate     =   38068
         End
         Begin MSComCtl2.DTPicker dtpFechaFin 
            Height          =   315
            Left            =   3180
            TabIndex        =   55
            Top             =   810
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
            Format          =   175898625
            CurrentDate     =   38068
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Rango Devengo:"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   19
            Left            =   360
            TabIndex        =   61
            Top             =   420
            Width           =   1455
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Del"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   11
            Left            =   2640
            TabIndex        =   60
            Top             =   450
            Width           =   300
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Al"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   9
            Left            =   2640
            TabIndex        =   59
            Top             =   840
            Width           =   180
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Frec. Devengo"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   22
            Left            =   5640
            TabIndex        =   58
            Top             =   1830
            Width           =   1275
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Aplicación Devengo"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   21
            Left            =   390
            TabIndex        =   57
            Top             =   1800
            Width           =   1710
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Devengo"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   5
            Left            =   360
            TabIndex        =   56
            Top             =   1320
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Fechas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1785
         Left            =   -74730
         TabIndex        =   26
         Top             =   3960
         Visible         =   0   'False
         Width           =   10725
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
            Left            =   2550
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   420
            Width           =   2655
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
            Left            =   2550
            Style           =   2  'Dropdown List
            TabIndex        =   35
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
            Left            =   6120
            TabIndex        =   34
            Top             =   840
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
            Left            =   7410
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   840
            Width           =   2505
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
            Left            =   7410
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   1320
            Width           =   2505
         End
         Begin VB.CheckBox chkIndicadorPagoFinMes 
            Caption         =   "Fin de Mes"
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
            Left            =   390
            TabIndex        =   27
            Top             =   1740
            Width           =   1245
         End
         Begin MSComCtl2.DTPicker dtpFechaPrimerPago 
            Height          =   315
            Left            =   2550
            TabIndex        =   28
            Top             =   1320
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
            Format          =   175898625
            CurrentDate     =   38068
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Modalidad Pago"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   31
            Left            =   360
            TabIndex        =   37
            Top             =   450
            Width           =   1380
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Pago"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   29
            Left            =   360
            TabIndex        =   36
            Top             =   900
            Width           =   885
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Cada"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   28
            Left            =   5520
            TabIndex        =   33
            Top             =   900
            Width           =   450
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Desplazamiento"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   24
            Left            =   5520
            TabIndex        =   31
            Top             =   1380
            Width           =   1785
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "1era. Fecha Pago"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   23
            Left            =   360
            TabIndex        =   29
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
         Height          =   6165
         Index           =   1
         Left            =   -74760
         TabIndex        =   8
         Top             =   480
         Width           =   10800
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
            TabIndex        =   62
            Top             =   3720
            Width           =   2655
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
            Left            =   2760
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   5280
            Visible         =   0   'False
            Width           =   4065
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
            Left            =   7680
            Style           =   2  'Dropdown List
            TabIndex        =   45
            Top             =   4680
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.ComboBox cboTipoGasto 
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
            Left            =   2760
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   4680
            Visible         =   0   'False
            Width           =   2745
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
            Left            =   9510
            TabIndex        =   14
            ToolTipText     =   "Buscar Proveedor"
            Top             =   2400
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
            Top             =   1080
            Width           =   7245
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
            Top             =   1755
            Width           =   7245
         End
         Begin MSComCtl2.DTPicker dtpFechaGasto 
            Height          =   315
            Left            =   2640
            TabIndex        =   1
            Top             =   360
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
            Format          =   175898625
            CurrentDate     =   38068
         End
         Begin TAMControls.TAMTextBox txtMontoGasto 
            Height          =   315
            Left            =   2610
            TabIndex        =   67
            Top             =   4320
            Width           =   1785
            _ExtentX        =   3149
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
            Container       =   "frmFondoActivoFijo.frx":0640
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   12648447
            EnterTab        =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   2
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   10
            Left            =   360
            TabIndex        =   65
            Top             =   3840
            Width           =   690
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Importe del Activo Fijo"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   360
            TabIndex        =   64
            Top             =   4350
            Width           =   1920
         End
         Begin VB.Label lblMagnitud 
            AutoSize        =   -1  'True
            Caption         =   "Magnitud"
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
            Left            =   4590
            TabIndex        =   63
            Top             =   4410
            Width           =   855
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Crédito Fiscal"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   12
            Left            =   720
            TabIndex        =   48
            Top             =   5280
            Visible         =   0   'False
            Width           =   1605
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Estado"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   17
            Left            =   6480
            TabIndex        =   46
            Top             =   4680
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Gasto"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   27
            Left            =   720
            TabIndex        =   44
            Top             =   4800
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "xxxx"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   26
            Left            =   1140
            TabIndex        =   25
            Top             =   7230
            Width           =   375
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "xxxx"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   8
            Left            =   1830
            TabIndex        =   24
            Top             =   7200
            Width           =   375
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "xxx"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   4
            Left            =   2490
            TabIndex        =   21
            Top             =   7170
            Width           =   285
         End
         Begin VB.Line Line3 
            X1              =   300
            X2              =   9870
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
            TabIndex        =   20
            Top             =   3120
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
            Left            =   7560
            TabIndex        =   19
            Top             =   3090
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
            TabIndex        =   18
            Top             =   3090
            Width           =   4815
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   13
            Left            =   360
            TabIndex        =   16
            Top             =   2445
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
            TabIndex        =   15
            Top             =   2400
            Width           =   6780
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
            TabIndex        =   4
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
            TabIndex        =   13
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
            TabIndex        =   12
            Top             =   1140
            Width           =   825
         End
         Begin VB.Label lblAnalitica 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "030-00000000"
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
            TabIndex        =   11
            Top             =   390
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Detalle"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   10
            Top             =   1800
            Width           =   615
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   9
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
         Height          =   1545
         Index           =   0
         Left            =   330
         TabIndex        =   6
         Top             =   480
         Width           =   10125
         Begin VB.CheckBox chkVerSoloGastosContabilizados 
            Caption         =   "Ver sólo act. fijo contabilizados"
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   6480
            TabIndex        =   71
            Top             =   1020
            Width           =   3435
         End
         Begin VB.CheckBox chkVerSoloGastosVigentes 
            Caption         =   "Ver sólo activos fijos vigentes"
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   360
            TabIndex        =   39
            Top             =   1020
            Width           =   3945
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
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   450
            Width           =   6315
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fondo"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   7
            Top             =   480
            Width           =   540
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmFondoActivoFijo.frx":065C
         Height          =   4185
         Left            =   360
         OleObjectBlob   =   "frmFondoActivoFijo.frx":0676
         TabIndex        =   66
         Top             =   2160
         Width           =   10125
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   20
         Left            =   -74490
         TabIndex        =   23
         Top             =   5160
         Width           =   75
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   18
         Left            =   -74550
         TabIndex        =   22
         Top             =   6030
         Width           =   405
      End
      Begin VB.Label lblCodProveedor 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   -74700
         TabIndex        =   17
         Top             =   8820
         Width           =   645
      End
   End
   Begin TAMControls.ucBotonEdicion cmdOpcion2 
      Height          =   390
      Left            =   720
      TabIndex        =   40
      Top             =   8520
      Visible         =   0   'False
      Width           =   4200
      _ExtentX        =   7408
      _ExtentY        =   688
      Buttons         =   3
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
      UserControlHeight=   390
      UserControlWidth=   4200
   End
   Begin TAMControls.ucBotonEdicion cmdAccion2 
      Height          =   390
      Left            =   5400
      TabIndex        =   41
      Top             =   8520
      Visible         =   0   'False
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
   Begin TAMControls.ucBotonEdicion cmdSalir2 
      Height          =   390
      Left            =   9600
      TabIndex        =   42
      Top             =   8520
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   688
      Caption0        =   "&Salir"
      Tag0            =   "9"
      Visible0        =   0   'False
      ToolTipText0    =   "Salir"
      UserControlHeight=   390
      UserControlWidth=   1200
   End
End
Attribute VB_Name = "frmFondoActivoFijo"
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
Dim arrModalidadPago()          As String, arrTipoDevengo()             As String
Dim arrAplicacionDevengo()      As String, arrPeriodoTasa()             As String
Dim arrTipoTasa()               As String, arrTipoGasto()               As String
Dim arrBaseCalculo()            As String


Dim strCodFondo                 As String, strCodTipoProvision          As String
Dim strCodMoneda                As String, strCodFrecuenciaDevengo      As String
Dim strCodCuenta                As String, strCodGasto                  As String
Dim strCodFile                  As String, strCodAnalitica              As String
Dim strCodTipoPago              As String, strCodDetalleGasto           As String
Dim strCodCreditoFiscal         As String, strCodTipoDesplazamiento     As String
Dim strEstado                   As String, strSQL                       As String
Dim strCodTipoValor             As String, strCodPeriodoPago            As String
Dim strCodModalidadPago         As String, strCodTipoDevengo            As String
Dim strCodAplicacionDevengo     As String, strCodPeriodoTasa            As String
Dim strCodTipoTasa              As String, strCodTipoGasto              As String
Dim strCodBaseCalculo           As String, strEstadoGasto                  As String

Dim intNumPeriodo               As Integer, strFechaInicio              As String
Dim strFechaFin                 As String, strFechaPago                 As String
Dim intCantDias                 As Integer, strIndVigente               As String
Dim intSecuencialGasto          As Integer, intNumSecuencial            As Integer
Dim adoConsulta                 As ADODB.Recordset


Public Sub Buscar()
    
    Set adoConsulta = New ADODB.Recordset
    
'    strSQL = "SELECT FG.CodGasto,FG.CodDetalleGasto,NumGasto,DCG.CodAnalitica,CG.DescripConcepto,DCG.DescripGasto,MontoGasto " & _
'        "FROM FondoGasto FG JOIN DetalleConceptoGasto DCG ON(DCG.CodDetalleGasto=FG.CodDetalleGasto AND DCG.CodGasto=FG.CodGasto) " & _
'        "JOIN ConceptoGasto CG ON(CG.CodGasto=DCG.CodGasto) " & _
'        "WHERE CodTipoGasto='" & strCodTipoProvision & "' AND CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND FG.IndVigente='X'" & _
'        "ORDER BY NumGasto"
'    strSQL = "SELECT FG.CodCuenta,FG.NumGasto,FG.CodAnalitica,PCG.DescripCuenta,FG.DescripGasto,FG.MontoGasto,FG.CodFile,INP.DescripPersona as DescripProveedor " & _
'        "FROM FondoGasto FG JOIN FondoConceptoGasto FCG ON(FCG.CodCuenta=FG.CodCuenta AND FCG.CodAdministradora=FG.CodAdministradora AND FCG.CodFondo=FG.CodFondo) " & _
'        "JOIN PlanContable PCG ON(PCG.CodCuenta=FG.CodCuenta) " & _
'        "JOIN InstitucionPersona INP ON(INP.CodPersona=FG.CodProveedor AND INP.TipoPersona = '" & Codigo_Tipo_Persona_Proveedor & "') " & _
'        "WHERE CodFile='099' AND CodTipoGasto='" & strCodTipoProvision & "' AND FG.CodFondo='" & strCodFondo & "' AND FG.CodAdministradora='" & gstrCodAdministradora & "' AND FG.IndVigente='X' " & _
'        " UNION " & _
'        "SELECT FG.CodCuenta,FG.NumGasto,FG.CodAnalitica,FG.DescripGasto DescripCuenta,DescripGasto,MontoGasto,CodFile,INP.DescripPersona as DescripProveedor " & _
'        "FROM FondoGasto FG " & _
'        "JOIN InstitucionPersona INP ON(INP.CodPersona=FG.CodProveedor AND INP.TipoPersona = '" & Codigo_Tipo_Persona_Proveedor & "') " & _
'        "WHERE CodFile='098' AND CodTipoGasto='" & strCodTipoProvision & "' AND CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND FG.IndVigente='X'"
                        
    strSQL = "SELECT FG.CodCuenta,FG.NumGasto,FG.CodAnalitica,FG.FechaDefinicion,PCG.DescripCuenta,PRV.DescripParametro as TipoGasto, FG.DescripGasto,FG.MontoGasto,FG.CodFile,INP.DescripPersona as DescripProveedor, FG.IndConfirma " & _
        "FROM FondoGasto FG JOIN FondoConceptoActivoFijo FCG ON(FCG.CodCuenta=FG.CodCuenta AND FCG.CodAdministradora=FG.CodAdministradora AND FCG.CodFondo=FG.CodFondo) " & _
        "JOIN PlanContable PCG ON(PCG.CodCuenta=FG.CodCuenta) " & _
        "JOIN InstitucionPersona INP ON(INP.CodPersona=FG.CodProveedor AND INP.TipoPersona = '" & Codigo_Tipo_Persona_Proveedor & "') " & _
        "JOIN AuxiliarParametro PRV ON(PRV.CodParametro=FG.CodTipoGasto AND PRV.CodTipoParametro='TIPPAG') " & _
        "WHERE FG.CodFondo='" & strCodFondo & "' AND FG.CodAdministradora='" & gstrCodAdministradora & "'"
                        
    If chkVerSoloGastosVigentes.Value = vbChecked Then
        strSQL = strSQL & " AND (FG.IndVigente = 'X') "
        If chkVerSoloGastosContabilizados.Value = vbChecked Then
            strSQL = strSQL & " AND (IndConfirma = 'X') "
        Else
            strSQL = strSQL & " AND (IndConfirma = '') "
        End If
    Else
        strSQL = strSQL & " AND (FG.IndVigente='') "
    End If
    
    strSQL = strSQL & " ORDER BY FG.NumGasto "
                        
    strEstado = Reg_Defecto
    
    With adoConsulta
        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open strSQL
    End With
        
    tdgConsulta.DataSource = adoConsulta
    
    Call AutoAjustarGrillas
        
    tdgConsulta.Refresh
    
    If adoConsulta.RecordCount > 0 Then strEstado = Reg_Consulta
            
End Sub


Private Sub CargarGastos()

    '*** Gastos del Fondo ***
'    strSQL = "SELECT (FCG.CodDetalleGasto + FCG.CodGasto + DCG.CodAnalitica) CODIGO,(RTRIM(CG.DescripConcepto) + '-' + RTRIM(DCG.DescripGasto)) DESCRIP " & _
'        "FROM FondoConceptoGasto FCG JOIN DetalleConceptoGasto DCG ON(DCG.CodDetalleGasto=FCG.CodDetalleGasto AND DCG.CodGasto=FCG.CodGasto) " & _
'        "JOIN ConceptoGasto CG ON (CG.CodGasto=DCG.CodGasto) " & _
'        "WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' " & _
'        "ORDER BY DCG.DescripGasto"
    strSQL = "SELECT FCG.CodCuenta CODIGO,(RTRIM(DescripCuenta)) DESCRIP " & _
        "FROM FondoConceptoActivoFijo FCG JOIN PlanContable PCG ON(PCG.CodCuenta=FCG.CodCuenta AND PCG.CodAdministradora=FCG.CodAdministradora) " & _
        "WHERE CodFondo='" & strCodFondo & "' AND FCG.CodAdministradora='" & gstrCodAdministradora & "' " & _
        "ORDER BY DescripCuenta"
    CargarControlLista strSQL, cboGasto, arrGasto(), Sel_Defecto
    
End Sub

Public Sub Salir()

    Unload Me
    
End Sub



Private Sub cboAplicacionDevengo_Click()

    strCodAplicacionDevengo = Valor_Caracter
    If cboAplicacionDevengo.ListIndex < 0 Then Exit Sub
    
    strCodAplicacionDevengo = Trim(arrAplicacionDevengo(cboAplicacionDevengo.ListIndex))

    If strCodAplicacionDevengo = Codigo_Aplica_Devengo_Inmediata Then
        lblDescrip(22).Visible = False
        cboFrecuenciaDevengo.Visible = False
        cboFrecuenciaDevengo.ListIndex = -1
    Else
        lblDescrip(22).Visible = True
        cboFrecuenciaDevengo.Visible = True
        cboFrecuenciaDevengo.ListIndex = -1
    End If

End Sub

'Private Sub cboBaseCalculo_Click()
'
'    strCodBaseCalculo = Valor_Caracter
'    If cboBaseCalculo.ListIndex < 0 Then Exit Sub
'
'    strCodBaseCalculo = Trim(arrBaseCalculo(cboBaseCalculo.ListIndex))
'
'End Sub

Private Sub cboEstado_Click()

    strEstadoGasto = Valor_Caracter
    If cboEstado.ListIndex < 0 Then Exit Sub
    
    strEstadoGasto = Trim(arrEstado(cboEstado.ListIndex))

End Sub

Private Sub cboModalidadPago_Click()

    strCodModalidadPago = Valor_Caracter
    If cboModalidadPago.ListIndex < 0 Then Exit Sub
    
    strCodModalidadPago = Trim(arrModalidadPago(cboModalidadPago.ListIndex))

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
    
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Fecha Vigente, Valor Cuota, Valor Cuota T-1, Moneda y Cantidad Inicial de Cuotas del Fondo ***
        .CommandText = "{ call up_ACSelDatosParametro(23,'" & strCodFondo & "','" & gstrCodAdministradora & "') }"
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


Private Sub cboFrecuenciaDevengo_Click()

    strCodFrecuenciaDevengo = Valor_Caracter
    If cboFrecuenciaDevengo.ListIndex < 0 Then Exit Sub
    
    strCodFrecuenciaDevengo = Trim(arrFrecuenciaDevengo(cboFrecuenciaDevengo.ListIndex))
    
End Sub


Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Listado de Activos Fijos del Fondo Vigentes"
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Text = "Listado de Activos Fijos del Fondo No Vigentes"

End Sub





Private Sub cboPeriodoPago_Click()

    strCodPeriodoPago = Valor_Caracter
    If cboPeriodoPago.ListIndex < 0 Then Exit Sub
    
    strCodPeriodoPago = Trim(arrPeriodoPago(cboPeriodoPago.ListIndex))
    
End Sub



Private Sub cboGasto_Click()

    strCodGasto = Valor_Caracter: strCodAnalitica = Valor_Caracter
    If cboGasto.ListIndex <= 0 Then Exit Sub
    
    strCodGasto = Trim(arrGasto(cboGasto.ListIndex))
    
    Dim adoRegistro     As ADODB.Recordset
    
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        .CommandText = "SELECT CodDetalleFile FROM InversionDetalleFile " & _
            "WHERE CodFile='" & strCodFile & "' AND DescripDetalleFile='" & strCodGasto & "'"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            strCodAnalitica = Format(adoRegistro("CodDetalleFile"), "00000000")
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    If strEstado = Reg_Adicion Then
        lblAnalitica.Caption = "030-" & strCodAnalitica
    Else
        lblAnalitica.Caption = Trim(tdgConsulta.Columns(9).Value) & "-" & strCodAnalitica
    End If
    
End Sub


Private Sub cboMoneda_Click()

    strCodMoneda = Valor_Caracter
    If cboMoneda.ListIndex < 0 Then Exit Sub
    
    strCodMoneda = Trim(arrMoneda(cboMoneda.ListIndex))
    
    'lblMoneda.Caption = ObtenerCodSignoMoneda(strCodMoneda)
    
    'If strCodTipoValor = Codigo_Tipo_Costo_Monto Then
    lblMagnitud.Caption = ObtenerCodSignoMoneda(strCodMoneda)
    'End If
    
End Sub







'Private Sub cboPeriodoTasa_Click()
'
'    strCodPeriodoTasa = Valor_Caracter
'    If cboPeriodoTasa.ListIndex < 0 Then Exit Sub
'
'    strCodPeriodoTasa = Trim(arrPeriodoTasa(cboPeriodoTasa.ListIndex))
'
'    lblMagnitud.Caption = "% " & cboPeriodoTasa.Text
'
'
'End Sub

Private Sub cboTipoDesplazamiento_Click()

    strCodTipoDesplazamiento = Valor_Caracter
    If cboTipoDesplazamiento.ListIndex < 0 Then Exit Sub
    
    strCodTipoDesplazamiento = Trim(arrTipoDesplazamiento(cboTipoDesplazamiento.ListIndex))

End Sub



Private Sub cboTipoDevengo_Click()

    strCodTipoDevengo = Valor_Caracter
    If cboTipoDevengo.ListIndex < 0 Then Exit Sub
    
    strCodTipoDevengo = Trim(arrTipoDevengo(cboTipoDevengo.ListIndex))
    
End Sub


Private Sub cboTipoGasto_Click()

    strCodTipoGasto = Valor_Caracter
    If cboTipoGasto.ListIndex < 0 Then Exit Sub
    
    strCodTipoGasto = Trim(arrTipoGasto(cboTipoGasto.ListIndex))

End Sub

Private Sub cboTipoPago_Click()

    strCodTipoPago = Valor_Caracter
    If cboTipoPago.ListIndex < 0 Then Exit Sub
    
    strCodTipoPago = Trim(arrTipoPago(cboTipoPago.ListIndex))
    
    If strCodTipoPago = "01" Then
        lblDescrip(28).Visible = True
        txtValorPeriodoPago.Visible = True
        txtValorPeriodoPago.Text = "1"
        cboPeriodoPago.ListIndex = -1
        cboPeriodoPago.Visible = True
    Else
        lblDescrip(28).Visible = False
        txtValorPeriodoPago.Visible = False
        txtValorPeriodoPago.Text = "0"
        cboPeriodoPago.ListIndex = -1
        cboPeriodoPago.Visible = False
    End If
    
End Sub


Private Sub chkVerSoloGastosContabilizados_Click()
    If chkVerSoloGastosContabilizados.Value = vbChecked Then
        chkVerSoloGastosVigentes.Value = vbChecked
    End If
    
    Call Buscar
End Sub

'Private Sub cboTipoTasa_Click()
'
'    strCodTipoTasa = Valor_Caracter
'    If cboTipoTasa.ListIndex < 0 Then Exit Sub
'
'    strCodTipoTasa = Trim(arrTipoTasa(cboTipoTasa.ListIndex))
'
'    If strCodTipoTasa <> "03" Then
'        cboPeriodoTasa.Enabled = True
'    Else
'        cboPeriodoTasa.ListIndex = -1
'        cboPeriodoTasa.Enabled = False
'        lblMagnitud.Caption = "% "
'    End If
'
'
'End Sub

'Private Sub cboTipoValor_Click()
'
'    strCodTipoValor = Valor_Caracter
'    If cboTipoValor.ListIndex < 0 Then Exit Sub
'
'    strCodTipoValor = Trim(arrTipoValor(cboTipoValor.ListIndex))
'
'    If strCodTipoValor = Codigo_Tipo_Costo_Monto Then
'        lblMagnitud.Caption = lblMoneda.Caption
'        txtBaseCalculo.Visible = False
'        lblMoneda.Visible = False
'        lblDescrip(16).Visible = False
'        lblDescrip(25).Visible = False
'        lblDescrip(30).Visible = False
'        lblDescrip(32).Visible = False
'        txtBaseCalculo.Text = ""
'        cboTipoTasa.ListIndex = -1
'        cboTipoTasa.Visible = False
'        cboPeriodoTasa.ListIndex = -1
'        cboPeriodoTasa.Visible = False
'        cboBaseCalculo.ListIndex = -1
'        cboBaseCalculo.Visible = False
'    End If
'
'    If strCodTipoValor = Codigo_Tipo_Costo_Porcentaje Then
'        lblMagnitud.Caption = "%"
'        txtBaseCalculo.Visible = True
'        lblMoneda.Visible = True
'        lblDescrip(16).Visible = True
'        lblDescrip(25).Visible = True
'        lblDescrip(30).Visible = True
'        lblDescrip(32).Visible = True
'        txtBaseCalculo.Text = "0.00"
'        cboTipoTasa.ListIndex = -1
'        cboTipoTasa.Visible = True
'        cboPeriodoTasa.ListIndex = -1
'        cboPeriodoTasa.Visible = True
'        cboBaseCalculo.ListIndex = -1
'        cboBaseCalculo.Visible = True
'    End If
'
'    Call txtMontoGasto_Change
'
'
'End Sub


Private Sub chkVerSoloGastosVigentes_Click()
    If chkVerSoloGastosVigentes.Value = vbUnchecked Then
        chkVerSoloGastosContabilizados.Value = Unchecked
        chkVerSoloGastosContabilizados.Enabled = False
    Else
        chkVerSoloGastosContabilizados.Enabled = True
    End If
    
    Call Buscar
End Sub

Private Sub cmdImprimir_Click()
    
    Call SubImprimir2(1)
    
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
        '.sSql = "{ call up_ACSelDatos(26) }"
        
        .sSql = "SELECT IP.CodPersona CODIGO, IP.TipoPersona, AP.DescripParametro TipoIdentidad, IP.NumIdentidad, " & _
                "IP.DescripPersona DESCRIP,IP.Direccion1 + IP.Direccion2 Direccion " & _
                "FROM InstitucionPersona IP " & _
                "JOIN AuxiliarParametro AP ON(AP.CodParametro=IP.TipoIdentidad AND AP.CodTipoParametro='TIPIDE') " & _
                "WHERE IP.TipoPersona='" & Codigo_Tipo_Persona_Proveedor & "' AND IP.IndVigente='" & Valor_Indicador & "' AND IP.IndBanco<>'" & Valor_Indicador & "' " & _
                "ORDER BY IP.DescripPersona "
        
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

    If dtpFechaFin.Value < gdatFechaActual Then
        dtpFechaFin.Value = gdatFechaActual
    End If
    
    If dtpFechaInicio.Value > dtpFechaFin.Value Then
        dtpFechaInicio.Value = dtpFechaFin.Value
    End If
    
End Sub

Private Sub Form_Load()

    Call InicializarValores
    Call CargarListas
    Call DarFormato
    Call Buscar
    Call CargarReportes
    
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
    
    For intCont = 0 To (fraGastos.Count - 1)
        Call FormatoMarco(fraGastos(intCont))
    Next
    
    For Each elemento In Me.Controls
    
        If TypeOf elemento Is TDBGrid Then
            Call FormatoGrilla(elemento)
        End If
    
    Next
            
End Sub
Private Sub CargarListas()
        
    Dim intRegistro         As Integer
    
    '*** Fondos ***
    strSQL = "{ call up_ACSelDatosParametro(74,'" & gstrCodAdministradora & "','" & gstrCodFondoContable & "') }"
    CargarControlLista strSQL, cboFondo, arrFondo(), Valor_Caracter
    
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
   
    '*** Tipo de Valor ***
'    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='VALCOM' ORDER BY DescripParametro"
'    CargarControlLista strSQL, cboTipoValor, arrTipoValor(), Valor_Caracter
    
    '*** Tipo de Gasto ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPPAG' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboTipoGasto, arrTipoGasto(), Valor_Caracter
    
    '*** Tipo de Desplazamiento ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPDES' ORDER BY CodParametro"
    CargarControlLista strSQL, cboTipoDesplazamiento, arrTipoDesplazamiento(), Valor_Caracter
    
    '*** Moneda ***
    strSQL = "{ call up_ACSelDatos(2) }"
    CargarControlLista strSQL, cboMoneda, arrMoneda(), Valor_Caracter
            
    '*** Tipos de Frecuencias ***
    strSQL = "{ call up_ACSelDatos(17) }"
    CargarControlLista strSQL, cboFrecuenciaDevengo, arrFrecuenciaDevengo(), Valor_Caracter
    
    '*** Tipos de Frecuencias ***
'    strSQL = "{ call up_ACSelDatos(17) }"
'    CargarControlLista strSQL, cboPeriodoTasa, arrPeriodoTasa(), Valor_Caracter
   
    '*** Base de Cálculo ***
'    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='BASANU' ORDER BY DescripParametro"
'    CargarControlLista strSQL, cboBaseCalculo, arrBaseCalculo(), Valor_Caracter
    
    '*** Tipo de Periodos ***
'    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='NATTAS' ORDER BY CodParametro"
'    CargarControlLista strSQL, cboTipoTasa, arrTipoTasa(), Valor_Caracter
      
    '*** Tipo de Periodos ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPPRD' ORDER BY CodParametro"
    CargarControlLista strSQL, cboPeriodoPago, arrPeriodoPago(), Valor_Caracter
    
    '*** Tipo de Pago ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPPAG' ORDER BY CodParametro"
    CargarControlLista strSQL, cboTipoPago, arrTipoPago(), Valor_Caracter
    
    '*** Modalidad de Pago ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='MODPAG' ORDER BY CodParametro"
    CargarControlLista strSQL, cboModalidadPago, arrModalidadPago(), Valor_Caracter
    
    '*** Tipo Crédito Fiscal ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP From AuxiliarParametro WHERE CodTipoParametro='CREFIS' ORDER BY CodParametro"
    CargarControlLista strSQL, cboCreditoFiscal, arrCreditoFiscal(), Valor_Caracter
        
    '*** Estados ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP From AuxiliarParametro WHERE CodTipoParametro='INDREG' AND CodParametro<>'03' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboEstado, arrEstado(), Valor_Caracter
        
    '*** Tipo de Devengo ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP From AuxiliarParametro WHERE CodTipoParametro='TIPDEV' ORDER BY CodParametro"
    CargarControlLista strSQL, cboTipoDevengo, arrTipoDevengo(), Valor_Caracter
        
    '*** Aplicacion de Devengo ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP From AuxiliarParametro WHERE CodTipoParametro='APLDEV' ORDER BY CodParametro"
    CargarControlLista strSQL, cboAplicacionDevengo, arrAplicacionDevengo(), Valor_Caracter
        
    '*** Tipo de desplazamiento
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP From AuxiliarParametro WHERE CodTipoParametro='TIPDES' ORDER BY CodParametro"
    CargarControlLista strSQL, cboTipoDesplazamiento, arrTipoDesplazamiento(), Valor_Caracter
    
        
        
End Sub
Private Sub InicializarValores()
                        
    '*** Valores Iniciales ***
    tabGasto.Tab = 0
    
    strCodFile = "030"
    strCodTipoValor = Codigo_Tipo_Costo_Monto
    
    tabGasto.TabEnabled(1) = False
    
    '** OCULTAR 3ER TAB **
    tabGasto.TabVisible(2) = False
    
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgConsulta.Columns(2).Width = tdgConsulta.Width * 0.01 * 6
    tdgConsulta.Columns(3).Width = tdgConsulta.Width * 0.01 * 10
    tdgConsulta.Columns(5).Width = tdgConsulta.Width * 0.01 * 15
    tdgConsulta.Columns(6).Width = tdgConsulta.Width * 0.01 * 30
    
    chkVerSoloGastosVigentes.Value = 1
    
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    Set cmdOpcion.FormularioActivo = Me
    
End Sub

Private Sub Form_Resize()
    Call AutoAjustarGrillas
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call OcultarReportes
    Set frmFondoActivoFijo = Nothing
    
End Sub

Public Sub Accion(nAccion As ButtonAction)
    
    Select Case nAccion
        
        Case vNew
            Call Adicionar
        Case vModify
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
        .TabEnabled(1) = False
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
    Dim strIndFinMes                    As String
    Dim intDiasProvision                As String
    Dim intDiasBaseAnual                As Integer
    Dim intNumPeriodoAnualTasa          As Integer
    Dim mensaje As String
    
    If strEstado = Reg_Consulta Then Exit Sub
    If Not TodoOK() Then Exit Sub
    
    On Error GoTo CtrlError
 
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
        strIndFinMes = Valor_Indicador
    Else
        strIndFinMes = Valor_Caracter
    End If
    
    dtpFechaPrimerPago.Value = dtpFechaFin.Value
    
    If strEstado = Reg_Adicion Then
    
        mensaje = Mensaje_Adicion
    
        Set adoRegistro = New ADODB.Recordset
        
        If MsgBox(mensaje, vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption) <> vbYes Then Exit Sub
                                        
        Me.MousePointer = vbHourglass
        
        intSecuencialGasto = 0
        
        '*** Guardar ***
        With adoComm
            '*** Obtener el número secuencial ***
            .CommandText = "SELECT MAX(NumGasto) NumSecuencial FROM FondoGasto " & _
                "WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
            Set adoRegistro = .Execute
            
            If Not adoRegistro.EOF Then
                If IsNull(adoRegistro("NumSecuencial")) Then
                    intSecuencialGasto = 1
                Else
                    intSecuencialGasto = CInt(adoRegistro("NumSecuencial")) + 1
                End If
            Else
                intSecuencialGasto = 1
            End If
            
            adoRegistro.Close: Set adoRegistro = Nothing
        
            .CommandText = "{ call up_GNManFondoGasto('" & strCodFondo & "','" & _
                gstrCodAdministradora & "'," & intSecuencialGasto & ",'" & Convertyyyymmdd(dtpFechaGasto.Value) & "','" & strCodGasto & "','" & _
                strCodFile & "','" & strCodAnalitica & "','" & Codigo_Tipo_Persona_Proveedor & "','" & Trim(lblCodProveedor.Caption) & "','" & Trim(txtDescripGasto.Text) & "','" & _
                Convertyyyymmdd(CVDate(Valor_Fecha)) & "','" & Convertyyyymmdd(dtpFechaInicio.Value) & "','" & Convertyyyymmdd(dtpFechaFin.Value) & "','" & _
                strCodTipoGasto & "','','" & strEstadoGasto & "'," & CDec(gdblTipoCambio) & ",'" & strCodMoneda & "','" & strCodTipoValor & "'," & _
                numMontoGasto & "," & numPorcenGasto & ",'" & strCodTipoTasa & "','" & strCodPeriodoTasa & "',0,'" & strCodBaseCalculo & "',0,'" & _
                strCodModalidadPago & "','" & strCodTipoPago & "'," & CLng(txtValorPeriodoPago.Text) & ",'" & strCodPeriodoPago & "','" & _
                Convertyyyymmdd(dtpFechaPrimerPago.Value) & "','" & strCodTipoDesplazamiento & "', 'X','" & strCodTipoDevengo & "','" & _
                strCodAplicacionDevengo & "','" & strCodFrecuenciaDevengo & "','" & strCodCreditoFiscal & "','" & strIndNoIncluyeBalancePreCierre & "','I','','') }"
            
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
            .TabEnabled(1) = False
            .Tab = 0
        End With
        'Call Buscar
    End If
    
    If strEstado = Reg_Edicion Then
        
        mensaje = Mensaje_Edicion
        
        If MsgBox(mensaje, vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption) <> vbYes Then Exit Sub
        
        Me.MousePointer = vbHourglass
        
        'intSecuencialGasto = CInt(tdgConsulta.Columns(2).Value)
        
        '*** Actualizar ***
        With adoComm
            
            .CommandText = "{ call up_GNManFondoGasto('" & strCodFondo & "','" & _
                gstrCodAdministradora & "'," & intSecuencialGasto & ",'" & Convertyyyymmdd(dtpFechaGasto.Value) & "','" & strCodGasto & "','" & _
                strCodFile & "','" & strCodAnalitica & "','" & Codigo_Tipo_Persona_Proveedor & "','" & Trim(lblCodProveedor.Caption) & "','" & Trim(txtDescripGasto.Text) & "','" & _
                Convertyyyymmdd(CVDate(Valor_Fecha)) & "','" & Convertyyyymmdd(dtpFechaInicio.Value) & "','" & Convertyyyymmdd(dtpFechaFin.Value) & "','" & _
                strCodTipoGasto & "','','" & strEstadoGasto & "'," & CDec(gdblTipoCambio) & ",'" & strCodMoneda & "','" & strCodTipoValor & "'," & _
                numMontoGasto & "," & numPorcenGasto & ",'" & strCodTipoTasa & "','" & strCodPeriodoTasa & "',0,'" & strCodBaseCalculo & "',0,'" & _
                strCodModalidadPago & "','" & strCodTipoPago & "'," & CLng(txtValorPeriodoPago.Text) & ",'" & strCodPeriodoPago & "','" & _
                Convertyyyymmdd(dtpFechaPrimerPago.Value) & "','" & strCodTipoDesplazamiento & "','" & strIndFinMes & "','" & strCodTipoDevengo & "','" & _
                strCodAplicacionDevengo & "','" & strCodFrecuenciaDevengo & "','" & strCodCreditoFiscal & "','" & strIndNoIncluyeBalancePreCierre & "','U','','') }"
            
            
'            .CommandText = "{ call up_GNManFondoGasto('" & strCodFondo & "','" & _
'                gstrCodAdministradora & "'," & intSecuencialGasto & ",'" & Convertyyyymmdd(dtpFechaGasto.Value) & "','" & strCodGasto & "','" & _
'                strCodFile & "','" & strCodAnalitica & "','" & Codigo_Tipo_Persona_Proveedor & "','" & Trim(lblCodProveedor.Caption) & "','" & Trim(txtDescripGasto.Text) & "','" & _
'                Convertyyyymmdd(CVDate(Valor_Fecha)) & "','" & Convertyyyymmdd(dtpFechaInicio.Value) & "','" & Convertyyyymmdd(dtpFechaFin.Value) & "','" & _
'                strCodTipoGasto & "','','" & strEstadoGasto & "'," & gdblTipoCambio & ",'" & strCodMoneda & "','" & strCodTipoValor & "'," & _
'                numMontoGasto & "," & numPorcenGasto & ",'" & strCodTipoTasa & "','" & strCodPeriodoTasa & "'," & strCodBaseCalculo & ",'" & strCodBaseCalculo & "',0,'" & _
'                strCodModalidadPago & "','" & strCodTipoPago & "'," & CLng(txtValorPeriodoPago.Text) & ",'" & strCodPeriodoPago & "','" & _
'                Convertyyyymmdd(dtpFechaPrimerPago.Value) & "','" & strCodTipoDesplazamiento & "','" & strIndFinMes & "','" & strCodTipoDevengo & "','" & _
'                strCodAplicacionDevengo & "','" & strCodFrecuenciaDevengo & "','" & strCodCreditoFiscal & "','" & strIndNoIncluyeBalancePreCierre & "','U') }"
            
            adoConn.Execute .CommandText
            
            Call GenerarPeriodos
            
        End With

        Me.MousePointer = vbDefault
                    
        MsgBox Mensaje_Edicion_Exitosa, vbExclamation
        
        frmMainMdi.stbMdi.Panels(3).Text = "Acción"
        
        cmdOpcion.Visible = True
        With tabGasto
            .TabEnabled(0) = True
            .TabEnabled(1) = False
            .Tab = 0
        End With
        'Call Buscar
    End If
    
    Call cboFondo_Click
    
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
    Dim strFechaDesde           As String, strFechaHasta        As String
    Dim strSeleccionRegistro    As String
    
    gstrNameRepo = "FondoActivoFijo"
    
    strSeleccionRegistro = "{Participe.FechaIngreso} IN 'Fch1' TO 'Fch2'"
    gstrSelFrml = strSeleccionRegistro
    frmRangoFecha.Show vbModal
    
    
    If gstrSelFrml <> "0" Then
                        
        Set frmReporte = New frmVisorReporte
    
        ReDim aReportParamS(4)
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
                aReportParamF(4) = "LISTADO DE ACTIVOS FIJOS VIGENTES"
                aReportParamS(2) = Valor_Indicador
            Case 2
                aReportParamF(4) = "LISTADO DE ACTIVOS FIJOS NO VIGENTES"
                aReportParamS(2) = Valor_Caracter
        End Select
        
         'aReportParamS(3) = Convertyyyymmdd(gstrFchDel)
         'aReportParamS(4) = Convertyyyymmdd(gstrFchAl)
        
        
        aReportParamS(3) = Convertyyyymmdd(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10))
        aReportParamS(4) = Convertyyyymmdd(DateAdd("d", 1, CVDate(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10))))

        
        
     End If
        
     
    If gstrSelFrml <> "0" Then
     
            gstrSelFrml = Valor_Caracter
            frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"

            Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())

            frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
            frmReporte.Show vbModal

            Set frmReporte = Nothing

            Screen.MousePointer = vbNormal
    
    End If
    
End Sub
Public Sub SubImprimir2(Index As Integer)

Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    
    gstrNameRepo = "FondoActivoFijoGrilla"
    
        Set frmReporte = New frmVisorReporte

        ReDim aReportParamS(4)
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
        aReportParamS(2) = Codigo_Tipo_Persona_Proveedor
       
       
         If chkVerSoloGastosVigentes.Value = vbChecked Then
         
            aReportParamS(3) = "X"
            
            If chkVerSoloGastosContabilizados.Value = vbChecked Then
            
                aReportParamS(4) = "X"
            Else
               aReportParamS(4) = " "
             End If
             
        Else
             aReportParamS(3) = " "
             aReportParamS(4) = "%"
        End If
       
          
    
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
       
 Dim adoPendientes As ADODB.Recordset
   
            If strEstado = Reg_Consulta Or strEstado = Reg_Edicion Then
                
                If chkVerSoloGastosVigentes.Value = 0 Then
                    MsgBox "Este registro ya esta anulado", vbOKOnly + vbCritical, Me.Caption
                    Exit Sub
                End If
                
                If tdgConsulta.Columns(11).Value = Valor_Indicador Then
                    MsgBox "No se puede anular un registro que ya ah sido contabilizado", vbOKOnly + vbCritical, Me.Caption
                    Exit Sub
                End If
                
                With adoComm
                
                    .CommandText = "SELECT  NumRegistro,DescripRegistro FROM RegistroCompra WHERE NumGasto='" & tdgConsulta.Columns(2).Value & "' and Estado='01'"
                    Set adoPendientes = adoComm.Execute
                    
                    If Not adoPendientes.EOF = True Then
                    
                          If MsgBox("Se procederá a eliminar el Activo Fijo ' " & tdgConsulta.Columns(2) & "' asociado al Número de Registro " & _
                          "'" & adoPendientes("NumRegistro") & "-" & adoPendientes("DescripRegistro") & " ' de la Opcion Comprobante de Pago del Modulo de Contabilidad." & _
                          vbNewLine & vbNewLine & "Seguro de continuar ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then

                                    adoComm.CommandText = "UPDATE FondoGasto SET IndVigente='" & Valor_Caracter & "', IndFinMes = 'X' " & _
                                    "WHERE NumGasto=" & CInt(tdgConsulta.Columns(2)) & " AND CodCuenta='" & tdgConsulta.Columns(1) & "' AND " & _
                                    "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
                                    adoConn.Execute adoComm.CommandText
                                    
                                    '--------- Tabla Registro Compra
                                    
                                    adoComm.CommandText = "UPDATE RegistroCompra SET Estado = '02' WHERE " & _
                                        "NumRegistro=" & adoPendientes("NumRegistro") & " AND CodFondo='" & _
                                        strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
                                    adoConn.Execute adoComm.CommandText
                                    
                                    tabGasto.TabEnabled(0) = True
                                    tabGasto.TabEnabled(1) = False
                                    tabGasto.Tab = 0
                                    Call Buscar
                                    Exit Sub
                        End If
                    
                    Else
                    
                        If MsgBox("Se procederá a eliminar el Activo Fijo." & vbNewLine & vbNewLine & "Seguro de continuar ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then

                                    adoComm.CommandText = "UPDATE FondoGasto SET IndVigente='" & Valor_Caracter & "' " & _
                                    "WHERE NumGasto=" & CInt(tdgConsulta.Columns(2)) & " AND CodCuenta='" & tdgConsulta.Columns(1) & "' AND " & _
                                    "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
                                    adoConn.Execute adoComm.CommandText
                                    tabGasto.TabEnabled(0) = True
                                    tabGasto.TabEnabled(1) = False
                                    tabGasto.Tab = 0
                                    Call Buscar
                                    Exit Sub
                        End If
                    
                    End If
                    
                    
                End With
                
            End If
       
       
End Sub
Public Sub Modificar()

    Dim adoRegistro As New ADODB.Recordset
    

    If strEstado = Reg_Consulta Then
        
        If chkVerSoloGastosVigentes.Value = 0 Then
            MsgBox "No se puede modificar un registro anulado", vbOKOnly + vbCritical, Me.Caption
            Exit Sub
        End If
        
        If tdgConsulta.Columns(11).Value = Valor_Indicador Then
            MsgBox "No se puede modificar un registro que ya ah sido contabilizado", vbOKOnly + vbCritical, Me.Caption
            Exit Sub
        End If
        
        strEstado = Reg_Edicion
        
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        
        With tabGasto
            .TabEnabled(0) = False
            .TabEnabled(1) = True
            .Tab = 1
        End With
        
    End If
        
End Sub

Private Sub LlenarFormulario(strModo As String)

    Dim adoRegistro   As ADODB.Recordset
    Dim intRegistro As Integer
    
    Select Case strModo
        
        Case Reg_Adicion
            
            fraGastos(1).Caption = "Fondo : " & Trim(cboFondo.Text) & Space(1) & "-" & Space(1) '& "Tipo Gasto : " & Trim(cboTipoProvision.Text)
            lblAnalitica.Caption = "030-????????"
            txtDescripGasto.Text = Valor_Caracter
            txtMontoGasto.Text = "0"
            lblSaldoProvision.Caption = "0"
            'txtBaseCalculo.Text = "0"
            txtValorPeriodoPago.Text = "0"
            dtpFechaGasto.Value = gdatFechaActual
            dtpFechaGasto.Enabled = False
            dtpFechaInicio.Value = gdatFechaActual
            dtpFechaFin.Value = gdatFechaActual
            dtpFechaPrimerPago.Value = gdatFechaActual
            dtpFechaInicio.Enabled = True
            dtpFechaFin.Enabled = True
            chkNoIncluyeEnBalancePrecierre.Value = vbUnchecked
                        
            lblProveedor.Caption = Valor_Caracter
            lblCodProveedor.Caption = Valor_Caracter
            lblTipoDocID.Caption = Valor_Caracter
            lblNumDocID.Caption = Valor_Caracter
            
            Call CargarGastos
            
            cboGasto.ListIndex = -1
            If cboGasto.ListCount > 0 Then cboGasto.ListIndex = 0
            
            cboTipoGasto.ListIndex = -1
            'If cboTipoGasto.ListCount > 0 Then cboTipoGasto.ListIndex = 0
            
            intRegistro = ObtenerItemLista(arrEstado(), Valor_Indicador)
            If intRegistro >= 0 Then cboEstado.ListIndex = intRegistro
                                  
            'Montos
            cboMoneda.ListIndex = -1
            'cboTipoValor.ListIndex = -1
            'cboTipoTasa.ListIndex = -1
            'cboPeriodoTasa.ListIndex = -1
            
            'Fechas
            cboModalidadPago.ListIndex = -1
            cboTipoPago.ListIndex = -1
            cboPeriodoPago.ListIndex = -1
            cboTipoDesplazamiento.ListIndex = -1
            
            'Condiciones Cobtables y Tributarias
            cboTipoDevengo.ListIndex = -1
            cboAplicacionDevengo.ListIndex = -1
            cboFrecuenciaDevengo.ListIndex = -1
            cboCreditoFiscal.ListIndex = -1
                        
            intRegistro = ObtenerItemLista(arrTipoDevengo(), Codigo_Tipo_Devengo_Provision)
            If intRegistro >= 0 Then cboTipoDevengo.ListIndex = intRegistro
                        
            intRegistro = ObtenerItemLista(arrAplicacionDevengo(), Codigo_Aplica_Devengo_Inmediata)
            If intRegistro >= 0 Then cboAplicacionDevengo.ListIndex = intRegistro
                                          
            intRegistro = ObtenerItemLista(arrCreditoFiscal(), Codigo_Tipo_Credito_RentaGravada)
            If intRegistro >= 0 Then cboCreditoFiscal.ListIndex = intRegistro
                        
            intRegistro = ObtenerItemLista(arrMoneda(), strCodMoneda)
            If intRegistro >= 0 Then cboMoneda.ListIndex = intRegistro
            
            intRegistro = ObtenerItemLista(arrTipoGasto(), Codigo_Tipo_Gasto_Unico)
            If intRegistro >= 0 Then cboTipoGasto.ListIndex = intRegistro
            
            intRegistro = ObtenerItemLista(arrModalidadPago(), Codigo_Modalidad_Pago_Vencimiento)
            If intRegistro >= 0 Then cboModalidadPago.ListIndex = intRegistro
            
            intRegistro = ObtenerItemLista(arrTipoPago(), Codigo_Tipo_Pago_Unico)
            If intRegistro >= 0 Then cboTipoPago.ListIndex = intRegistro
            
            intRegistro = ObtenerItemLista(arrTipoDesplazamiento(), Tipo_Desplazamiento_Ningun_Desplazamiento)
            If intRegistro >= 0 Then cboTipoDesplazamiento.ListIndex = intRegistro
           
'            intRegistro = ObtenerItemLista(arrTipoValor(), Codigo_Tipo_Costo_Monto)
'            If intRegistro >= 0 Then cboTipoValor.ListIndex = intRegistro
           
'            intRegistro = ObtenerItemLista(arrBaseCalculo(), Codigo_Base_30_360)
'            If intRegistro >= 0 Then cboBaseCalculo.ListIndex = intRegistro
                      
            'lblMagnitud.Caption = lblMoneda.Caption
            
            cboGasto.SetFocus
                        
        Case Reg_Edicion
        
            Call CargarGastos
            
            Set adoRegistro = New ADODB.Recordset
            
            If tdgConsulta.AllowRowSelect = True Then
                adoComm.CommandText = "SELECT FG.CodBaseAnual,FG.NumGasto,FG.CodProveedor,FG.IndNoIncluyeEnBalancePreCierre, FG.IndVigente,FG.CodTipoGasto,FG.CodTipoValor,FG.PorcenGasto,FG.MontoBaseCalculo,FG.CodCuenta,FG.CodFile,FG.CodAnalitica,FG.DescripGasto,FG.CodCreditoFiscal,FG.CodMoneda,FG.CodFrecuenciaDevengo,FG.CodPeriodoPago,FG.CodTipoPago," & _
                    "FG.MontoGasto,FG.MontoDevengo,FG.FechaDefinicion,FG.FechaInicial,FG.FechaFinal,FG.FechaPrimerPago,FG.CodTipoDesplazamiento,FG.IndFinMes,FG.CodAplicacionDevengo,FG.CodTipoTasa,FG.CodPeriodoTasa,FG.CodTipoDevengo,FG.CodModalidadPago,FG.FechaPrimerPago,FG.IndFinMes," & _
                    "AP.DescripParametro TipoIdentidad,INP.DescripPersona,INP.NumIdentidad " & _
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
                fraGastos(1).Caption = "Fondo : " & Trim(cboFondo.Text) & Space(1) & "-" & Space(1) ' & "Tipo Gasto : " & Trim(cboTipoProvision.Text)
                
                intSecuencialGasto = adoRegistro("NumGasto")
                
                strCodFile = Trim(adoRegistro("CodFile"))
                lblAnalitica.Caption = Trim(adoRegistro("CodFile")) & "-" & Trim(adoRegistro("CodAnalitica"))
                txtDescripGasto.Text = Trim(adoRegistro("DescripGasto"))
            
                intRegistro = ObtenerItemLista(arrGasto(), adoRegistro("CodCuenta"))
                If intRegistro >= 0 Then cboGasto.ListIndex = intRegistro
            
                strCodTipoGasto = CStr(adoRegistro("CodTipoGasto"))
                
                lblProveedor.Caption = Trim(adoRegistro("DescripPersona"))
                lblTipoDocID.Caption = adoRegistro("TipoIdentidad")
                lblNumDocID.Caption = adoRegistro("NumIdentidad")
                lblCodProveedor.Caption = adoRegistro("CodProveedor")
               
                dtpFechaGasto.Value = adoRegistro("FechaDefinicion")
                dtpFechaInicio.Value = adoRegistro("FechaInicial")
                dtpFechaFin.Value = adoRegistro("FechaFinal")
                                             
'                If dtpFechaFin.Value < gdatFechaActual Then
'                    dtpFechaInicio.Enabled = False
'                    dtpFechaFin.Enabled = False
'                Else
'                    dtpFechaInicio.Enabled = True
'                    dtpFechaFin.Enabled = True
'                End If
               
               
                intRegistro = ObtenerItemLista(arrTipoGasto(), adoRegistro("CodTipoGasto"))
                If intRegistro >= 0 Then cboTipoGasto.ListIndex = intRegistro
                                   
                intRegistro = ObtenerItemLista(arrEstado(), adoRegistro("IndVigente"))
                If intRegistro >= 0 Then cboEstado.ListIndex = intRegistro
                
                lblSaldoProvision.Caption = CStr(adoRegistro("MontoDevengo"))
                                   
                'Montos
                intRegistro = ObtenerItemLista(arrMoneda(), adoRegistro("CodMoneda"))
                If intRegistro >= 0 Then cboMoneda.ListIndex = intRegistro
                                   
'                intRegistro = ObtenerItemLista(arrTipoValor(), adoRegistro("CodTipoValor"))
'                If intRegistro >= 0 Then cboTipoValor.ListIndex = intRegistro
                                   
'                intRegistro = ObtenerItemLista(arrTipoTasa(), adoRegistro("CodTipoTasa"))
'                If intRegistro >= 0 Then cboTipoTasa.ListIndex = intRegistro
                                   
'                intRegistro = ObtenerItemLista(arrPeriodoTasa(), adoRegistro("CodPeriodoTasa"))
'                If intRegistro >= 0 Then cboPeriodoTasa.ListIndex = intRegistro
                                   
                If adoRegistro("CodTipoValor") = Codigo_Tipo_Costo_Monto Then
                    txtMontoGasto.Text = CStr(adoRegistro("MontoGasto"))
                End If
                
                If adoRegistro("CodTipoValor") = Codigo_Tipo_Costo_Porcentaje Then
                    txtMontoGasto.Text = CStr(adoRegistro("PorcenGasto"))
                End If
                
'                txtBaseCalculo.Text = CStr(adoRegistro("MontoBaseCalculo"))
                                   
'                intRegistro = ObtenerItemLista(arrBaseCalculo(), adoRegistro("CodBaseAnual"))
'                If intRegistro >= 0 Then cboBaseCalculo.ListIndex = intRegistro
                                   
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
                                   
                If adoRegistro("IndFinMes") = Valor_Indicador Then
                    chkIndicadorPagoFinMes.Value = vbChecked
                Else
                    chkIndicadorPagoFinMes.Value = vbUnchecked
                End If
                                   
                'Condiciones Contables y Tributarias
                intRegistro = ObtenerItemLista(arrTipoDevengo(), adoRegistro("CodTipoDevengo"))
                If intRegistro >= 0 Then cboTipoDevengo.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrAplicacionDevengo(), adoRegistro("CodAplicacionDevengo"))
                If intRegistro >= 0 Then cboAplicacionDevengo.ListIndex = intRegistro
                                                                        
                intRegistro = ObtenerItemLista(arrFrecuenciaDevengo(), adoRegistro("CodFrecuenciaDevengo"))
                If intRegistro >= 0 Then cboFrecuenciaDevengo.ListIndex = intRegistro
                            
                intRegistro = ObtenerItemLista(arrCreditoFiscal(), adoRegistro("CodCreditoFiscal"))
                If intRegistro >= 0 Then cboCreditoFiscal.ListIndex = intRegistro
                
                If adoRegistro("IndNoIncluyeEnBalancePreCierre") = Valor_Indicador Then
                    chkNoIncluyeEnBalancePrecierre.Value = vbChecked
                Else
                    chkNoIncluyeEnBalancePrecierre.Value = vbUnchecked
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
        .TabEnabled(1) = True
        .Tab = 1
    End With
      
End Sub
Private Function TodoOK()

    TodoOK = False
    
    If Trim(strCodGasto) = Valor_Caracter Then
        MsgBox "Debe Seleccionar el Concepto.", vbCritical
        cboGasto.SetFocus
        Exit Function
    End If
      
'    If Trim(strCodCreditoFiscal) = Valor_Caracter Then
'        MsgBox "Debe Seleccionar el Tipo de Crédito Fiscal.", vbCritical
'        cboCreditoFiscal.SetFocus
'        Exit Function
'    End If
    
    If Trim(txtDescripGasto.Text) = Valor_Caracter Then
        MsgBox "Debe Ingresar la Descripción del Activo Fijo.", vbCritical
        txtDescripGasto.SetFocus
        Exit Function
    End If
    
    If cboMoneda.ListIndex = -1 Then
        MsgBox "Debe Ingresar la Moneda del Activo Fijo.", vbCritical
        cboMoneda.SetFocus
        Exit Function
    End If
                
    If CDec(txtMontoGasto.Text) < 0 Then
        MsgBox "El Valor del Gasto no Puede Ser Menor que 0.", vbCritical
        txtMontoGasto.SetFocus
        Exit Function
    End If
    
    If Trim(lblCodProveedor.Caption) = "" Then
        MsgBox "Debe Indicar el Proveedor del Activo Fijo.", vbCritical
        txtMontoGasto.SetFocus
        Exit Function
    End If
    
    If cboAplicacionDevengo.ListIndex = -1 Then
        MsgBox "Debe Seleccionar el Tipo de Aplicación de Devengo del Activo Fijo.", vbCritical
        cboAplicacionDevengo.SetFocus
        Exit Function
    End If
    
    If strCodAplicacionDevengo = Codigo_Aplica_Devengo_Periodica And cboFrecuenciaDevengo.ListIndex = -1 Then
        MsgBox "Debe Seleccionar la Frecuencia de Aplicación de Devengo del Activo Fijo.", vbCritical
        cboFrecuenciaDevengo.SetFocus
        Exit Function
    End If
    
     If txtMontoGasto.Text = "0.00" Then
        MsgBox "Debe Ingresar el monto del Activo Fijo.", vbCritical
        txtMontoGasto.SetFocus
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


Private Sub tdgConsulta_HeadClick(ByVal ColIndex As Integer)
    Static numColindex As Integer

    tdgConsulta.Splits(0).Columns(numColindex).HeadingStyle.ForegroundPicture = Null

    Call OrdenarDBGrid(ColIndex, adoConsulta, tdgConsulta)
    
    numColindex = ColIndex
End Sub

'Private Sub txtBaseCalculo_Change()
'
'    Call FormatoCajaTexto(txtBaseCalculo, Decimales_Monto)
'
'End Sub

'Private Sub txtBaseCalculo_KeyPress(KeyAscii As Integer)
'
'    Call ValidaCajaTexto(KeyAscii, "M", txtBaseCalculo, Decimales_Monto)
'
'End Sub

Private Sub txtMontoGasto_Change()

'    If strCodTipoValor = Codigo_Tipo_Costo_Monto Then
'        Call FormatoCajaTexto(txtMontoGasto, Decimales_Monto)
'    Else
'        Call FormatoCajaTexto(txtMontoGasto, Decimales_Tasa)
'    End If
    
End Sub


'Private Sub txtMontoGasto_KeyPress(KeyAscii As Integer)
'
''    If strCodTipoValor = Codigo_Tipo_Costo_Monto Then
''        Call ValidaCajaTexto(KeyAscii, "M", txtMontoGasto, Decimales_Monto)
''    Else
''        Call ValidaCajaTexto(KeyAscii, "M", txtMontoGasto, Decimales_Tasa)
''    End If
'
'End Sub

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
   
    With adoComm
        Set adoConsulta = New ADODB.Recordset

        '*** Obtener el número de días del peridodo de pago ***
        .CommandText = "SELECT ValorParametro FROM AuxiliarParametro WHERE CodTipoParametro='TIPFRE' AND CodParametro='" & strCodPeriodoPago & "'"
        Set adoConsulta = .Execute

        If Not adoConsulta.EOF Then
            intNumDiasPeriodo = CInt(adoConsulta("ValorParametro")) '*** Días del periodo  ***
            intNumMesesPeriodo = CInt(intNumDiasPeriodo / 30)       '*** Meses del periodo ***
            'intNumPeriodosAnual = CInt(12 / intNumMesesPeriodo)     '*** Periodos al año   ***
        End If
        adoConsulta.Close: Set adoConsulta = Nothing
    
    End With
    
    Do While datFechaIniCupon <= dtpFechaFin.Value
        intNumSecuencial = intNumSecuencial + 1
       
        '*** Fecha de corte del primer cupón ***
        '*** Mes calendario ***
        If intNumSecuencial > 1 Then
            datFechaFinCupon = DateAdd("m", intNumMesesPeriodo, datFechaIniCupon) - 1
            If chkIndicadorPagoFinMes.Value = vbChecked Then
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
            strFechaPago = Convertyyyymmdd(DesplazamientoDiaUtil(datFechaFinCupon, strCodTipoDesplazamiento))

            '*** Grabar en temporal ***
            Call GrabarFechaCorteTmp
            
            blnCupon = False: Exit Do
        End If
        
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
        strFechaPago = Convertyyyymmdd(DesplazamientoDiaUtil(datFechaFinCupon, strCodTipoDesplazamiento))
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

Private Sub AutoAjustarGrillas()
    
    Dim i As Integer
    
    If Not adoConsulta.EOF Then
        If adoConsulta.RecordCount > 0 Then
            For i = 0 To tdgConsulta.Columns.Count - 1
                tdgConsulta.Columns(i).AutoSize
            Next
            
            tdgConsulta.Columns(11).AutoSize
        End If
    End If
    
    tdgConsulta.Refresh

End Sub

