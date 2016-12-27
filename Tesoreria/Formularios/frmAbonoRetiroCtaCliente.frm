VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmAbonoRetiroCtaCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Abono y Retiro de Cuenta de Cliente"
   ClientHeight    =   9435
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9435
   ScaleWidth      =   11040
   Begin TAMControls2.ucBotonEdicion2 cmdIdentificar 
      Height          =   735
      Left            =   270
      TabIndex        =   59
      Top             =   8670
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
      Caption0        =   "&Identificar"
      Tag0            =   "3"
      Visible0        =   0   'False
      ToolTipText0    =   "Nuevo"
      UserControlWidth=   1200
   End
   Begin VB.CommandButton cmdReservar 
      Caption         =   "Reversar"
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
      Left            =   6405
      Picture         =   "frmAbonoRetiroCtaCliente.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   8670
      Width           =   1200
   End
   Begin TAMControls2.ucBotonEdicion2 cmdAccion2 
      Height          =   735
      Left            =   7980
      TabIndex        =   50
      Top             =   8670
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   1296
      Buttons         =   2
      Caption0        =   "&Imprimir"
      Tag0            =   "6"
      Visible0        =   0   'False
      ToolTipText0    =   "Imprimir"
      Caption1        =   "&Salir"
      Tag1            =   "9"
      Visible1        =   0   'False
      ToolTipText1    =   "Salir"
      UserControlWidth=   2700
   End
   Begin TAMControls2.ucBotonEdicion2 cmdOpcion 
      Height          =   735
      Left            =   270
      TabIndex        =   49
      Top             =   8670
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   1296
      Buttons         =   4
      Caption0        =   "&Nuevo"
      Tag0            =   "0"
      Visible0        =   0   'False
      ToolTipText0    =   "Nuevo"
      Caption1        =   "&Modificar"
      Tag1            =   "1"
      Visible1        =   0   'False
      ToolTipText1    =   "Modificar"
      Caption2        =   "&Buscar"
      Tag2            =   "5"
      Visible2        =   0   'False
      ToolTipText2    =   "Buscar"
      Caption3        =   "&Anular"
      Tag3            =   "4"
      Visible3        =   0   'False
      ToolTipText3    =   "Anular"
      UserControlWidth=   5700
   End
   Begin TabDlg.SSTab tabPagos 
      Height          =   8445
      Left            =   60
      TabIndex        =   2
      Top             =   150
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   14896
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
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
      TabPicture(0)   =   "frmAbonoRetiroCtaCliente.frx":0625
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "tdgConsulta"
      Tab(0).Control(1)=   "fraCriterio"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "frmAbonoRetiroCtaCliente.frx":0641
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraDatos"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdAccion"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   7200
         TabIndex        =   57
         Top             =   7560
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
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmAbonoRetiroCtaCliente.frx":065D
         Height          =   4185
         Left            =   -74760
         OleObjectBlob   =   "frmAbonoRetiroCtaCliente.frx":0677
         TabIndex        =   11
         Top             =   4020
         Width           =   10245
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
         Height          =   3345
         Left            =   -74760
         TabIndex        =   6
         Top             =   510
         Width           =   10245
         Begin VB.CommandButton cmdEnviarBackOffice 
            Caption         =   "&Procesar"
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
            Left            =   7920
            Picture         =   "frmAbonoRetiroCtaCliente.frx":7727
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   2160
            Width           =   1200
         End
         Begin VB.ComboBox cboEstadoOperacionBusqueda 
            Height          =   315
            Left            =   2130
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   1920
            Width           =   2745
         End
         Begin VB.TextBox txtCodParticipeBusqueda 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2130
            MaxLength       =   20
            TabIndex        =   14
            Top             =   1020
            Width           =   2280
         End
         Begin VB.CommandButton cmdBusqueda 
            Caption         =   "..."
            Height          =   315
            Left            =   4440
            TabIndex        =   13
            ToolTipText     =   "Búsqueda de Partícipe"
            Top             =   1020
            Width           =   390
         End
         Begin VB.ComboBox cboFondo 
            Height          =   315
            Left            =   2130
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   540
            Width           =   7635
         End
         Begin MSComCtl2.DTPicker dtpFechaOperacionDesde 
            Height          =   315
            Left            =   2010
            TabIndex        =   51
            Top             =   2880
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
            Format          =   206766081
            CurrentDate     =   38785
         End
         Begin MSComCtl2.DTPicker dtpFechaOperacionHasta 
            Height          =   315
            Left            =   5160
            TabIndex        =   52
            Top             =   2880
            Width           =   1515
            _ExtentX        =   2672
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
            Format          =   206766081
            CurrentDate     =   38785
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
            Index           =   25
            Left            =   4080
            TabIndex        =   55
            Top             =   2880
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
            Index           =   42
            Left            =   270
            TabIndex        =   54
            Top             =   2880
            Width           =   555
         End
         Begin VB.Line Line2 
            X1              =   240
            X2              =   6840
            Y1              =   2760
            Y2              =   2760
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Operacion"
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
            Left            =   3000
            TabIndex        =   53
            Top             =   2520
            Width           =   1470
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Estado Operación"
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
            Left            =   390
            TabIndex        =   29
            Top             =   1950
            Width           =   1515
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
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
            Left            =   360
            TabIndex        =   16
            Top             =   1050
            Width           =   600
         End
         Begin VB.Label lblDescripParticipeBusqueda 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2130
            TabIndex        =   15
            Top             =   1470
            Width           =   7635
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
            Index           =   2
            Left            =   360
            TabIndex        =   7
            Top             =   570
            Width           =   540
         End
      End
      Begin VB.Frame fraDatos 
         Caption         =   "Datos"
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
         Height          =   6825
         Left            =   390
         TabIndex        =   3
         Top             =   600
         Width           =   10245
         Begin VB.ComboBox cboFormaPago 
            Height          =   315
            Left            =   6960
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   2310
            Width           =   2655
         End
         Begin VB.TextBox txtNroCheque 
            Height          =   285
            Left            =   6960
            TabIndex        =   42
            Top             =   2730
            Visible         =   0   'False
            Width           =   2625
         End
         Begin VB.ComboBox cboCuentasControl 
            Height          =   315
            Left            =   2310
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   6210
            Width           =   7410
         End
         Begin VB.CheckBox chkSinIdentificar 
            Caption         =   "Participe sin Identificar"
            Height          =   375
            Left            =   4980
            TabIndex        =   39
            Top             =   1530
            Width           =   2775
         End
         Begin VB.ComboBox cboCuentas 
            Height          =   315
            Left            =   2310
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   3750
            Width           =   7350
         End
         Begin VB.TextBox txtDescripObservaciones 
            Height          =   975
            Left            =   2310
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   27
            Top             =   5070
            Width           =   7425
         End
         Begin VB.CommandButton cmdBusquedaParticipe 
            Caption         =   "..."
            Height          =   315
            Left            =   4410
            TabIndex        =   25
            ToolTipText     =   "Búsqueda de Partícipe"
            Top             =   1530
            Width           =   390
         End
         Begin VB.TextBox txtCodParticipe 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2310
            MaxLength       =   20
            TabIndex        =   24
            Top             =   1530
            Width           =   2040
         End
         Begin VB.ComboBox cboTipoMov 
            Height          =   315
            Left            =   2310
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   1110
            Width           =   3135
         End
         Begin VB.ComboBox cboMoneda 
            Height          =   315
            Left            =   2310
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   3270
            Width           =   3105
         End
         Begin VB.TextBox txtNroVoucher 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2310
            MaxLength       =   12
            TabIndex        =   17
            Text            =   " "
            Top             =   4620
            Width           =   1830
         End
         Begin MSComCtl2.DTPicker dtpFechaActual 
            Height          =   315
            Left            =   2310
            TabIndex        =   1
            Top             =   2340
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   206766081
            CurrentDate     =   38949
         End
         Begin VB.TextBox txtMontoPago 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2310
            MaxLength       =   12
            TabIndex        =   4
            Text            =   " "
            Top             =   4170
            Width           =   1800
         End
         Begin MSComCtl2.DTPicker dtpFechaObligacion 
            Height          =   315
            Left            =   2310
            TabIndex        =   31
            Top             =   2760
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   556
            _Version        =   393216
            Format          =   206766081
            CurrentDate     =   38949
         End
         Begin VB.Label lblMonedaOrigen 
            AutoSize        =   -1  'True
            Caption         =   "(USD)"
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
            Left            =   9630
            TabIndex        =   48
            Top             =   1200
            Width           =   525
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblSaldoCliente 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7620
            TabIndex        =   47
            Top             =   1170
            Width           =   1995
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Disponible"
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
            Left            =   6060
            TabIndex        =   46
            Top             =   1170
            Width           =   1440
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Forma de Pago"
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
            Index           =   16
            Left            =   5370
            TabIndex        =   45
            Top             =   2340
            Width           =   1365
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Nro Cheque"
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
            Index           =   17
            Left            =   5400
            TabIndex        =   44
            Top             =   2730
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta Control"
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
            Index           =   15
            Left            =   360
            TabIndex        =   41
            Top             =   6240
            Width           =   1275
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Disponible"
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
            Left            =   5940
            TabIndex        =   38
            Top             =   4230
            Width           =   1440
         End
         Begin VB.Label lblSaldoCuenta 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7620
            TabIndex        =   37
            Top             =   4170
            Width           =   1995
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta"
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
            Left            =   420
            TabIndex        =   36
            Top             =   3810
            Width           =   615
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Nro. Operación"
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
            Left            =   6150
            TabIndex        =   34
            Top             =   300
            Width           =   1290
         End
         Begin VB.Label lblNumOperacion 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "GENERADO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   7770
            TabIndex        =   33
            Top             =   270
            Width           =   1845
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Obligacion"
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
            Left            =   330
            TabIndex        =   32
            Top             =   2820
            Width           =   1500
         End
         Begin VB.Label lblSignoMoneda 
            Caption         =   "PEN"
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
            Left            =   4260
            TabIndex        =   30
            Top             =   4200
            Width           =   465
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comitente"
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
            Left            =   330
            TabIndex        =   26
            Top             =   1590
            Width           =   855
         End
         Begin VB.Line Line1 
            X1              =   360
            X2              =   9570
            Y1              =   3150
            Y2              =   3150
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Movimiento"
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
            Left            =   330
            TabIndex        =   23
            Top             =   1170
            Width           =   1410
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones"
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
            Left            =   360
            TabIndex        =   21
            Top             =   5100
            Width           =   1275
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
            Height          =   240
            Index           =   7
            Left            =   360
            TabIndex        =   20
            Top             =   3330
            Width           =   690
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Nro. Voucher"
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
            Left            =   360
            TabIndex        =   18
            Top             =   4650
            Width           =   1140
         End
         Begin VB.Label lblDescripParticipe 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2310
            TabIndex        =   12
            Top             =   1920
            Width           =   7245
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto"
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
            Left            =   360
            TabIndex        =   10
            Top             =   4230
            Width           =   540
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
            Index           =   3
            Left            =   360
            TabIndex        =   9
            Top             =   750
            Width           =   540
         End
         Begin VB.Label lblDescripFondo 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2310
            TabIndex        =   8
            Top             =   690
            Width           =   7305
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Registro"
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
            Left            =   330
            TabIndex        =   5
            Top             =   2400
            Width           =   1305
         End
      End
   End
End
Attribute VB_Name = "frmAbonoRetiroCtaCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strCodFondo         As String, strCodParticipe      As String
Dim strEstado           As String, strSQL               As String
Dim strTipMov           As String
Dim curMontoEmitido     As Currency, strNroVocuher      As String
Dim arrMoneda()         As String, strCodMoneda         As String, strSignoMoneda   As String, strCodSignoMoneda As String
Dim arrCuenta()         As String, strCodCuenta         As String
Dim arrCuentas()        As String
Dim arrCuentasControl() As String
Dim arrFormaPago()      As String
Dim arrTipMov()         As String, strEstadoOperacion   As String, strFormPago  As String
Dim strCodFile          As String, strCodBanco  As String, strCodAnalitica As String
Dim adoConsulta         As ADODB.Recordset
Dim adoRegistroAux      As ADODB.Recordset
Dim arrEstadoOperacionBusqueda()  As String
Dim strEstadoOperacionBusqueda As String
Dim strCodParticipeBusqueda As String
Dim strNumOperacion As String


Private Sub cboCuentas_Click()

    Dim curSaldoCuenta As Double, strFecha As String, strFechaMas1Dia As String

    
    strCodFile = Valor_Caracter: strCodAnalitica = Valor_Caracter
    strCodBanco = Valor_Caracter: strCodCuenta = Valor_Caracter
    curSaldoCuenta = 0
    
    If cboCuentas.ListIndex < 0 Then Exit Sub
   
    strCodFile = Left(Trim(arrCuentas(cboCuentas.ListIndex)), 3)
    strCodAnalitica = Mid(Trim(arrCuentas(cboCuentas.ListIndex)), 4, 8)
    strCodBanco = Mid(Trim(arrCuentas(cboCuentas.ListIndex)), 12, 8)
    strCodCuenta = Trim(Right(arrCuentas(cboCuentas.ListIndex), 10))
      
    strFecha = gstrFechaActual  'Convertyyyymmdd(dtpFechaContable.Value)
    strFechaMas1Dia = Convertyyyymmdd(DateAdd("d", 1, gdatFechaActual))  'dtpFechaContable.Value))
    
    '*** Obtener los saldos de la cuenta ***
    lblSaldoCuenta.Caption = "0"
    curSaldoCuenta = ObtenerSaldoFinalCuenta(strCodFondo, gstrCodAdministradora, strCodFile, strCodAnalitica, strFecha, strFechaMas1Dia, strCodCuenta, strCodMoneda) ', strCodMonedaFondo)
    lblSaldoCuenta.Caption = curSaldoCuenta

    
    'Call ObtenerSaldos

End Sub

Private Sub ObtenerSaldos()

    Dim adoTemporal As ADODB.Recordset
    Dim strFecha    As String, strFechaMas1Dia  As String
    
    strFecha = gstrFechaActual  'Convertyyyymmdd(dtpFechaContable.Value)
    strFechaMas1Dia = Convertyyyymmdd(DateAdd("d", 1, gdatFechaActual))  'dtpFechaContable.Value))
    
    Set adoTemporal = New ADODB.Recordset
    With adoComm
        .CommandText = "{ call up_ACObtenerSaldoCuentaContable('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
            strCodFile & "','" & strCodAnalitica & "','" & strFecha & "','" & strFechaMas1Dia & "','" & _
            strCodCuenta & "','" & strCodMoneda & "') }"
            
        Set adoTemporal = .Execute
        
        If Not adoTemporal.EOF Then
            lblSaldoCuenta.Caption = CStr(adoTemporal("SaldoCuenta"))
        Else
            lblSaldoCuenta.Caption = "0"
        End If
        adoTemporal.Close: Set adoTemporal = Nothing
    End With
    
End Sub

Private Sub cboEstadoOperacionBusqueda_Click()

    strEstadoOperacionBusqueda = Valor_Caracter
    If cboEstadoOperacionBusqueda.ListIndex < 0 Then Exit Sub
    strEstadoOperacionBusqueda = Trim(arrEstadoOperacionBusqueda(cboEstadoOperacionBusqueda.ListIndex))

    If strEstadoOperacionBusqueda = "04" Then
        cmdOpcion.Visible = False
        cmdIdentificar.Visible = True
               
    Else
        cmdOpcion.Visible = True
        cmdIdentificar.Visible = False
                
    End If
    Call Buscar

End Sub


Private Sub cboFondo_Click()

    Dim adoRegistro As ADODB.Recordset
    
    strCodFondo = Valor_Caracter
    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(garrFondo(cboFondo.ListIndex))
    
    With adoComm
        '*** Fecha Vigente, Valor Cuota, Valor Cuota T-1, Moneda y Cantidad Inicial de Cuotas del Fondo ***
        .CommandText = "{ call up_ACSelDatosParametro(23,'" & strCodFondo & "','" & gstrCodAdministradora & "') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            dtpFechaActual.Value = adoRegistro("FechaCuota")
            dtpFechaObligacion.Value = adoRegistro("FechaCuota")
        Else
            MsgBox "Periodo contable no vigente ! Debe aperturar primero un periodo contable!", vbExclamation + vbOKOnly, Me.Caption
            Exit Sub
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
        
    End With

End Sub

Private Sub cboFormaPago_Click()

     strFormPago = Valor_Caracter
     If cboFormaPago.ListIndex < 0 Then Exit Sub
     strFormPago = Trim(arrFormaPago(cboFormaPago.ListIndex))

        Select Case strFormPago
        
        Case Codigo_FormaPago_Cheque
        
            txtNroCheque.Visible = True
            lblDescrip(17).Visible = True
       
        Case Else
        
            txtNroCheque.Visible = False
            lblDescrip(17).Visible = False

        End Select

End Sub

Private Sub cboMoneda_Click()
    strCodMoneda = Valor_Caracter: strSignoMoneda = Valor_Caracter
    If cboMoneda.ListIndex < 0 Then Exit Sub
    
    strCodMoneda = Trim(arrMoneda(cboMoneda.ListIndex))
    
    strSignoMoneda = ObtenerSignoMoneda(strCodMoneda)
    strCodSignoMoneda = ObtenerCodSignoMoneda(strCodMoneda)
    
    lblSignoMoneda.Caption = strCodSignoMoneda
     
     lblMonedaOrigen(0).Caption = ObtenerSignoMoneda(strCodMoneda)
     
     
    'Call ObtenerSaldoCliente
    Call CargarCuentasBancarias

    
End Sub


Private Sub CargarCuentasBancarias()

    Dim strSQL As String
        
    strSQL = "SELECT (CodFile + CodAnalitica + CodBanco + CodCuentaActivo) CODIGO,(RTRIM(DescripCuenta) + SPACE(1) + NumCuenta) DESCRIP FROM BancoCuenta " & _
            "WHERE CodMoneda='" & strCodMoneda & "' AND IndVigente='X' AND " & _
            "CodFondo='" & gstrCodFondoContable & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
    CargarControlLista strSQL, cboCuentas, arrCuentas(), Sel_Defecto
    'cboCtaAhorro.Visible = True: cboCtaCte.Visible = False
    If cboCuentas.ListCount > 0 Then cboCuentas.ListIndex = 0
    
    
    If chkSinIdentificar.Value = 1 Or tdgConsulta.Columns("EstadoOperacion") = Codigo_Movimiento_Deposito_No_Identificado Then
        cboCuentasControl.Visible = True
        lblDescrip(15).Visible = True
        strSQL = "SELECT (CodFile + CodAnalitica + CodBanco + CodCuentaActivo) CODIGO,(RTRIM(DescripCuenta) + SPACE(1) + NumCuenta) DESCRIP FROM BancoCuenta " & _
            "WHERE CodMoneda='" & strCodMoneda & "' AND IndVigente='X' AND CodCuentaActivo like '18911%' AND " & _
            "CodFondo='" & gstrCodFondoContable & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
        CargarControlLista strSQL, cboCuentasControl, arrCuentasControl(), Valor_Caracter
        If cboCuentasControl.ListCount > 0 Then cboCuentasControl.ListIndex = 0
    Else
        cboCuentasControl.Visible = False
        lblDescrip(15).Visible = False
    End If
        
        
End Sub

Private Sub cboTipoMov_Click()
     
     strTipMov = Valor_Caracter
     If cboTipoMov.ListIndex < 0 Then Exit Sub
     strTipMov = Trim(arrTipMov(cboTipoMov.ListIndex))
     
    If strTipMov = Codigo_Movimiento_Retiro Then
    
        lblDescrip(16).Visible = True
        lblDescrip(17).Visible = True
        cboFormaPago.Visible = True
        txtNroCheque.Visible = True
    
     '*** Forma de Pago ***
     
     strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='MEDPAG' ORDER BY DescripParametro"
     CargarControlLista strSQL, cboFormaPago, arrFormaPago(), Valor_Caracter
     If cboFormaPago.ListCount > 0 Then cboFormaPago.ListIndex = 0
     
    Else
    
        lblDescrip(16).Visible = False
        lblDescrip(17).Visible = False
        cboFormaPago.Visible = False
        txtNroCheque.Visible = False
    End If
     
     

End Sub

Private Sub chkSinIdentificar_Click()

     Dim intRegistro                 As Integer

    Me.txtCodParticipe.Enabled = IIf(chkSinIdentificar.Value, False, True)
    Me.cmdBusquedaParticipe.Enabled = IIf(chkSinIdentificar.Value, False, True)
    If chkSinIdentificar.Value = 1 Then
        strCodParticipe = ""
        Me.lblDescripParticipe.Caption = ""
        Me.txtCodParticipe.Text = ""
        'Verificar que solo se realizen  Abonos
        intRegistro = ObtenerItemLista(arrTipMov(), Codigo_Movimiento_Deposito)
        If intRegistro >= 0 Then cboTipoMov.ListIndex = intRegistro
    End If
    
    cboTipoMov.Enabled = Not cboTipoMov.Enabled
    
    Call cboMoneda_Click
End Sub

Private Sub cmdBusqueda_Click()

    gstrFormulario = "frmAbonoRetiroCtaCliente"
    
    frmBusquedaParticipe.Show vbModal
    
    If Trim(gstrCodParticipe) <> "" Then strCodParticipeBusqueda = gstrCodParticipe
    
    Me.txtCodParticipeBusqueda.SetFocus
    
End Sub



Private Sub cmdBusquedaParticipe_Click()

    
    gstrFormulario = Me.Name
    
    frmBusquedaParticipe.Show vbModal
    
    If Trim(gstrCodParticipe) <> "" Then strCodParticipe = gstrCodParticipe
    
    Me.txtCodParticipe.SetFocus
    

End Sub

Private Sub cmdEnviarBackOffice_Click()

    
    Dim intContador                 As Integer
    Dim intRegistro                 As Integer
    Dim strTesoreriaOperacionXML    As String
    Dim objTesoreriaOperacionXML    As DOMDocument60
    Dim strFechaGrabar              As String
    Dim strMsgError                 As String
    Dim strTipoOperacion            As String
    Dim adoRegistro                 As ADODB.Recordset
        
        
    If TodoOkBackOffice() Then
        '*** Realizar proceso de contabilización ***
        If MsgBox("Datos correctos. ¿ Procedemos a enviar estas operaciones a Backoffice de Tesoreria?", vbQuestion + vbYesNo, "Observación") = vbNo Then Exit Sub
    
        intContador = tdgConsulta.SelBookmarks.Count - 1
               
        strFechaGrabar = Convertyyyymmdd(dtpFechaActual.Value) & Space(1) & Format(Time, "hh:mm")
                   
        Call ConfiguraRecordsetAuxiliar
            
        'adoRegistroAux.Open
               
        Set adoRegistro = New ADODB.Recordset
        
        With adoComm
        
            Set objTesoreriaOperacionXML = Nothing
            strTesoreriaOperacionXML = ""
                      
                      
            For intRegistro = 0 To intContador
                
                adoConsulta.MoveFirst
                
                adoConsulta.Move CLng(tdgConsulta.SelBookmarks(intRegistro) - 1), 0
                                
                tdgConsulta.Refresh
                                
                adoRegistroAux.AddNew
                adoRegistroAux.Fields("CodFondo") = strCodFondo
                adoRegistroAux.Fields("CodAdministradora") = gstrCodAdministradora
                adoRegistroAux.Fields("NumOperacion") = tdgConsulta.Columns("NumOperacion")
                
                
                strTipoOperacion = tdgConsulta.Columns("TipoOperacion")
          
            Next
           
            Call XMLADORecordset(objTesoreriaOperacionXML, "TesoreriaOperacion", "Operacion", adoRegistroAux, strMsgError)
            strTesoreriaOperacionXML = objTesoreriaOperacionXML.xml
                
            If strTipoOperacion <> "50" Then
                .CommandText = "{ call up_TEProcMovimientoFondoCliente('" & _
                    strCodFondo & "','" & gstrCodAdministradora & "','" & _
                    strFechaGrabar & "','" & _
                    strTesoreriaOperacionXML & "') }"
            Else
                
                .CommandText = "{ call up_TEProcAbonoNoIdentificado('" & _
                strCodFondo & "','" & gstrCodAdministradora & "','" & _
                strFechaGrabar & "','" & _
                strTesoreriaOperacionXML & "') }"
            End If
            
            .Execute .CommandText
                                             
        End With
        
        Me.MousePointer = vbDefault
        
        MsgBox Mensaje_Proceso_Exitoso, vbExclamation
        
        frmMainMdi.stbMdi.Panels(3).Text = "Acción"
        
        'cmdOpcion.Visible = True
        With tabPagos
            .TabEnabled(0) = True
            .Tab = 0
        End With
        
        Call Buscar
        tdgConsulta.ReBind
        Me.Refresh
    End If
    

End Sub


Public Sub Accion(nAccion As ButtonAction)
    
    Select Case nAccion
        
        Case vNew
            Call Adicionar
        Case vQuery
            Call Modificar
        Case vDelete
            Call Eliminar
        Case vSearch
            Call Buscar
        Case vPrint
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
    Call SubImprimir(1)
End Sub

Public Sub Eliminar()

    If strEstado = Reg_Consulta Or strEstado = Reg_Edicion Then
    
       If strEstadoOperacionBusqueda = "02" Then
            MsgBox "No se puede anular un registro ya procesado ", vbOKOnly + vbCritical, Me.Caption
            Exit Sub
        ElseIf strEstadoOperacionBusqueda = "03" Then
            MsgBox "Este registro ya esta anulado ", vbOKOnly + vbCritical, Me.Caption
            Exit Sub
        ElseIf strEstadoOperacionBusqueda = "05" Then
            MsgBox "No se puede anular un registro ya Reversado ", vbOKOnly + vbCritical, Me.Caption
            Exit Sub
        End If
    
    
        strEstadoOperacion = "03"
        If MsgBox("Se procederá a anular el Abono del Comitente." & vbNewLine & vbNewLine & "Seguro de continuar ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            adoComm.CommandText = "UPDATE TesoreriaOperacion SET EstadoOperacion='" & strEstadoOperacion & "' " & _
            "WHERE NumOperacion='" & tdgConsulta.Columns("NumOperacion") & "' AND " & _
            "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
            adoConn.Execute adoComm.CommandText
            tabPagos.TabEnabled(0) = True
            tabPagos.Tab = 0
            Call Buscar
            Exit Sub
        End If
    End If
End Sub
Public Sub Salir()

    Unload Me
    
End Sub
Public Sub Cancelar()

    cmdOpcion.Visible = True
    With tabPagos
        .TabEnabled(0) = True
        .Tab = 0
    End With
    strEstado = Reg_Consulta
    
End Sub
Public Sub Grabar()
                        
    Dim intAccion As Integer, lngNumError   As Integer
    Dim dblMontoPago As Double, strNumSolicitud As String
    Dim strDescripOperacion As String, strNumDocumento As String
    Dim strAccion As String, strDescripObservacion As String
    Dim strFechaObligacion As String, strFechaGrabar As String
    Dim strTipoOperacion As String
    Dim strBancoDestino As String, strBancoOrigen As String, strNumCheque As String
    Dim strTransitoriaDestino As String, strTransitoriaOrigen As String
    
    
    If strEstado = Reg_Consulta Then Exit Sub
            
   'On Error GoTo CtrlError
        
    If strEstadoOperacionBusqueda = "02" Or strEstadoOperacionBusqueda = "03" Or strEstadoOperacionBusqueda = "05" Then
            
        MsgBox "No puede Realizar esta Operacion", vbCritical, Me.Caption
        
    Else
    
    If TodoOk() Then
    
        strBancoOrigen = "'','',''"
        strBancoDestino = "'','',''"
        strTransitoriaOrigen = "'','',''"
        strTransitoriaDestino = "'','',''"
        
        strDescripOperacion = IIf(lblDescripParticipe.Caption = "", "Participe sin Identificar", lblDescripParticipe.Caption)
        If strTipMov = Codigo_Movimiento_Deposito Then
            strDescripOperacion = "Abono de Participe - " + strDescripOperacion
            strTipoOperacion = "32"
            strBancoDestino = "'" & strCodCuenta & "','" & strCodFile & "','" & strCodAnalitica & "'"
            
            If chkSinIdentificar.Value = 1 Then
                strCodFile = Left(Trim(arrCuentasControl(cboCuentasControl.ListIndex)), 3)
                strCodAnalitica = Mid(Trim(arrCuentasControl(cboCuentasControl.ListIndex)), 4, 8)
                'strCodBanco = Mid(Trim(arrCuentasControl(cboCuentasControl.ListIndex)), 12, 8)
                strCodCuenta = Trim(Right(arrCuentasControl(cboCuentasControl.ListIndex), 10))
                'strBancoOrigen = "'" & strCodCuenta & "','" & strCodFile & "','" & strCodAnalitica & "'"
                If tdgConsulta.Columns("EstadoOperacion") <> Codigo_Movimiento_Deposito_No_Identificado Then
                    strTransitoriaOrigen = "'" & strCodCuenta & "','" & strCodFile & "','" & strCodAnalitica & "'"
                    strTipoOperacion = "50"
                End If
            End If
            
            If tdgConsulta.Columns("EstadoOperacion") = Codigo_Movimiento_Deposito_No_Identificado Then
                strCodFile = Left(Trim(arrCuentasControl(cboCuentasControl.ListIndex)), 3)
                strCodAnalitica = Mid(Trim(arrCuentasControl(cboCuentasControl.ListIndex)), 4, 8)
                'strCodBanco = Mid(Trim(arrCuentasControl(cboCuentasControl.ListIndex)), 12, 8)
                strCodCuenta = Trim(Right(arrCuentasControl(cboCuentasControl.ListIndex), 10))
                strBancoDestino = "'" & strCodCuenta & "','" & strCodFile & "','" & strCodAnalitica & "'"
                
                
            End If
            
        ElseIf Codigo_Movimiento_Retiro Then
            strDescripOperacion = "Retiro de Participe - " + strDescripOperacion
            strTipoOperacion = "33"
            strBancoDestino = "'','',''"
            strBancoOrigen = "'" & strCodCuenta & "','" & strCodFile & "','" & strCodAnalitica & "'"
        End If
        
        dblMontoPago = IIf(strTipMov = Codigo_Movimiento_Retiro, CDbl(txtMontoPago.Text) * -1, CDbl(txtMontoPago.Text))

        strNumDocumento = Trim(txtNroVoucher.Text)
        
        strNumCheque = Trim(txtNroCheque.Text)
        
        strDescripObservacion = Trim(txtDescripObservaciones.Text)
    
        If strEstado = Reg_Adicion Then
            If strNumOperacion = Valor_Caracter Then
                strAccion = "I"
                strNumOperacion = Valor_Caracter
            Else
                strAccion = "D"
            End If
        End If
        
        If strEstado = Reg_Edicion Then
            strAccion = "U"
            strNumOperacion = lblNumOperacion.Caption
        End If
        
        strFechaGrabar = Convertyyyymmdd(dtpFechaActual.Value) & Space(1) & Format(Time, "hh:mm")
        strFechaObligacion = Convertyyyymmdd(dtpFechaObligacion.Value)
        
        strNumOperacion = lblNumOperacion.Caption
               
        Me.MousePointer = vbHourglass
        
        ''------------------------------
        ''---Registro en BD-------------
                
        With adoComm
                        
            
            .CommandText = "{ call up_TEManTesoreriaOperacion ('" & _
                     strCodFondo & "','" & gstrCodAdministradora & "','" & strNumOperacion & "'," & _
                     "'000', '00000000', '000'," & _
                     "'000','" & strTipoOperacion & "','','" & strTipMov & "','" & strCodParticipe & "','',''," & _
                     "'" & strDescripOperacion & "','" & strFechaGrabar & "','" & strFechaObligacion & "'," & _
                     "'19000101','" & strCodMoneda & "'," & dblMontoPago & "," & _
                     "'', 0," & _
                     "'','', 0,0,0, " & _
                     "'','',0, " & _
                     "'13','" & strNumDocumento & "','" & strDescripObservacion & "'," & _
                     strBancoOrigen & "," & _
                     strBancoDestino & "," & _
                     strTransitoriaOrigen & "," & _
                     strTransitoriaDestino & "," & _
                     "'','','','" & strFormPago & "', '" & strNumCheque & "'," & _
                     "'01','" & strAccion & "') }"
        
            .Execute
            
            
        
        End With
        
        
        Me.MousePointer = vbDefault
        MsgBox Mensaje_Adicion_Exitosa, vbExclamation
        frmMainMdi.stbMdi.Panels(3).Text = "Acción"
    
        cmdOpcion.Visible = True
        With tabPagos
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
    
End Sub

Private Function TodoOk() As Boolean
        
    TodoOk = False
    Dim adoRegistro As ADODB.Recordset
    Dim str_msg As String, str_Tipo As String
    Dim str_pwd As String
    
    If Trim(txtCodParticipe.Text) = "" And chkSinIdentificar.Value = 0 Then
        MsgBox "Seleccione el comitente!.", vbCritical, gstrNombreEmpresa
        Exit Function
    End If
    
    If dtpFechaActual.Value > dtpFechaObligacion.Value Then
        MsgBox "La fecha de obligacion no puede ser menor a la fecha actual!.", vbCritical, gstrNombreEmpresa
        Exit Function
    End If
        
    If CCur(txtMontoPago.Text) = 0 Then
        MsgBox "El Monto de Pago no puede ser cero!.", vbCritical, gstrNombreEmpresa
        Exit Function
    End If
    
    If Trim(txtNroVoucher.Text) = "" Then
        MsgBox "Ingrese el Numero de Voucher o Transaccion!.", vbCritical, gstrNombreEmpresa
        Exit Function
    End If
    
    If cboMoneda.ListIndex = 0 Then
        MsgBox "Seleccione una Moneda", vbCritical, gstrNombreEmpresa
        cboMoneda.SetFocus
        Exit Function
    End If
    
    
    If cboCuentas.ListIndex = 0 Then
        MsgBox "Seleccione una Cuenta", vbCritical, gstrNombreEmpresa
        cboCuentas.SetFocus
        Exit Function
    End If
    
    If chkSinIdentificar.Value = 1 And strTipMov = Codigo_Movimiento_Retiro Then
        MsgBox "No es posible realizar un Retiro sin Identificar", vbCritical, gstrNombreEmpresa
        Exit Function
    End If

    
'    If strTipMov = Codigo_Movimiento_Retiro Then
'        strSQL = "{ call up_IVOrdenRetiroValidacion('" & gstrCodFondoContable & "','" & gstrCodAdministradora & "','" & _
'                     strCodParticipe & "'," & CDec(txtMontoPago.Text) & ",'" & _
'                     strCodMoneda & "') }"
'
'        Set adoRegistro = New ADODB.Recordset
'
'        With adoComm
'            .CommandText = strSQL
'            Set adoRegistro = .Execute
'            If Not adoRegistro.EOF Then
'                str_msg = adoRegistro.Fields("Mensaje")
'                str_Tipo = adoRegistro.Fields("Tipo")
'            End If
'            adoRegistro.Close: Set adoRegistro = Nothing
'        End With
'
'        If str_Tipo = "1" Then
'            MsgBox str_msg, vbCritical, Me.Caption
'            Exit Function
'        End If
'    End If
    '*** Si todo paso OK ***
    TodoOk = True
  
End Function
Private Function TodoOkBackOffice() As Boolean
    
    
    TodoOkBackOffice = False
        
    If tdgConsulta.SelBookmarks.Count - 1 = -1 Then
        MsgBox "Debe seleccionar registros para enviar a Backoffice!", vbCritical, gstrNombreEmpresa
        Exit Function
    End If
    
    If cboEstadoOperacionBusqueda.ListIndex >= 0 Then
        If strEstadoOperacionBusqueda <> Estado_Acuerdo_Ingresado Then
            MsgBox "No se pueden enviar a Backoffice las operaciones seleccionadas!", vbCritical, gstrNombreEmpresa
            Exit Function
        End If
    End If
    
    
        
    '*** Si todo paso OK ***
    TodoOkBackOffice = True
  
End Function
Public Sub SubImprimir(index As Integer)
    
   

    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    Dim strFechaDesde           As String, strFechaHasta        As String

       
    
    If strCodParticipeBusqueda = Valor_Caracter Then
            strCodParticipeBusqueda = Valor_Comodin
    End If


            gstrNameRepo = "AbonoRetiroCtaClienteGrilla"

                Set frmReporte = New frmVisorReporte

                ReDim aReportParamS(6)
                ReDim aReportParamFn(3)
                ReDim aReportParamF(3)

                aReportParamFn(0) = "Fondo"
                aReportParamFn(1) = "NombreEmpresa"
                aReportParamFn(2) = "Hora"
                aReportParamFn(3) = "Usuario"
                
                
                aReportParamF(0) = Trim(cboFondo.Text)
                aReportParamF(1) = gstrNombreEmpresa & Space(1)
                aReportParamF(2) = Format(Time(), "hh:mm:ss")
                aReportParamF(3) = gstrLogin
                
                aReportParamS(0) = strCodFondo
                aReportParamS(1) = gstrCodAdministradora
                aReportParamS(2) = Convertyyyymmdd(dtpFechaOperacionDesde.Value)
                aReportParamS(3) = Convertyyyymmdd(dtpFechaOperacionHasta.Value)
                aReportParamS(4) = strEstadoOperacionBusqueda
                aReportParamS(5) = strCodParticipeBusqueda
                aReportParamS(6) = cboEstadoOperacionBusqueda.Text
                
    gstrSelFrml = Valor_Caracter
    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"

    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())

    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
    frmReporte.Show vbModal

    Set frmReporte = Nothing

    Screen.MousePointer = vbNormal


End Sub

Public Sub Buscar()
        
    Dim strSQL As String
    
    Set adoConsulta = Nothing
    
    Set adoConsulta = New ADODB.Recordset
    
    If Trim(txtCodParticipeBusqueda.Text) <> Valor_Caracter Then
        strSQL = "SELECT " & _
                 "NumOperacion, CodFile, CodAnalitica, DescripOperacion, " & _
                 "TOPE.CodParticipe, DescripParticipe, FechaOperacion, FechaObligacion, " & _
                 "TOPE.CodMoneda, MO.CodSigno as CodSignoMoneda, MontoOperacion, " & _
                 "TipoDocumento, NumDocumento, TOPE.TipoOperacion, TOPE.EstadoOperacion " & _
                 "FROM TesoreriaOperacion TOPE " & _
                 "JOIN Moneda MO ON (MO.CodMoneda = TOPE.CodMoneda) " & _
                 "JOIN ParticipeContrato PC ON (PC.CodParticipe = TOPE.CodParticipe) " & _
                 "WHERE " & _
                 "TOPE.CodFondo = '" & strCodFondo & "' AND " & _
                 "TOPE.CodAdministradora = '" & gstrCodAdministradora & "' AND " & _
                 "TOPE.CodParticipe = '" & strCodParticipeBusqueda & "' AND " & _
                 "TOPE.TipoOperacion in ('32','33','50') AND " & _
                 "TOPE.EstadoOperacion = '" & strEstadoOperacionBusqueda & "' AND " & _
                 "dbo.uf_ACObtenerFechaCorta(TOPE.FechaOperacion)>='" & Convertyyyymmdd(dtpFechaOperacionDesde.Value) & "' AND " & _
                 "dbo.uf_ACObtenerFechaCorta(TOPE.FechaOperacion)<='" & Convertyyyymmdd(dtpFechaOperacionHasta.Value) & "' " & _
                 "ORDER BY NumOperacion"
    Else
        strSQL = "SELECT " & _
                 "NumOperacion, CodFile, CodAnalitica, DescripOperacion, " & _
                 "TOPE.CodParticipe, DescripParticipe, FechaOperacion, FechaObligacion, " & _
                 "TOPE.CodMoneda, MO.CodSigno as CodSignoMoneda, MontoOperacion, " & _
                 "TipoDocumento, NumDocumento, TOPE.TipoOperacion, TOPE.EstadoOperacion " & _
                 "FROM TesoreriaOperacion TOPE " & _
                 "JOIN Moneda MO ON (MO.CodMoneda = TOPE.CodMoneda) " & _
                 "LEFT JOIN ParticipeContrato PC ON (PC.CodParticipe = TOPE.CodParticipe) " & _
                 "WHERE " & _
                 "TOPE.CodFondo = '" & strCodFondo & "' AND " & _
                 "TOPE.CodAdministradora = '" & gstrCodAdministradora & "' AND " & _
                 "TOPE.TipoOperacion in ('32','33','50') AND " & _
                 "TOPE.EstadoOperacion = '" & strEstadoOperacionBusqueda & "' AND " & _
                 "dbo.uf_ACObtenerFechaCorta(TOPE.FechaOperacion)>='" & Convertyyyymmdd(dtpFechaOperacionDesde.Value) & "' AND " & _
                 "dbo.uf_ACObtenerFechaCorta(TOPE.FechaOperacion)<='" & Convertyyyymmdd(dtpFechaOperacionHasta.Value) & "' " & _
                 "ORDER BY NumOperacion"
    End If
        
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

Public Sub Modificar()

            If strEstadoOperacionBusqueda = "05" Then
                MsgBox "No se puede Modificar un registro ya Reversado ", vbOKOnly + vbCritical, Me.Caption
                Exit Sub
            Else
        
                If strEstado = Reg_Consulta Then
                    strEstado = Reg_Edicion
                    LlenarFormulario strEstado
                    cmdOpcion.Visible = False
                    With tabPagos
                        .TabEnabled(0) = False
                        .Tab = 1
                    End With
                    If strEstadoOperacionBusqueda = Codigo_Movimiento_Deposito_No_Identificado Then
                        strEstado = Reg_Adicion
                        lblNumOperacion.Caption = "GENERADO"
                        strNumOperacion = Valor_Caracter
                    End If
                End If
                
            End If
End Sub

Private Sub LlenarFormulario(strModo As String)

    Dim intNumSecuencial    As Integer, intRegistro As Integer
    Dim adoRegistro As New ADODB.Recordset
    Dim strSQL As String
    
    Select Case strModo
        Case Reg_Edicion
        
            If tdgConsulta.Columns("EstadoOperacion") = Codigo_Movimiento_Deposito_No_Identificado Then
                strSQL = "SELECT TOPE.CodFile, TOPE.CodAnalitica, CodDetalleFile, CodSubDetalleFile, TipoMovimiento, TOPE.CodParticipe," & _
                         "ISNULL(PC.DescripParticipe,'') DescripParticipe, DescripOperacion, FechaOperacion, " & _
                         "FechaObligacion, FechaLiquidacion, TOPE.CodMoneda, MontoOperacion, " & _
                         "TipoDocumento, NumDocumento, DescripObservacion, EstadoOperacion, (C.CodFile + C.CodAnalitica + C.CodBanco + C.CodCuentaActivo) cuenta " & _
                         "FROM TesoreriaOperacion TOPE " & _
                         "LEFT JOIN ParticipeContrato PC ON (PC.CodParticipe = TOPE.CodParticipe) " & _
                         "LEFT JOIN BancoCuenta C ON (C.CodFile + C.CodAnalitica  + C.CodCuentaActivo = CodFileBancoTransitoriaOrigen + CodAnaliticaBancoTransitoriaOrigen + CodCuentaBancoTransitoriaOrigen )" & _
                         " WHERE " & _
                         "TOPE.CodFondo = '" & strCodFondo & "' AND " & _
                         "TOPE.CodAdministradora = '" & gstrCodAdministradora & "' AND " & _
                         "TOPE.NumOperacion = '" & Trim(tdgConsulta.Columns("NumOperacion")) & "'"
            Else
            
                If tdgConsulta.Columns("TipoOperacion") = "32" Then
            
                    strSQL = "SELECT TOPE.CodFile, TOPE.CodAnalitica, CodDetalleFile, CodSubDetalleFile, TipoMovimiento, TOPE.CodParticipe," & _
                             "ISNULL(PC.DescripParticipe,'') DescripParticipe, DescripOperacion, FechaOperacion, " & _
                             "FechaObligacion, FechaLiquidacion, TOPE.CodMoneda, MontoOperacion, " & _
                             "TipoDocumento, NumDocumento, DescripObservacion, EstadoOperacion, (C.CodFile + C.CodAnalitica + C.CodBanco + C.CodCuentaActivo) cuenta  " & _
                             "FROM TesoreriaOperacion TOPE " & _
                             "LEFT JOIN ParticipeContrato PC ON (PC.CodParticipe = TOPE.CodParticipe) " & _
                             "LEFT JOIN BancoCuenta C ON (C.CodFile + C.CodAnalitica  + C.CodCuentaActivo = CodFileBancoClienteDestino + CodAnaliticaBancoClienteDestino + CodCuentaBancoClienteDestino )" & _
                             " WHERE " & _
                             "TOPE.CodFondo = '" & strCodFondo & "' AND " & _
                             "TOPE.CodAdministradora = '" & gstrCodAdministradora & "' AND " & _
                             "TOPE.NumOperacion = '" & tdgConsulta.Columns("NumOperacion") & "' AND " & _
                             "C.IndVigente = 'X'"
                Else '"TipoOperacion" = "33"'
                    strSQL = "SELECT TOPE.CodFile, TOPE.CodAnalitica, CodDetalleFile, CodSubDetalleFile, TipoMovimiento, TOPE.CodParticipe," & _
                             "ISNULL(PC.DescripParticipe,'') DescripParticipe, DescripOperacion, FechaOperacion, " & _
                             "FechaObligacion, FechaLiquidacion, TOPE.CodMoneda, MontoOperacion, " & _
                             "TipoDocumento, NumDocumento, DescripObservacion, EstadoOperacion, (C.CodFile + C.CodAnalitica + C.CodBanco + C.CodCuentaActivo) cuenta  " & _
                             "FROM TesoreriaOperacion TOPE " & _
                             "LEFT JOIN ParticipeContrato PC ON (PC.CodParticipe = TOPE.CodParticipe) " & _
                             "LEFT JOIN BancoCuenta C ON (C.CodFile + C.CodAnalitica  + C.CodCuentaActivo = CodFileBancoClienteOrigen + CodAnaliticaBancoClienteOrigen + CodCuentaBancoClienteOrigen )" & _
                             " WHERE " & _
                             "TOPE.CodFondo = '" & strCodFondo & "' AND " & _
                             "TOPE.CodAdministradora = '" & gstrCodAdministradora & "' AND " & _
                             "TOPE.NumOperacion = '" & tdgConsulta.Columns("NumOperacion") & "' AND " & _
                             "C.IndVigente = 'X'"
                End If

            End If
            adoComm.CommandText = strSQL

            Set adoRegistro = adoComm.Execute
            
            If Not adoRegistro.EOF Then
                            
                If tdgConsulta.Columns("EstadoOperacion") = Codigo_Movimiento_Deposito_No_Identificado Then chkSinIdentificar.Value = 0
                
                lblDescripFondo.Caption = Trim(cboFondo.Text)
                
                lblNumOperacion.Caption = tdgConsulta.Columns("NumOperacion")
                
                lblDescripFondo.Caption = cboFondo.Text
                
                txtCodParticipe.Text = Trim(txtCodParticipeBusqueda.Text)
                
                intRegistro = ObtenerItemLista(arrTipMov(), adoRegistro.Fields("TipoMovimiento"))
                If intRegistro >= 0 Then cboTipoMov.ListIndex = intRegistro

                dtpFechaObligacion.Value = adoRegistro.Fields("FechaObligacion")
                
                intRegistro = ObtenerItemLista(arrMoneda(), adoRegistro.Fields("CodMoneda"))
                If intRegistro >= 0 Then cboMoneda.ListIndex = intRegistro
                
                txtMontoPago.Text = adoRegistro.Fields("MontoOperacion")
                txtNroVoucher.Text = adoRegistro.Fields("NumDocumento")
                
                txtDescripObservaciones.Text = adoRegistro.Fields("DescripObservacion")
                
                txtCodParticipe.Text = adoRegistro.Fields("CodParticipe")
                                
                lblDescripParticipe.Caption = adoRegistro.Fields("DescripParticipe")
            
                intRegistro = ObtenerItemLista(arrCuentas(), adoRegistro.Fields("cuenta"))
                If intRegistro >= 0 Then cboCuentas.ListIndex = intRegistro
                
            End If
            
            
            If strEstadoOperacionBusqueda = "02" Or strEstadoOperacionBusqueda = "03" Then
                Call DeshabilitarDetalle
            End If
            
            
            adoRegistro.Close
            
            Set adoRegistro = Nothing
        
        Case Reg_Adicion
            ''Llenar los combos del formulario
            
            Call HabilitarDetalle
                        
            lblDescripFondo.Caption = Trim(cboFondo.Text)
                        
            lblNumOperacion.Caption = "GENERADO"
            
            strNumOperacion = Valor_Caracter
            
            txtCodParticipe.Text = Valor_Caracter
            
            lblDescripParticipe.Caption = Valor_Caracter
            
            dtpFechaObligacion.Value = dtpFechaActual.Value
            
            cboTipoMov.ListIndex = 0
            
            cboMoneda.ListIndex = 0
            
            txtMontoPago.Text = "0.00"
            
            txtNroVoucher.Text = Valor_Caracter
            
            txtDescripObservaciones.Text = Valor_Caracter
            
    End Select
    
End Sub

Private Sub DeshabilitarDetalle()
            
            txtCodParticipe.Enabled = False
            
            dtpFechaObligacion.Enabled = False
            
            cboTipoMov.Enabled = False
            
            cboMoneda.Enabled = False
            
            txtMontoPago.Enabled = False
            
            txtNroVoucher.Enabled = False
            
            txtDescripObservaciones.Enabled = False
            
            cmdBusquedaParticipe.Enabled = False
            
            cboCuentas.Enabled = False
            
            cboCuentasControl.Enabled = False
            
            chkSinIdentificar.Enabled = False
            
            
            

End Sub


Private Sub HabilitarDetalle()
            
            txtCodParticipe.Enabled = True
            
            dtpFechaObligacion.Enabled = True
            
            cboTipoMov.Enabled = True
            
            cboMoneda.Enabled = True
            
            txtMontoPago.Enabled = True
            
            txtNroVoucher.Enabled = True
            
            txtDescripObservaciones.Enabled = True
            
            cmdBusquedaParticipe.Enabled = True
            
            cboCuentas.Enabled = True
            
            cboCuentasControl.Enabled = True
            
            chkSinIdentificar.Enabled = True
            
            
            

End Sub

Public Sub Adicionar()
    
    strEstado = Reg_Adicion
    LlenarFormulario strEstado
    cmdOpcion.Visible = False
    With tabPagos
        .TabEnabled(0) = False
        .Tab = 1
    End With
                
End Sub




Private Sub cmdReservar_Click()
    Call Reversar
End Sub

Private Sub Form_Load()

    Call InicializarValores
    Call CargarListas
'    Call Buscar
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
Private Sub CargarListas()
    
    '*** Fondos ***
    strSQL = "{ call up_ACSelDatosParametro(74,'" & gstrCodAdministradora & "','" & gstrCodFondoContable & "') }"
    CargarControlLista strSQL, cboFondo, garrFondo(), Valor_Caracter
    
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
    
    strSQL = "select CodParametro CODIGO,DescripParametro DESCRIP from AuxiliarParametro Where CodTipoParametro = 'MOVCTA' AND CodParametro IN ('01','02') Order By DescripParametro"
    CargarControlLista strSQL, cboTipoMov, arrTipMov(), Valor_Caracter
    If cboTipoMov.ListCount > 0 Then cboTipoMov.ListIndex = 0
    
    
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP from AuxiliarParametro Where CodTipoParametro = 'ESTACU' Order By CodParametro"
    CargarControlLista strSQL, cboEstadoOperacionBusqueda, arrEstadoOperacionBusqueda(), Valor_Caracter
    If cboEstadoOperacionBusqueda.ListCount > 0 Then cboEstadoOperacionBusqueda.ListIndex = 0
   
    '*** Moneda ***
    strSQL = "{ call up_ACSelDatos(2) }"
    CargarControlLista strSQL, cboMoneda, arrMoneda(), Sel_Defecto
    
    If cboMoneda.ListCount > 0 Then cboMoneda.ListIndex = 0
        
End Sub
Private Sub InicializarValores()

    '*** Valores Iniciales ***
    strEstado = Reg_Defecto
    
    tabPagos.Tab = 0
    tabPagos.TabEnabled(1) = False
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdAccion2.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    Set cmdIdentificar.FormularioActivo = Me
    
    dtpFechaOperacionDesde.Value = gdatFechaActual
    dtpFechaOperacionHasta.Value = gdatFechaActual
 
End Sub

Private Sub tabPagos_Click(PreviousTab As Integer)

    Select Case tabPagos.Tab
        Case 1
            If gstrFormulario = "frmConfirmacionSolicitud" Then tabPagos.Tab = 0
            If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vQuery)
            If strEstado = Reg_Defecto Then tabPagos.Tab = 0
        
    End Select
    
End Sub

Private Sub tdgConsulta_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
        
    If ColIndex = 3 Then
        Call DarFormatoValor(Value, Decimales_Monto)
    End If
        
End Sub


Private Sub tdgConsulta_HeadClick(ByVal ColIndex As Integer)

    Static numColindex As Integer

    tdgConsulta.Splits(0).Columns(numColindex).HeadingStyle.ForegroundPicture = Null

    Call OrdenarDBGrid(ColIndex, adoConsulta, tdgConsulta)
    
    numColindex = ColIndex

End Sub

    Private Sub txtCodParticipe_LostFocus()

    Dim rst As Boolean
    Dim adoRegistro As ADODB.Recordset
    
    If Trim(txtCodParticipe.Text) <> Valor_Caracter Then
        rst = False
        
       txtCodParticipe.Text = Format(txtCodParticipe.Text, "00000000000000000000")
    
        With adoComm
            Set adoRegistro = New ADODB.Recordset
                    
            .CommandText = ""
            strSQL = "SELECT CodParticipe,AP1.DescripParametro TipoIdentidad,NumIdentidad,DescripParticipe,FechaIngreso,TipoIdentidad CodTipoIdentidad,AP2.DescripParametro TipoMancomuno "
            strSQL = strSQL & "FROM ParticipeContrato JOIN AuxiliarParametro AP1 ON(AP1.CodParametro=ParticipeContrato.TipoIdentidad AND AP1.CodTipoParametro='TIPIDE') "
            strSQL = strSQL & "JOIN AuxiliarParametro AP2 ON(AP2.CodParametro=ParticipeContrato.TipoMancomuno AND AP2.CodTipoParametro='TIPMAN') "
            strSQL = strSQL & "WHERE CodParticipe='" & Trim(txtCodParticipe.Text) & "'"
            .CommandText = strSQL
            Set adoRegistro = .Execute
    
            Do Until adoRegistro.EOF
                rst = True
                lblDescripParticipe.Caption = Trim(adoRegistro("DescripParticipe"))
                strCodParticipe = Trim(adoRegistro("CodParticipe"))
                adoRegistro.MoveNext
                
                'Obtener el saldo del Cliente
                'Call ObtenerSaldoCliente
                
            Loop
            adoRegistro.Close: Set adoRegistro = Nothing
            
            If Not rst Then
                strCodParticipe = ""
                txtCodParticipe.Text = ""
                lblDescripParticipe.Caption = ""
                MsgBox "Codigo de Cliente Incorrecto", vbCritical, Me.Caption
            End If
            
        End With
    Else
        strCodParticipe = ""
        txtCodParticipe.Text = ""
        lblDescripParticipe.Caption = ""
    End If


End Sub

Private Sub ObtenerSaldoCliente()
    
    Dim adoRegistro As ADODB.Recordset
    If strCodParticipe = Valor_Caracter Then Exit Sub
    If strCodMoneda = Valor_Caracter Then Exit Sub
    
    strSQL = "{ call up_IVOrdenRetiroValidacion('" & gstrCodFondoContable & "','" & gstrCodAdministradora & "','" & _
                     strCodParticipe & "',0,'" & _
                     strCodMoneda & "') }"
                     
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        .CommandText = strSQL
        Set adoRegistro = .Execute
        If Not adoRegistro.EOF Then
            lblSaldoCliente.Caption = adoRegistro.Fields("Saldo")
            
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
End Sub

Private Sub txtCodParticipeBusqueda_LostFocus()
    
    Dim rst As Boolean
    Dim adoRegistro As ADODB.Recordset
    
    If Trim(txtCodParticipeBusqueda.Text) <> Valor_Caracter Then
        rst = False
        
     txtCodParticipeBusqueda.Text = Format(txtCodParticipeBusqueda.Text, "00000000000000000000")
    
        With adoComm
            Set adoRegistro = New ADODB.Recordset
                    
            .CommandText = ""
            strSQL = "SELECT CodParticipe,AP1.DescripParametro TipoIdentidad,NumIdentidad,DescripParticipe,FechaIngreso,TipoIdentidad CodTipoIdentidad,AP2.DescripParametro TipoMancomuno "
            strSQL = strSQL & "FROM ParticipeContrato JOIN AuxiliarParametro AP1 ON(AP1.CodParametro=ParticipeContrato.TipoIdentidad AND AP1.CodTipoParametro='TIPIDE') "
            strSQL = strSQL & "JOIN AuxiliarParametro AP2 ON(AP2.CodParametro=ParticipeContrato.TipoMancomuno AND AP2.CodTipoParametro='TIPMAN') "
            strSQL = strSQL & "WHERE CodParticipe ='" & Trim(txtCodParticipeBusqueda.Text) & "'"
            .CommandText = strSQL
            Set adoRegistro = .Execute
    
            Do Until adoRegistro.EOF
                rst = True
                lblDescripParticipeBusqueda.Caption = Trim(adoRegistro("DescripParticipe"))
                strCodParticipeBusqueda = Trim(adoRegistro("CodParticipe"))
                adoRegistro.MoveNext
            Loop
            adoRegistro.Close: Set adoRegistro = Nothing
            
            If Not rst Then
                strCodParticipeBusqueda = ""
                txtCodParticipeBusqueda.Text = ""
                lblDescripParticipeBusqueda.Caption = ""
                MsgBox "Codigo de Cliente Incorrecto", vbCritical, Me.Caption
            End If
            
        End With
    Else
        strCodParticipeBusqueda = ""
        txtCodParticipeBusqueda.Text = ""
        lblDescripParticipeBusqueda.Caption = ""
    End If
    
    Call Buscar
    
End Sub

Private Sub txtMontoPago_GotFocus()
    Call FormatoGotFocus(txtMontoPago)

End Sub

Private Sub txtMontoPago_LostFocus()
    Call FormatoCajaTexto(txtMontoPago, Decimales_Monto)
End Sub

Private Sub FormatoGotFocus(txtCrl As TextBox)
    Call FormatoCajaTexto(txtCrl, Decimales_Monto)
    With txtMontoPago
        .SelStart = 0
        .SelLength = Len(.Text)
         Call FormatoCajaTexto(txtCrl, Decimales_Monto)
    End With

End Sub

Private Sub ConfiguraRecordsetAuxiliar()

    Set adoRegistroAux = New ADODB.Recordset

    With adoRegistroAux
       .CursorLocation = adUseClient
       .Fields.Append "CodFondo", adVarChar, 3
       .Fields.Append "CodAdministradora", adVarChar, 3
       .Fields.Append "NumOperacion", adVarChar, 10
       .LockType = adLockBatchOptimistic
    End With

    adoRegistroAux.Open

End Sub

Private Function TodoOkReversar() As Boolean
    
    
    TodoOkReversar = False
        
    If tdgConsulta.SelBookmarks.Count - 1 = -1 Or tdgConsulta.SelBookmarks.Count - 1 > 0 Then
        MsgBox "Debe seleccionar un registro para reversar", vbCritical, gstrNombreEmpresa
        Exit Function
    End If
        
    If cboEstadoOperacionBusqueda.ListIndex >= 0 Then
        If strEstadoOperacionBusqueda <> "02" Then
            MsgBox "Solo se puede reversar operaciones ya procesadas", vbOKOnly + vbCritical, Me.Caption
            Exit Function
        End If
    Else
        MsgBox "Debe seleccionar algun estado de operacion", vbCritical, gstrNombreEmpresa
        If cboEstadoOperacionBusqueda.Enabled Then cboEstadoOperacionBusqueda.SetFocus
        Exit Function
    End If
        
    '*** Si todo paso OK ***
    TodoOkReversar = True
  
End Function


Private Sub Reversar()


    Dim strFechaGrabar  As String
    Dim strNumOperacion  As String 'se usa
    Dim motivo          As String
    Dim str_msg, str_pwd         As String
    Dim adoRegistro     As ADODB.Recordset

    If strEstado <> Reg_Edicion Then
        If strEstado <> Reg_Consulta Then Exit Sub
    End If
    
    If Not TodoOkReversar() Then Exit Sub

    If MsgBox("Desea reversar el Movimiento Cambiario Nro. " & tdgConsulta.Columns("NumOperacion").Value & " ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
        
        Me.MousePointer = vbHourglass
        
'        If gdatFechaActual > tdgConsulta.Columns(1).Value Then 'cambiar la condicion por la fecha
'            str_msg = str_msg + " Para continuar se requiere la Autorización, continuar?"
'            If MsgBox(str_msg, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
'                Exit Sub
'            Else
'                'Call inputbox_Password(frmOrdenRentaVariableCliente, "*")
'                str_pwd = InputBox(" Ingrese la contraseña de Autorización", App.Title)
'
'                strSQL = "SELECT  * FROM AuxiliarParametro WHERE CodTipoParametro = 'PWDORD' " _
'                 & "AND CodParametro= '01' and ValorParametro = '" & str_pwd & "'"
'
'                Set adoRegistro = New ADODB.Recordset
'                With adoComm
'                    .CommandText = strSQL
'                    Set adoRegistro = .Execute
'                    If adoRegistro.EOF Then
'                        MsgBox "La contraseña ingresada es incorrecta...", vbCritical, Me.Caption
'                        Exit Sub
'                    End If
'                    adoRegistro.Close: Set adoRegistro = Nothing
'                End With
'
'            End If
'        End If
        motivo = InputBox("Ingrese el motivo", App.Title)

    Else

        Exit Sub

    End If


    On Error GoTo Ctrl_Error
                                        
    With adoComm
        
        .CommandType = adCmdText
        
        strFechaGrabar = Convertyyyymmdd(gdatFechaActual) & Space(1) & Format(Time, "hh:ss")
    
        strNumOperacion = tdgConsulta.Columns("NumOperacion").Value
        
        '*** Cabecera ***
        .CommandText = "{ call up_TEProcAnularOperacionTesoreria('" & _
            strCodFondo & "','" & gstrCodAdministradora & "','" & _
            Valor_NumOpeCajaBancos & "','" & strNumOperacion & "','" & Space(1) + motivo + Space(1) + "') }"
            
        adoConn.Execute .CommandText
      
        Me.MousePointer = vbDefault
       
        Call Buscar
        
        MsgBox Mensaje_Proceso_Exitoso, vbExclamation
            
        frmMainMdi.stbMdi.Panels(3).Text = "Acción"

                                                                       
    End With
    
    Exit Sub
    
Ctrl_Error:
    
    MsgBox Mensaje_Proceso_NoExitoso, vbCritical
    Me.MousePointer = vbDefault
    
    

End Sub








