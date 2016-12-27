VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{830D5A06-9B70-4F8C-98B6-7A19C4E7760B}#1.0#0"; "TAMControls.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmMovimientoCambiario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Operaciones de Cambio"
   ClientHeight    =   10485
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   13590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10485
   ScaleWidth      =   13590
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
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
      Left            =   10080
      Picture         =   "frmMovimientoCambiario.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   122
      Top             =   9600
      Width           =   1200
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
      Left            =   6720
      Picture         =   "frmMovimientoCambiario.frx":05EC
      Style           =   1  'Graphical
      TabIndex        =   121
      Top             =   9600
      Width           =   1200
   End
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   11520
      TabIndex        =   120
      Top             =   9600
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
      Caption0        =   "&Salir"
      Tag0            =   "9"
      Visible0        =   0   'False
      ToolTipText0    =   "Salir"
      UserControlWidth=   1200
   End
   Begin TAMControls2.ucBotonEdicion2 cmdOpcion 
      Height          =   735
      Left            =   720
      TabIndex        =   119
      Top             =   9600
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
   Begin TabDlg.SSTab tabTipoCambio 
      Height          =   9345
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   16484
      _Version        =   393216
      Style           =   1
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
      TabPicture(0)   =   "frmMovimientoCambiario.frx":0C11
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tdgConsulta"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraCriterio"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Datos de Operación"
      TabPicture(1)   =   "frmMovimientoCambiario.frx":0C2D
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraCuenta(0)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Datos de Liquidación"
      TabPicture(2)   =   "frmMovimientoCambiario.frx":0C49
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraCuenta(1)"
      Tab(2).ControlCount=   1
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
         Height          =   3120
         Left            =   450
         TabIndex        =   79
         Top             =   690
         Width           =   11895
         Begin VB.ComboBox cboFondo 
            Height          =   315
            Left            =   2130
            Style           =   2  'Dropdown List
            TabIndex        =   84
            Top             =   540
            Width           =   7605
         End
         Begin VB.CommandButton cmdBusqueda 
            Caption         =   "..."
            Height          =   315
            Index           =   1
            Left            =   4440
            TabIndex        =   83
            ToolTipText     =   "Búsqueda de Partícipe"
            Top             =   1020
            Width           =   390
         End
         Begin VB.TextBox txtCodParticipeBusqueda 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2130
            MaxLength       =   20
            TabIndex        =   82
            Top             =   1020
            Width           =   2280
         End
         Begin VB.ComboBox cboEstadoOperacionBusqueda 
            Height          =   315
            Left            =   2130
            Style           =   2  'Dropdown List
            TabIndex        =   81
            Top             =   1950
            Width           =   2745
         End
         Begin VB.CommandButton cmdProcesar 
            Caption         =   "Procesar"
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
            Left            =   9120
            Picture         =   "frmMovimientoCambiario.frx":0C65
            Style           =   1  'Graphical
            TabIndex        =   80
            Top             =   2160
            Width           =   1200
         End
         Begin MSComCtl2.DTPicker dtpFechaDesde 
            Height          =   315
            Left            =   2130
            TabIndex        =   89
            Top             =   2430
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
            Format          =   207159297
            CurrentDate     =   38068
         End
         Begin MSComCtl2.DTPicker dtpFechaHasta 
            Height          =   315
            Left            =   5820
            TabIndex        =   90
            Top             =   2430
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
            Format          =   207159297
            CurrentDate     =   38068
         End
         Begin VB.Label lblDescrip 
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
            Index           =   19
            Left            =   360
            TabIndex        =   92
            Top             =   2460
            Width           =   1065
         End
         Begin VB.Label lblDescrip 
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
            Index           =   1
            Left            =   4320
            TabIndex        =   91
            Top             =   2490
            Width           =   705
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
            Index           =   34
            Left            =   360
            TabIndex        =   88
            Top             =   570
            Width           =   540
         End
         Begin VB.Label lblDescripParticipeBusqueda 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2160
            TabIndex        =   87
            Top             =   1530
            Width           =   7635
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Participe"
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
            Left            =   360
            TabIndex        =   86
            Top             =   1050
            Width           =   765
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
            Index           =   0
            Left            =   360
            TabIndex        =   85
            Top             =   1980
            Width           =   1515
         End
      End
      Begin VB.Frame fraCuenta 
         Caption         =   "Cuentas Bancarias"
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
         Height          =   6705
         Index           =   1
         Left            =   -74640
         TabIndex        =   58
         Top             =   630
         Width           =   12750
         Begin VB.Frame Frame1 
            Caption         =   "Cuentas Bancarias"
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
            Height          =   3285
            Left            =   450
            TabIndex        =   59
            Top             =   810
            Width           =   11565
            Begin VB.ComboBox cboCuentaBancoResultado 
               Height          =   315
               Left            =   3960
               Style           =   2  'Dropdown List
               TabIndex        =   98
               Top             =   2640
               Width           =   4635
            End
            Begin VB.ComboBox cboCuentaBancoTransitoriaOrigen 
               Height          =   315
               Left            =   6390
               Style           =   2  'Dropdown List
               TabIndex        =   65
               Top             =   870
               Width           =   4695
            End
            Begin VB.ComboBox cboCuentaBancoTransitoriaDestino 
               Height          =   315
               Left            =   6390
               Style           =   2  'Dropdown List
               TabIndex        =   64
               Top             =   1380
               Width           =   4695
            End
            Begin VB.ComboBox cboCuentaBancoClienteDestino 
               Height          =   315
               Left            =   1170
               Style           =   2  'Dropdown List
               TabIndex        =   62
               Top             =   1380
               Width           =   4635
            End
            Begin VB.ComboBox cboCuentaBancoClienteOrigen 
               Height          =   315
               Left            =   1170
               Style           =   2  'Dropdown List
               TabIndex        =   60
               Top             =   870
               Width           =   4635
            End
            Begin VB.Line LinUtilidad 
               BorderColor     =   &H00008000&
               BorderStyle     =   4  'Dash-Dot
               X1              =   6120
               X2              =   6120
               Y1              =   1710
               Y2              =   2430
            End
            Begin VB.Label lblUtilidad 
               Caption         =   "v"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   225
               Left            =   6060
               TabIndex        =   100
               Top             =   2310
               Width           =   135
            End
            Begin VB.Label lblDescrip 
               AutoSize        =   -1  'True
               Caption         =   "Utilidad"
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
               Left            =   3120
               TabIndex        =   99
               Top             =   2700
               Width           =   660
            End
            Begin VB.Label lblFechaComitente 
               BackStyle       =   0  'Transparent
               Caption         =   "<"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   255
               Left            =   5910
               TabIndex        =   97
               Top             =   1410
               Width           =   135
            End
            Begin VB.Label lblFlechaTransitoria 
               BackStyle       =   0  'Transparent
               Caption         =   ">"
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
               Height          =   255
               Left            =   6210
               TabIndex        =   96
               Top             =   900
               Width           =   135
            End
            Begin VB.Line linDirecto 
               BorderColor     =   &H00000080&
               Visible         =   0   'False
               X1              =   6210
               X2              =   6210
               Y1              =   1020
               Y2              =   1530
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               Caption         =   "TRANSITORIAS"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   7950
               TabIndex        =   69
               Top             =   360
               Width           =   1785
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "PARTICIPE"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   2880
               TabIndex        =   68
               Top             =   330
               Width           =   1275
            End
            Begin VB.Line Line2 
               BorderColor     =   &H00000080&
               X1              =   6390
               X2              =   11040
               Y1              =   660
               Y2              =   660
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00000080&
               X1              =   1200
               X2              =   5790
               Y1              =   660
               Y2              =   660
            End
            Begin VB.Label lblTransitoriaComitente 
               Caption         =   "---"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   255
               Left            =   6000
               TabIndex        =   67
               Top             =   1410
               Width           =   255
            End
            Begin VB.Label lblComitenteTransitoria 
               BackStyle       =   0  'Transparent
               Caption         =   "---"
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
               Height          =   255
               Left            =   6000
               TabIndex        =   66
               Top             =   900
               Width           =   435
            End
            Begin VB.Label lblDescrip 
               AutoSize        =   -1  'True
               Caption         =   "Destino"
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
               Index           =   16
               Left            =   360
               TabIndex        =   63
               Top             =   1410
               Width           =   660
            End
            Begin VB.Label lblDescrip 
               AutoSize        =   -1  'True
               Caption         =   "Origen"
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
               Left            =   360
               TabIndex        =   61
               Top             =   900
               Width           =   570
            End
         End
      End
      Begin VB.Frame fraCuenta 
         Caption         =   "Operación"
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
         Height          =   8415
         Index           =   0
         Left            =   -74640
         TabIndex        =   38
         Top             =   630
         Width           =   12840
         Begin TAMControls2.ucBotonEdicion2 cmdGrabar 
            Height          =   735
            Left            =   8910
            TabIndex        =   118
            Top             =   7440
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
         Begin VB.Frame fraTipoLiquidacion 
            Caption         =   "Tipo de Liquidacion"
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
            Height          =   1035
            Left            =   6720
            TabIndex        =   113
            Top             =   2550
            Width           =   5865
            Begin VB.OptionButton optTipoLiquidacion 
               Caption         =   "Sin Recursos"
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
               Height          =   255
               Index           =   1
               Left            =   2610
               TabIndex        =   115
               Top             =   480
               Width           =   1545
            End
            Begin VB.OptionButton optTipoLiquidacion 
               Caption         =   "Con Recursos"
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
               Height          =   255
               Index           =   0
               Left            =   360
               TabIndex        =   114
               Top             =   480
               Value           =   -1  'True
               Width           =   1755
            End
         End
         Begin VB.Frame fraTipoOperacion 
            Caption         =   "Tipo de Operación"
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
            Height          =   1035
            Left            =   360
            TabIndex        =   110
            Top             =   2550
            Width           =   5865
            Begin VB.OptionButton optTipo 
               Caption         =   "Compra"
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
               Height          =   255
               Index           =   0
               Left            =   360
               TabIndex        =   112
               Top             =   480
               Value           =   -1  'True
               Width           =   1095
            End
            Begin VB.OptionButton optTipo 
               Caption         =   "Venta"
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
               Height          =   255
               Index           =   1
               Left            =   2610
               TabIndex        =   111
               Top             =   480
               Width           =   1095
            End
         End
         Begin VB.ComboBox cboEstadoOperacion 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2400
            Style           =   2  'Dropdown List
            TabIndex        =   77
            Top             =   7800
            Visible         =   0   'False
            Width           =   2985
         End
         Begin VB.Frame fraLiquidacion 
            BorderStyle     =   0  'None
            Height          =   1365
            Left            =   6360
            TabIndex        =   70
            Top             =   4560
            Width           =   5745
            Begin VB.TextBox txtNumVoucher 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2520
               MaxLength       =   40
               TabIndex        =   73
               Top             =   0
               Width           =   2385
            End
            Begin TAMControls.TAMTextBox txtMontoRecibidoBanco 
               Height          =   315
               Left            =   2520
               TabIndex        =   101
               Top             =   1050
               Width           =   2355
               _ExtentX        =   4154
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
               Container       =   "frmMovimientoCambiario.frx":11CD
               Text            =   "0.00"
               Decimales       =   2
               Estilo          =   4
               CambiarConFoco  =   -1  'True
               ColorEnfoque    =   12648447
               EnterTab        =   -1  'True
               Apariencia      =   1
               Borde           =   1
               DecimalesValue  =   4
            End
            Begin TAMControls.TAMTextBox txtMontoEntregadoBanco 
               Height          =   315
               Left            =   2520
               TabIndex        =   102
               Top             =   540
               Width           =   2355
               _ExtentX        =   4154
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
               Container       =   "frmMovimientoCambiario.frx":11E9
               Text            =   "0.00"
               Decimales       =   2
               Estilo          =   4
               CambiarConFoco  =   -1  'True
               ColorEnfoque    =   12648447
               EnterTab        =   -1  'True
               Apariencia      =   1
               Borde           =   1
               DecimalesValue  =   4
            End
            Begin VB.Label lblSignoMonedaRecibidoBanco 
               AutoSize        =   -1  'True
               Caption         =   "PEN"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   4920
               TabIndex        =   76
               Top             =   1110
               Width           =   330
            End
            Begin VB.Label lblDescrip 
               AutoSize        =   -1  'True
               Caption         =   "Monto Recibido"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   33
               Left            =   300
               TabIndex        =   75
               Top             =   1110
               Width           =   1125
            End
            Begin VB.Label lblDescrip 
               AutoSize        =   -1  'True
               Caption         =   "Nro. de Voucher"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   30
               Left            =   300
               TabIndex        =   74
               Top             =   30
               Width           =   1170
            End
            Begin VB.Label lblDescrip 
               AutoSize        =   -1  'True
               Caption         =   "Monto Entregado"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   32
               Left            =   300
               TabIndex        =   72
               Top             =   600
               Width           =   1230
            End
            Begin VB.Label lblSignoMonedaEntregadoBanco 
               AutoSize        =   -1  'True
               Caption         =   "PEN"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   4920
               TabIndex        =   71
               Top             =   600
               Width           =   330
            End
         End
         Begin VB.TextBox txtObservacion 
            Height          =   675
            Left            =   2130
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   49
            Top             =   6720
            Width           =   9645
         End
         Begin VB.TextBox txtCodParticipe 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2190
            MaxLength       =   20
            TabIndex        =   46
            Top             =   510
            Width           =   2415
         End
         Begin VB.CommandButton cmdBusqueda 
            Caption         =   "..."
            Height          =   315
            Index           =   0
            Left            =   4710
            TabIndex        =   45
            ToolTipText     =   "Búsqueda de Partícipe"
            Top             =   510
            Width           =   375
         End
         Begin VB.ComboBox cboMoneda 
            Height          =   315
            Left            =   8850
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   1860
            Width           =   2985
         End
         Begin MSComCtl2.DTPicker dtpFechaCambio 
            Height          =   315
            Left            =   2190
            TabIndex        =   40
            Top             =   1530
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
            Format          =   207159297
            CurrentDate     =   38067
         End
         Begin TAMControls.ucBotonEdicion cmdGrabar2 
            Height          =   390
            Left            =   8910
            TabIndex        =   51
            Top             =   7470
            Visible         =   0   'False
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   688
            Buttons         =   2
            Caption0        =   "&Guardar"
            Tag0            =   "2"
            Visible0        =   0   'False
            ToolTipText0    =   "Nuevo"
            Caption1        =   "&Cancelar"
            Tag1            =   "8"
            Visible1        =   0   'False
            ToolTipText1    =   "Buscar"
            UserControlHeight=   390
            UserControlWidth=   2700
         End
         Begin TAMControls.TAMTextBox txtMontoRecibido 
            Height          =   315
            Left            =   2610
            TabIndex        =   93
            Top             =   4560
            Width           =   2355
            _ExtentX        =   4154
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
            Container       =   "frmMovimientoCambiario.frx":1205
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   12648447
            EnterTab        =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   4
         End
         Begin TAMControls.TAMTextBox txtMontoEntregado 
            Height          =   315
            Left            =   2610
            TabIndex        =   94
            Top             =   5070
            Width           =   2355
            _ExtentX        =   4154
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
            Container       =   "frmMovimientoCambiario.frx":1221
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   12648447
            EnterTab        =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   4
         End
         Begin MSComCtl2.DTPicker dtpFechaValor 
            Height          =   315
            Left            =   2190
            TabIndex        =   103
            Top             =   2040
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
            Format          =   207159297
            CurrentDate     =   38067
         End
         Begin TAMControls.TAMTextBox txtMontoUtilidad 
            Height          =   315
            Left            =   8880
            TabIndex        =   107
            Top             =   6120
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   556
            BackColor       =   32768
            ForeColor       =   16777215
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
            ForeColor       =   16777215
            Locked          =   -1  'True
            Container       =   "frmMovimientoCambiario.frx":123D
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
         Begin TAMControls.TAMTextBox txtValorTipoCambioCosto 
            Height          =   315
            Left            =   8880
            TabIndex        =   116
            Top             =   4050
            Width           =   2355
            _ExtentX        =   4154
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
            Container       =   "frmMovimientoCambiario.frx":1259
            Text            =   "0.000000"
            Decimales       =   6
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   12648447
            EnterTab        =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   6
         End
         Begin TAMControls.TAMTextBox txtValorTipoCambio 
            Height          =   315
            Left            =   2610
            TabIndex        =   117
            Top             =   4020
            Width           =   2355
            _ExtentX        =   4154
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
            Container       =   "frmMovimientoCambiario.frx":1275
            Text            =   "0.000000"
            Decimales       =   6
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   12648447
            EnterTab        =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   6
         End
         Begin VB.Label lblSignoMonedaUtilidad 
            AutoSize        =   -1  'True
            Caption         =   "PEN"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   11280
            TabIndex        =   109
            Top             =   6180
            Width           =   330
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Utilidad Operación"
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
            Index           =   36
            Left            =   6630
            TabIndex        =   108
            Top             =   6150
            Width           =   1485
         End
         Begin VB.Label lblSignoMonedaParBanco 
            AutoSize        =   -1  'True
            Caption         =   "PEN / USD"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   11280
            TabIndex        =   106
            Top             =   4080
            Width           =   840
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Cambio Banco"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   29
            Left            =   6630
            TabIndex        =   105
            Top             =   4050
            Width           =   1395
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00800000&
            X1              =   6330
            X2              =   6330
            Y1              =   3690
            Y2              =   6510
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Valor"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   35
            Left            =   360
            TabIndex        =   104
            Top             =   2070
            Width           =   855
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Estado Operación"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   21
            Left            =   480
            TabIndex        =   78
            Top             =   7800
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.Label lblSignoMonedaPar 
            AutoSize        =   -1  'True
            Caption         =   "PEN / USD"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   5010
            TabIndex        =   57
            Top             =   4050
            Width           =   840
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
            Left            =   10020
            TabIndex        =   56
            Top             =   510
            Width           =   1755
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Nro. Operación"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   20
            Left            =   8580
            TabIndex        =   55
            Top             =   540
            Width           =   1080
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Moneda de Cambio"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   18
            Left            =   6570
            TabIndex        =   54
            Top             =   1860
            Width           =   1380
         End
         Begin VB.Label lblSignoMonedaEntregado 
            AutoSize        =   -1  'True
            Caption         =   "PEN"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   5010
            TabIndex        =   53
            Top             =   5130
            Width           =   330
         End
         Begin VB.Label lblSignoMonedaRecibido 
            AutoSize        =   -1  'True
            Caption         =   "USD"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   5010
            TabIndex        =   52
            Top             =   4620
            Width           =   345
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones"
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
            Index           =   45
            Left            =   390
            TabIndex        =   50
            Top             =   6690
            Width           =   1245
         End
         Begin VB.Label lblNombreCliente 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2190
            TabIndex        =   48
            Top             =   990
            Width           =   9615
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Participe"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   31
            Left            =   360
            TabIndex        =   47
            Top             =   540
            Width           =   1215
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto Entregado"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   27
            Left            =   360
            TabIndex        =   44
            Top             =   5100
            Width           =   1230
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto Recibido"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   26
            Left            =   360
            TabIndex        =   43
            Top             =   4590
            Width           =   1125
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Cambio Comitente"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   25
            Left            =   360
            TabIndex        =   42
            Top             =   4050
            Width           =   1635
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Operación"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   24
            Left            =   360
            TabIndex        =   41
            Top             =   1560
            Width           =   1230
         End
      End
      Begin VB.Frame fraConfirmacion 
         Caption         =   "Orden de..."
         Height          =   6075
         Left            =   -74760
         TabIndex        =   2
         Top             =   420
         Width           =   10875
         Begin VB.TextBox txtDescripMotivo 
            Height          =   285
            Left            =   2520
            TabIndex        =   11
            Text            =   " "
            Top             =   4515
            Width           =   7365
         End
         Begin VB.CheckBox ChkMonDiferente 
            Caption         =   "Transacción en distinta moneda"
            Enabled         =   0   'False
            Height          =   255
            Left            =   270
            TabIndex        =   10
            Top             =   5490
            Visible         =   0   'False
            Width           =   3255
         End
         Begin VB.ComboBox cboCuentas 
            Height          =   315
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   2700
            Width           =   7920
         End
         Begin VB.ComboBox cboMonedaPago 
            Height          =   315
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1950
            Width           =   2565
         End
         Begin VB.TextBox txtTipoCambio 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   7500
            Locked          =   -1  'True
            TabIndex        =   7
            Text            =   " 1"
            Top             =   1920
            Width           =   1845
         End
         Begin VB.ComboBox cboFormaPago 
            Height          =   315
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   4920
            Width           =   2295
         End
         Begin VB.TextBox txtNroDocumento 
            Height          =   285
            Left            =   7500
            TabIndex        =   5
            Top             =   4950
            Visible         =   0   'False
            Width           =   2385
         End
         Begin VB.ComboBox cboTipoDocumento 
            Height          =   315
            Left            =   2520
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   420
            Width           =   2565
         End
         Begin VB.TextBox txtNumDocumento 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   7500
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   3
            Text            =   " "
            Top             =   420
            Width           =   2775
         End
         Begin MSComCtl2.DTPicker dtpFechaContable 
            Height          =   315
            Left            =   2520
            TabIndex        =   12
            Top             =   1005
            Width           =   1935
            _ExtentX        =   3413
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
            Format          =   207159297
            CurrentDate     =   38068
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Movimiento Cuenta"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   10
            Left            =   5280
            TabIndex        =   35
            Top             =   3240
            Width           =   1875
         End
         Begin VB.Label lblTipoCambio 
            Caption         =   "Tipo de Cambio Arbitraje"
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   5250
            TabIndex        =   34
            Top             =   1980
            Width           =   2115
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Moneda de Pago"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   8
            Left            =   360
            TabIndex        =   33
            Top             =   1980
            Width           =   1455
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Disponible"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   7
            Left            =   5280
            TabIndex        =   32
            Top             =   3660
            Width           =   1440
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Motivo"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   5
            Left            =   360
            TabIndex        =   31
            Top             =   4545
            Width           =   1395
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Depositar a la Cuenta"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   4
            Left            =   390
            TabIndex        =   30
            Top             =   2760
            Width           =   1860
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Contabilización"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   360
            TabIndex        =   29
            Top             =   1050
            Width           =   1890
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   5280
            TabIndex        =   28
            Top             =   1050
            Width           =   540
         End
         Begin VB.Label lblMonto 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7500
            TabIndex        =   27
            Top             =   990
            Width           =   2415
         End
         Begin VB.Label lblSaldoCuenta 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7530
            TabIndex        =   26
            Top             =   3660
            Width           =   2475
         End
         Begin VB.Label lblMovCuenta 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7530
            TabIndex        =   25
            Top             =   3240
            Width           =   2475
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   1
            X1              =   300
            X2              =   10620
            Y1              =   4200
            Y2              =   4200
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Forma de Pago"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   11
            Left            =   360
            TabIndex        =   24
            Top             =   4935
            Width           =   1455
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Nro Cheque"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   12
            Left            =   5310
            TabIndex        =   23
            Top             =   4980
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   0
            X1              =   270
            X2              =   10590
            Y1              =   2460
            Y2              =   2460
         End
         Begin VB.Label lblMonedaTC 
            Caption         =   "(PEN/USD)"
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
            Height          =   225
            Left            =   9390
            TabIndex        =   22
            Top             =   1980
            Width           =   975
         End
         Begin VB.Label lblMonedaCuentaLiquidacion 
            Caption         =   "PEN"
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
            Index           =   0
            Left            =   10080
            TabIndex        =   21
            Top             =   3270
            Width           =   465
         End
         Begin VB.Label lblMonedaCuentaLiquidacion 
            Caption         =   "PEN"
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
            Height          =   225
            Index           =   1
            Left            =   10080
            TabIndex        =   20
            Top             =   3690
            Width           =   465
         End
         Begin VB.Label lblMonedaOrden 
            Caption         =   "PEN"
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
            Height          =   225
            Left            =   9990
            TabIndex        =   19
            Top             =   1050
            Width           =   465
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "XXXXX"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   6
            Left            =   3570
            TabIndex        =   18
            Top             =   5520
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label lblMonedaContable 
            Caption         =   "PEN"
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
            Height          =   225
            Left            =   9990
            TabIndex        =   17
            Top             =   1470
            Width           =   465
         End
         Begin VB.Label lblMontoContable 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7500
            TabIndex        =   16
            Top             =   1410
            Width           =   2415
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto Contable"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   13
            Left            =   5280
            TabIndex        =   15
            Top             =   1440
            Width           =   1350
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   2
            X1              =   300
            X2              =   10620
            Y1              =   1800
            Y2              =   1800
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Documento"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   14
            Left            =   360
            TabIndex        =   14
            Top             =   480
            Width           =   1410
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Nro. de Documento"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   15
            Left            =   5280
            TabIndex        =   13
            Top             =   450
            Width           =   1665
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   3
            X1              =   300
            X2              =   10620
            Y1              =   840
            Y2              =   840
         End
      End
      Begin TAMControls.ucBotonEdicion cmdAccion 
         Height          =   390
         Left            =   -66930
         TabIndex        =   1
         Top             =   6570
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   688
         Buttons         =   2
         Caption0        =   "&Procesar"
         Tag0            =   "2"
         Visible0        =   0   'False
         ToolTipText0    =   "Procesar"
         Caption1        =   "&Cancelar"
         Tag1            =   "8"
         Visible1        =   0   'False
         ToolTipText1    =   "Cancelar"
         UserControlHeight=   390
         UserControlWidth=   2700
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Height          =   4755
         Left            =   450
         OleObjectBlob   =   "frmMovimientoCambiario.frx":1291
         TabIndex        =   95
         Top             =   4140
         Width           =   11895
      End
   End
   Begin TAMControls.ucBotonEdicion cmdSalir2 
      Height          =   390
      Left            =   11520
      TabIndex        =   36
      Top             =   9840
      Visible         =   0   'False
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
   Begin TAMControls.ucBotonEdicion cmdOpcion2 
      Height          =   390
      Left            =   720
      TabIndex        =   37
      Top             =   9720
      Visible         =   0   'False
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   688
      Buttons         =   2
      Caption0        =   "&Nuevo"
      Tag0            =   "0"
      Visible0        =   0   'False
      ToolTipText0    =   "Nuevo"
      Caption1        =   "&Buscar"
      Tag1            =   "5"
      Visible1        =   0   'False
      ToolTipText1    =   "Buscar"
      UserControlHeight=   390
      UserControlWidth=   2700
   End
End
Attribute VB_Name = "frmMovimientoCambiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrFondo()              As String, arrFondoOrden()              As String
Dim arrTipoInstrumento()    As String, arrTipoInstrumentoOrden()    As String
Dim arrEstado()             As String, arrTipoCambio()               As String
Dim arrOperacion()          As String, arrNegociacion()             As String
Dim arrEmisor()             As String, arrMoneda()                  As String
Dim arrBaseAnual()          As String, arrTipoTasa()                As String
Dim arrOrigen()             As String, arrClaseInstrumento()        As String
Dim arrTitulo()             As String, arrSubClaseInstrumento()     As String
Dim arrAgente()             As String, arrConceptoCosto()           As String
Dim arrModalidadOrden()     As String, arrEstadoOperacionBusqueda() As String
Dim arrEstadoOperacion()    As String

Dim arrCuentaBancoClienteOrigen() As String
Dim arrCuentaBancoClienteDestino() As String
Dim arrCuentaBancoTransitoriaOrigen() As String
Dim arrCuentaBancoTransitoriaDestino() As String
Dim arrCuentaBancoResultado() As String

Dim strCuentaBancoClienteOrigen As String
Dim strCuentaBancoClienteDestino As String
Dim strCuentaBancoTransitoriaOrigen As String
Dim strCuentaBancoTransitoriaDestino As String
Dim strCuentaBancoResultado As String

Dim strCodCuentaBancoClienteOrigen As String
Dim strCodFileBancoClienteOrigen As String
Dim strCodAnaliticaBancoClienteOrigen As String

Dim strCodCuentaBancoTransitoriaDestino As String
Dim strCodFileBancoTransitoriaDestino As String
Dim strCodAnaliticaBancoTransitoriaDestino As String

Dim strCodCuentaBancoTransitoriaOrigen As String
Dim strCodFileBancoTransitoriaOrigen As String
Dim strCodAnaliticaBancoTransitoriaOrigen As String

Dim strCodCuentaBancoClienteDestino As String
Dim strCodFileBancoClienteDestino As String
Dim strCodAnaliticaBancoClienteDestino As String

Dim strCodCuentaBancoResultado As String
Dim strCodFileBancoResultado As String
Dim strCodAnaliticaBancoResultado As String

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
Dim strCodEstadoBusqueda    As String

Dim strCodFile              As String, strCodAnalitica              As String
Dim strCodGrupo             As String, strCodCiiu                   As String
Dim strEstadoOrden          As String, strCodCategoria              As String
Dim strCodRiesgo            As String, strCodSubRiesgo              As String
Dim strCalcVcto             As String, strCodSector                 As String
Dim strCodTipoCostoBolsa    As String, strCodTipoCostoConasev       As String
Dim strCodTipoCostoFondo    As String, strCodTipoCavali             As String
Dim strIndCuponCero         As String, strCodGarantia               As String
Dim strCodTipoCostoFondoG   As String, strCodMonedaPago             As String
Dim dblTipoCambio           As Double, dblComisionFondoG            As Double
Dim dblComisionBolsa        As Double, dblComisionConasev           As Double
Dim dblComisionFondo        As Double, dblComisionCavali            As Double
Dim strTipoDocumento        As String, strNumDocumento              As String
Dim intDiasVigencia         As Integer, strTituloCambio             As String
Dim strCodParticipe         As String, strCodModalidad              As String
Dim strTipoMov              As String, strSignoMoneda               As String
Dim strCodSignoMoneda       As String, strCodMonedaParPorDefecto    As String
Dim strCodTipoCambio        As String, strValorCambio               As String
Dim strCodMonedaRecibido    As String, strCodMonedaEntregado        As String
Dim strEstadoOperacion      As String
Dim adoRegistroAux          As ADODB.Recordset
Dim adoConsulta             As ADODB.Recordset

'************************************
Dim strCodParticipeBusqueda As String

Private Sub cboCuentaBancoClienteDestino_Click()

    strCuentaBancoClienteDestino = Valor_Caracter
    
    If cboCuentaBancoClienteDestino.ListIndex < 0 Then Exit Sub
    
    strCuentaBancoClienteDestino = Trim(arrCuentaBancoClienteDestino(cboCuentaBancoClienteDestino.ListIndex))

    strCodCuentaBancoClienteDestino = Mid(strCuentaBancoClienteDestino, 1, 10)
    strCodFileBancoClienteDestino = Mid(strCuentaBancoClienteDestino, 11, 3)
    strCodAnaliticaBancoClienteDestino = Mid(strCuentaBancoClienteDestino, 14, 8)


End Sub

Private Sub cboCuentaBancoClienteOrigen_Click()

    strCuentaBancoClienteOrigen = Valor_Caracter
    
    If cboCuentaBancoClienteOrigen.ListIndex < 0 Then Exit Sub
    
    strCuentaBancoClienteOrigen = Trim(arrCuentaBancoClienteOrigen(cboCuentaBancoClienteOrigen.ListIndex))

    strCodCuentaBancoClienteOrigen = Mid(strCuentaBancoClienteOrigen, 1, 10)
    strCodFileBancoClienteOrigen = Mid(strCuentaBancoClienteOrigen, 11, 3)
    strCodAnaliticaBancoClienteOrigen = Mid(strCuentaBancoClienteOrigen, 14, 8)

End Sub

Private Sub cboCuentaBancoTransitoriaDestino_Click()

    strCuentaBancoTransitoriaDestino = Valor_Caracter
    
    If cboCuentaBancoTransitoriaDestino.ListIndex < 0 Then Exit Sub
    
    strCuentaBancoTransitoriaDestino = Trim(arrCuentaBancoTransitoriaDestino(cboCuentaBancoTransitoriaDestino.ListIndex))

    strCodCuentaBancoTransitoriaDestino = Mid(strCuentaBancoTransitoriaDestino, 1, 10)
    strCodFileBancoTransitoriaDestino = Mid(strCuentaBancoTransitoriaDestino, 11, 3)
    strCodAnaliticaBancoTransitoriaDestino = Mid(strCuentaBancoTransitoriaDestino, 14, 8)


End Sub

Private Sub cboCuentaBancoTransitoriaOrigen_Click()

    strCuentaBancoTransitoriaOrigen = Valor_Caracter
    
    If cboCuentaBancoTransitoriaOrigen.ListIndex < 0 Then Exit Sub
    
    strCuentaBancoTransitoriaOrigen = Trim(arrCuentaBancoTransitoriaOrigen(cboCuentaBancoTransitoriaOrigen.ListIndex))

    strCodCuentaBancoTransitoriaOrigen = Mid(strCuentaBancoTransitoriaOrigen, 1, 10)
    strCodFileBancoTransitoriaOrigen = Mid(strCuentaBancoTransitoriaOrigen, 11, 3)
    strCodAnaliticaBancoTransitoriaOrigen = Mid(strCuentaBancoTransitoriaOrigen, 14, 8)

End Sub

Private Sub cboCuentaBancoResultado_Click()

    strCuentaBancoResultado = Valor_Caracter
    If cboCuentaBancoResultado.ListIndex < 0 Then Exit Sub
    
    strCuentaBancoResultado = Trim(arrCuentaBancoResultado(cboCuentaBancoResultado.ListIndex))

    strCodCuentaBancoResultado = Mid(strCuentaBancoResultado, 1, 10)
    strCodFileBancoResultado = Mid(strCuentaBancoResultado, 11, 3)
    strCodAnaliticaBancoResultado = Mid(strCuentaBancoResultado, 14, 8)

End Sub

Private Sub cboEstadoOperacionBusqueda_Click()

    strCodEstadoBusqueda = Valor_Caracter
    If cboEstadoOperacionBusqueda.ListIndex < 0 Then Exit Sub
    
    strCodEstadoBusqueda = Trim(arrEstadoOperacionBusqueda(cboEstadoOperacionBusqueda.ListIndex))

    Call Buscar
    
End Sub

Private Sub cboFondo_Click()

    strCodFondo = Valor_Caracter
    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(arrFondo(cboFondo.ListIndex))


End Sub

Private Sub cboMoneda_Click()

    Dim strCodMonedaParEvaluacion   As String
    Dim dblTipoCambioOrden          As Double
    
    strCodMoneda = Valor_Caracter: strSignoMoneda = Valor_Caracter
    If cboMoneda.ListIndex < 0 Then Exit Sub
    
    strCodMoneda = Trim(arrMoneda(cboMoneda.ListIndex))
    
    If strTipoMov = "V" Then
        strCodMonedaRecibido = strCodMoneda
            
        strCodMonedaEntregado = Codigo_Moneda_Local 'por defecto es asi
    End If

    If strTipoMov = "C" Then
        strCodMonedaRecibido = Codigo_Moneda_Local  'por defecto es asi
        strCodMonedaEntregado = strCodMoneda
    End If
    
    lblSignoMonedaRecibido.Caption = ObtenerCodSignoMoneda(strCodMonedaRecibido)
      
    lblSignoMonedaEntregado.Caption = ObtenerCodSignoMoneda(strCodMonedaEntregado)
    
    lblSignoMonedaRecibidoBanco.Caption = lblSignoMonedaEntregado.Caption
    lblSignoMonedaEntregadoBanco.Caption = lblSignoMonedaRecibido.Caption
   
   
        If strTipoMov = "C" Then
            lblSignoMonedaUtilidad.Caption = lblSignoMonedaRecibido.Caption
        Else
            lblSignoMonedaUtilidad.Caption = lblSignoMonedaEntregado.Caption
            
        End If
   
    'lblSignoMonedaUtilidad.Caption = lblSignoMonedaRecibido.Caption

   
    strCodMonedaParEvaluacion = strCodMonedaRecibido & strCodMonedaEntregado
    
    strCodMonedaParPorDefecto = ObtenerMonedaParPorDefecto(strCodTipoCambio, strCodMonedaParEvaluacion)
        
    If strCodMonedaParPorDefecto = "0000" Then strCodMonedaParPorDefecto = strCodMonedaParEvaluacion
    
    If strCodMonedaParPorDefecto <> Valor_Caracter Then
        lblSignoMonedaPar.Caption = Trim(ObtenerCodSignoMoneda(Mid(strCodMonedaParPorDefecto, 1, 2))) + "/" + Trim(ObtenerCodSignoMoneda(Mid(strCodMonedaParPorDefecto, 3, 2)))
    Else
        lblSignoMonedaPar.Caption = Valor_Caracter
    End If
    
    dblTipoCambioOrden = CStr(ObtenerTipoCambioMoneda2(strCodTipoCambio, strValorCambio, dtpFechaCambio.Value, Mid(strCodMonedaParPorDefecto, 1, 2), Mid(strCodMonedaParPorDefecto, 3, 2)))
    If dblTipoCambioOrden = 0 Then dblTipoCambioOrden = 1
    
    txtValorTipoCambio.Text = 0#
    txtValorTipoCambio.Text = dblTipoCambioOrden 'CStr(ObtenerTipoCambioMoneda2(strCodTipoCambio, strValorCambio, datFechaConsulta, Mid(strCodMonedaParPorDefecto, 1, 2), Mid(strCodMonedaParPorDefecto, 3, 2)))
    
    txtValorTipoCambioCosto.Text = txtValorTipoCambio.Value

    '*** Cuentas Ordinarias Origen ***
    strSQL = "{ call up_ACSelDatosParametro(71,'" & strCodFondo & "','" & gstrCodAdministradora & "','" & strCodMonedaRecibido & "') }"
    CargarControlLista strSQL, cboCuentaBancoClienteOrigen, arrCuentaBancoClienteOrigen(), ""
   
    If cboCuentaBancoClienteOrigen.ListCount > 0 Then cboCuentaBancoClienteOrigen.ListIndex = 0
    
    '*** Cuentas Ordinarias Destino ***
    strSQL = "{ call up_ACSelDatosParametro(71,'" & strCodFondo & "','" & gstrCodAdministradora & "','" & strCodMonedaEntregado & "') }"
    CargarControlLista strSQL, cboCuentaBancoClienteDestino, arrCuentaBancoClienteDestino(), ""
    
    If cboCuentaBancoClienteDestino.ListCount > 0 Then cboCuentaBancoClienteDestino.ListIndex = 0
   
    '*** Cuentas Administrativas Origen ***
    strSQL = "{ call up_ACSelDatosParametro(72,'" & strCodFondo & "','" & gstrCodAdministradora & "','" & strCodMonedaRecibido & "') }"
    CargarControlLista strSQL, cboCuentaBancoTransitoriaOrigen, arrCuentaBancoTransitoriaOrigen(), ""
    
    If cboCuentaBancoTransitoriaOrigen.ListCount > 0 Then cboCuentaBancoTransitoriaOrigen.ListIndex = 0

    '*** Cuentas Administrativas Destino ***
    strSQL = "{ call up_ACSelDatosParametro(72,'" & strCodFondo & "','" & gstrCodAdministradora & "','" & strCodMonedaEntregado & "') }"
    CargarControlLista strSQL, cboCuentaBancoTransitoriaDestino, arrCuentaBancoTransitoriaDestino(), ""
    
    If cboCuentaBancoTransitoriaDestino.ListCount > 0 Then cboCuentaBancoTransitoriaDestino.ListIndex = 0

    '*** Cuentas de Resultado ***
    strSQL = "{ call up_ACSelDatosParametro(72,'" & strCodFondo & "','" & gstrCodAdministradora & "','" & strCodMonedaRecibido & "') }"
    CargarControlLista strSQL, cboCuentaBancoResultado, arrCuentaBancoResultado(), ""
    
    If cboCuentaBancoResultado.ListCount > 0 Then cboCuentaBancoResultado.ListIndex = 0



End Sub

'Private Sub cboTipoCambio_Click()
'    strTipoMov = Valor_Caracter
'    If cboTipoCambio.ListIndex < 0 Then Exit Sub
'     strTipoMov = Trim(arrTipoCambio(cboTipoCambio.ListIndex))
'End Sub

Private Sub cmdBusqueda_Click(index As Integer)
    
    Select Case index
    
    Case 0:
        
        gstrFormulario = "frmMovimientoCambiario"
    
        frmBusquedaParticipe.Show vbModal
        
        If Trim(gstrCodParticipe) <> "" Then strCodParticipe = gstrCodParticipe
        
        Me.txtCodParticipe.SetFocus
        
    Case 1:
    
        gstrFormulario = "frmMovimientoCambiarioBusqueda"
    
        frmBusquedaParticipe.Show vbModal
        
        If Trim(gstrCodParticipe) <> "" Then strCodParticipeBusqueda = gstrCodParticipe
        
        Me.txtCodParticipe.SetFocus
    
    End Select
    
    


End Sub

Private Sub cmdImprimir_Click()
    Call SubImprimir2(1)
End Sub



Private Sub cmdProcesar_Click()

    Dim intContador                 As Integer
    Dim intRegistro                 As Integer
    Dim strTesoreriaOperacionXML    As String
    Dim objTesoreriaOperacionXML    As DOMDocument60
    Dim strFechaGrabar              As String
    Dim strMsgError                 As String
    'Dim adoRegistro                 As ADODB.Recordset
        
    On Error GoTo ErrorHandler
        
    If TodoOkProceso() Then
        '*** Realizar proceso de contabilización ***
        If MsgBox("Datos correctos. ¿Procedemos a procesar esta(s) operacion(es)?", vbQuestion + vbYesNo, "Observación") = vbNo Then Exit Sub
        
        Me.MousePointer = vbHourglass
        
        intContador = tdgConsulta.SelBookmarks.Count - 1
               
        dtpFechaCambio.Value = gdatFechaActual
               
        strFechaGrabar = Convertyyyymmdd(dtpFechaCambio.Value) & Space(1) & Format(Time, "hh:mm")
                   
        Call ConfiguraRecordsetAuxiliar
        
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
          
            Next
           
            Call XMLADORecordset(objTesoreriaOperacionXML, "TesoreriaOperacion", "Operacion", adoRegistroAux, strMsgError)
            strTesoreriaOperacionXML = objTesoreriaOperacionXML.xml
                
            .CommandText = "{ call up_TEProcOperacionCambio('" & _
                strCodFondo & "','" & gstrCodAdministradora & "','" & _
                strFechaGrabar & "','" & _
                strTesoreriaOperacionXML & "') }"
            
            .Execute .CommandText
                                             
        End With
        
        Me.MousePointer = vbDefault
        
        MsgBox Mensaje_Proceso_Exitoso, vbExclamation
        
        frmMainMdi.stbMdi.Panels(3).Text = "Acción"
        
        'cmdOpcion.Visible = True
        With tabTipoCambio
            .TabEnabled(0) = True
            .Tab = 0
        End With
                
        Call Buscar
        tdgConsulta.ReBind
        Me.Refresh
    End If

ErrorHandler:
    
    If err.Number <> 0 Then
        MsgBox err.Number & " " & err.Description, vbCritical + vbOKOnly, Me.Caption
        Me.MousePointer = vbDefault
    End If

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

Private Sub Form_Load()

    Call InicializarValores
    Call CargarListas
    Call Buscar
    Call DarFormato
    
    Call ValidarPermisoUsoControl(Trim(gstrLoginUS), Me, Trim(App.Title) + Separador_Codigo_Objeto + _
    gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)
    
    CentrarForm Me
    
    Call ValidaExisteTipoCambio(Codigo_TipoCambio_SBS, gstrFechaActual)
    
End Sub



Private Sub CargarListas()
    
    '*** Fondos ***
    strSQL = "{ call up_ACSelDatosParametro(74,'" & gstrCodAdministradora & "','" & gstrCodFondoContable & "') }"
    CargarControlLista strSQL, cboFondo, arrFondo(), ""
    
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
     
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP from AuxiliarParametro Where CodTipoParametro = 'ESTES2' Order By CodParametro"
    CargarControlLista strSQL, cboEstadoOperacionBusqueda, arrEstadoOperacionBusqueda(), Valor_Caracter
    If cboEstadoOperacionBusqueda.ListCount > 0 Then cboEstadoOperacionBusqueda.ListIndex = 0
     
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP from AuxiliarParametro Where CodTipoParametro = 'ESTTES' Order By CodParametro"
    CargarControlLista strSQL, cboEstadoOperacion, arrEstadoOperacion(), Valor_Caracter
    If cboEstadoOperacion.ListCount > 0 Then cboEstadoOperacion.ListIndex = 0
          
    '*** Moneda ***
    strSQL = "{ call up_ACSelDatos(39) }"
    CargarControlLista strSQL, cboMoneda, arrMoneda(), "" 'Sel_Defecto
    
    If cboMoneda.ListCount > 0 Then cboMoneda.ListIndex = 0
     
       
End Sub

'Public Sub CentrarForm(ByVal Forma As Form)
'
'    Forma.Left = (frmMainMdi.Width - Forma.Width) / 2
'    Forma.Top = (frmMainMdi.Height - frmMainMdi.tlbMdi.Height - frmMainMdi.stbMdi.Height - 780 - Forma.Height) / 2
'
'End Sub

Private Sub InicializarValores()

    '*** Valores Iniciales ***
    tabTipoCambio.Tab = 0
    tabTipoCambio.TabEnabled(1) = False
    tabTipoCambio.TabEnabled(2) = False
    '*** Valores Por Defecto
    
    strCodTipoCambio = gstrCodClaseTipoCambioOperacionFondo     'Codigo_TipoCambio_Conasev
    'strValorCambio = gstrValorTipoCambioOperacion            'Codigo_Valor_TipoCambioCompra

    dtpFechaDesde.Value = gdatFechaActual
    dtpFechaHasta.Value = dtpFechaDesde.Value
        
    Call optTipoLiquidacion_Click(0)
        
    'Call optTipo_Click(0)
        
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgConsulta.Columns(0).Width = tdgConsulta.Width * 0.01 * 5
    tdgConsulta.Columns(1).Width = tdgConsulta.Width * 0.01 * 10
    tdgConsulta.Columns(3).Width = tdgConsulta.Width * 0.01 * 16
    tdgConsulta.Columns(5).Width = tdgConsulta.Width * 0.01 * 9
    
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdGrabar.FormularioActivo = Me
    'Set cmdProcesar.FormularioActivo = Me
    
    
End Sub

Public Sub Accion(nAccion As ButtonAction)
    
    Select Case nAccion
        
        Case vNew
            Call Adicionar
        Case vModify
            Call Modificar
       Case vSearch
            Call Buscar
        Case vSave
            Call Grabar
        Case vCancel
            Call Cancelar
        Case vDelete
            Call Eliminar
        Case vExit
            Call Salir
        
    End Select
    
End Sub

Public Sub Modificar()

 frmMainMdi.stbMdi.Panels(3).Text = "Movimiento Cambiario de la SAB..."
    
    If strCodEstadoBusqueda = "01" Then
            strEstado = Reg_Edicion
            LlenarFormulario strEstado
            cmdOpcion.Visible = False
            cmdReservar.Visible = False
            With tabTipoCambio
                .TabEnabled(0) = False
                .TabEnabled(1) = True
                .TabEnabled(2) = True
                .Tab = 1
            End With
    Else
        
        MsgBox "No puede Modificar un registro " & cboEstadoOperacionBusqueda.Text & "", vbCritical, Me.Caption
        
    End If


End Sub

Public Sub Salir()

    Unload Me
    
End Sub

Public Sub Adicionar()

   frmMainMdi.stbMdi.Panels(3).Text = "Movimiento Cambiario de la SAB..."

    strEstado = Reg_Adicion
    LlenarFormulario strEstado
    cmdOpcion.Visible = False
    cmdReservar.Visible = False
    With tabTipoCambio
        .TabEnabled(0) = False
        .TabEnabled(1) = True
        .TabEnabled(2) = True
        .Tab = 1
    End With
    
End Sub

Public Sub Buscar()
        
    Dim strFechaMovimBCDesde          As String
    Dim strFechaMovimBCHasta          As String
    Dim datFechaSiguiente             As Date
    Dim strSQL As String
                      
    Me.MousePointer = vbHourglass

    If Not IsNull(dtpFechaDesde.Value) Or Not IsNull(dtpFechaHasta.Value) Then
        strFechaMovimBCDesde = Convertyyyymmdd(dtpFechaDesde.Value)
        datFechaSiguiente = DateAdd("d", 1, dtpFechaHasta.Value)
        strFechaMovimBCHasta = Convertyyyymmdd(datFechaSiguiente)
    End If
        
    strSQL = "{ call up_TEListarOperacionCambio ('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
            strFechaMovimBCDesde & "','" & strFechaMovimBCHasta & "','" & strCodEstadoBusqueda & "','" & strCodParticipeBusqueda & "' ) }"
    
    Set adoConsulta = New ADODB.Recordset

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

    Me.MousePointer = vbDefault
    
End Sub


Public Sub Grabar()
                        
    Dim intAccion As Integer, lngNumError   As Integer
    Dim strNumSolicitud As String
    Dim strDescripOperacion As String, strNumDocumento As String
    Dim strAccion As String, strDescripObservacion As String
    Dim strFechaObligacion As String, strFechaGrabar As String
    Dim strTipoOperacion As String, strMensaje As String
    Dim curMontoRecibido As Double, curMontoEntregado As Double, dblValorTipoCambio As Double
    Dim strTipoContraparte As String, strCodContraparte As String, strTipoMovimiento As String
    Dim strNumOperacion As String, strCodTipoMovimiento As String, strModalidadOperacion As String
    
    If strEstado = Reg_Consulta Then Exit Sub
            
    On Error GoTo ErrorHandler
        
    If TodoOk() Then
    
        'Confirmar la Operacion
        strMensaje = "Se procede a registrar la siguiente Operacion de Cambio:" & Chr(vbKeyReturn) & _
        "Tipo de Operación          : " & Space(4) & IIf(optTipo(0), "Compra", "Venta") & Chr(vbKeyReturn) & _
        "Monto Recibido             : " & txtMontoRecibido.Text & " " & lblSignoMonedaRecibido.Caption & Chr(vbKeyReturn) & _
        "Monto Entregado            : " & txtMontoEntregado.Text & " " & lblSignoMonedaEntregado.Caption & Chr(vbKeyReturn) & _
        "Tipo de Cambio Cliente     : " & Space(4) & txtValorTipoCambio.Text & Chr(vbKeyReturn) & _
        "Tipo de Cambio Costo       : " & Space(4) & txtValorTipoCambioCosto.Text

        If MsgBox(strMensaje, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then Exit Sub
        
        curMontoRecibido = CCur(txtMontoRecibido.Text)
        curMontoEntregado = CCur(txtMontoEntregado.Text)
        dblValorTipoCambio = CDbl(txtValorTipoCambio.Text)
        
        strTipoContraparte = ""
        strCodContraparte = ""
        strCodTipoMovimiento = ""
                
        strTipoOperacion = Codigo_Caja_Operacion_Cambio
        strDescripOperacion = "Operación de Cambio Comitente - " + Trim(lblNombreCliente.Caption)
                 
        strNumDocumento = Trim(txtNroDocumento.Text)
        
        strDescripObservacion = Trim(txtObservacion.Text)
    
        strTipoMovimiento = Codigo_Movimiento_Retiro ' Retiro
    
        If optTipoLiquidacion(0).Value Then strModalidadOperacion = "01" 'CON RECURSOS
        
        If optTipoLiquidacion(1).Value Then strModalidadOperacion = "02" 'SIN RECURSOS
        
    
        If strEstado = Reg_Adicion Then
            strAccion = "I"
            strNumOperacion = Valor_Caracter
        End If
        
        If strEstado = Reg_Edicion Then
            strAccion = "U"
            strNumOperacion = lblNumOperacion.Caption
        End If
        
        strFechaGrabar = Convertyyyymmdd(dtpFechaCambio.Value) & Space(1) & Format(Time, "hh:mm")
        strFechaObligacion = Convertyyyymmdd(dtpFechaCambio.Value)
        
        strNumOperacion = lblNumOperacion.Caption
               
        Me.MousePointer = vbHourglass
        
        ''------------------------------
        ''---Registro en BD-------------
                                
        With adoComm
            
            .CommandText = "{ call up_TEManTesoreriaOperacion ('" & _
                                 strCodFondo & "','" & gstrCodAdministradora & "','" & strNumOperacion & "'," & _
                                 "'000', '00000000', '000'," & _
                                 "'000','" & strTipoOperacion & "','" & strModalidadOperacion & "','" & strCodTipoMovimiento & "','" & strCodParticipe & "','" & strTipoContraparte & "','" & strCodContraparte & "'," & _
                                 "'" & strDescripOperacion & "','" & strFechaGrabar & "','" & strFechaObligacion & "'," & _
                                 "'19000101','" & strCodMonedaRecibido & "'," & CCur(txtMontoRecibido.Text) & "," & _
                                 "'" & strCodMonedaEntregado & "'," & CCur(txtMontoEntregado.Text) & "," & _
                                 "'" & gstrCodClaseTipoCambioOperacionFondo & "','" & gstrValorTipoCambioOperacion & "'," & CDbl(txtValorTipoCambio.Text) & "," & _
                                 CDbl(txtMontoEntregadoBanco.Text) & "," & CDbl(txtMontoRecibidoBanco.Text) & "," & _
                                 "'" & gstrCodClaseTipoCambioOperacionFondo & "','" & gstrValorTipoCambioOperacion & "'," & CDbl(txtValorTipoCambioCosto.Text) & "," & _
                                 "'13','" & strNumDocumento & "','" & strDescripObservacion & "'," & _
                                 "'" & strCodCuentaBancoClienteOrigen & "','" & strCodFileBancoClienteOrigen & "','" & strCodAnaliticaBancoClienteOrigen & "'," & _
                                 "'" & strCodCuentaBancoClienteDestino & "','" & strCodFileBancoClienteDestino & "','" & strCodAnaliticaBancoClienteDestino & "'," & _
                                 "'" & strCodCuentaBancoTransitoriaOrigen & "','" & strCodFileBancoTransitoriaOrigen & "','" & strCodAnaliticaBancoTransitoriaOrigen & "'," & _
                                 "'" & strCodCuentaBancoTransitoriaDestino & "','" & strCodFileBancoTransitoriaDestino & "','" & strCodAnaliticaBancoTransitoriaDestino & "'," & _
                                 "'" & strCodCuentaBancoResultado & "','" & strCodFileBancoResultado & "','" & strCodAnaliticaBancoResultado & "','', '', " & _
                                 "'01','" & strAccion & "') }"
            
            .Execute
            
        
        End With
        
        
        Me.MousePointer = vbDefault
    
        MsgBox Mensaje_Adicion_Exitosa, vbExclamation
    
        frmMainMdi.stbMdi.Panels(3).Text = "Acción"
    
        cmdOpcion.Visible = True
        With tabTipoCambio
            .TabEnabled(0) = True
            .TabEnabled(1) = False
            .TabEnabled(2) = False
            .Tab = 0
        End With
        
        Call Limpiar
        Call Buscar
        
        cmdReservar.Visible = True
        
    End If

    Exit Sub
    
ErrorHandler:
    
    If err.Number <> 0 Then
        MsgBox err.Number & " " & err.Description, vbCritical + vbOKOnly, Me.Caption
        Me.MousePointer = vbDefault
    End If

'CtrlError:
'    Me.MousePointer = vbDefault
'    intAccion = ControlErrores
'    Select Case intAccion
'        Case 0: Resume
'        Case 1: Resume Next
'        Case 2: Exit Sub
'        Case Else
'            lngNumError = err.Number
'            err.Raise Number:=lngNumError
'            err.Clear
'    End Select
    
End Sub

Private Function TodoOkAnular()

    TodoOkAnular = False

    If tdgConsulta.SelBookmarks.Count = 0 Or tdgConsulta.SelBookmarks.Count > 1 Then
        MsgBox "Debe seleccionar un registro para Anular", vbCritical + vbOKOnly, gstrNombreEmpresa
        Exit Function
    End If
        
    If cboEstadoOperacionBusqueda.ListIndex > -1 Then
        If strCodEstadoBusqueda = "02" Then
            MsgBox "No se puede anular un registro ya procesado", vbOKOnly + vbCritical, Me.Caption
            Exit Function
        ElseIf strCodEstadoBusqueda = "03" Then
            MsgBox "Este registro ya esta anulado", vbOKOnly + vbCritical, Me.Caption
            Exit Function
        End If
    Else
        MsgBox "Debe seleccionar algun estado de operacion", vbCritical + vbOKOnly, gstrNombreEmpresa
        If cboEstadoOperacionBusqueda.Enabled Then cboEstadoOperacionBusqueda.SetFocus
        Exit Function
    End If
        
    '*** Si todo paso OK ***
    TodoOkAnular = True

End Function

Private Sub Eliminar()

    If strEstado = Reg_Consulta Then
        
        strEstadoOperacion = "03"
        
        If Not TodoOkAnular() Then Exit Sub
        
        If MsgBox("Se procederá a anular el Movimiento Cambiario Nro. " & tdgConsulta.Columns("NumOperacion").Value & vbNewLine & vbNewLine & "Seguro de continuar ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
        
            adoComm.CommandText = "UPDATE TesoreriaOperacion SET EstadoOperacion='" & strEstadoOperacion & "' " & _
            "WHERE NumOperacion='" & tdgConsulta.Columns("NumOperacion").Value & "' AND " & _
            "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' "
            
            adoConn.Execute adoComm.CommandText
            
            Call Buscar
            
        End If
        
    End If

End Sub

Public Function TodoOk() As Boolean

    TodoOk = False
    
'    If Trim(txtNDocumento.Text) = "" Then
'        MsgBox "Debe Ingresar un Numero de Documento!", vbCritical, gstrNombreEmpresa
'        txtNumDocumento.SetFocus
'        Exit Function
'    End If
    
    If Trim(strCodParticipe) = "" Then
        MsgBox "Debe Ingresar el codigo de Participe!", vbCritical, gstrNombreEmpresa
        txtNumDocumento.SetFocus
        Exit Function
    End If
    
    
    If Trim(txtValorTipoCambio.Text) = "" Then
        MsgBox "Debe ingresar el Tipo de Cambio!", vbCritical, gstrNombreEmpresa
        txtValorTipoCambio.SetFocus
        Exit Function
    End If
    
    
    If Trim(txtMontoRecibido.Value) = "" Then
        MsgBox "Debe ingresar el Monto!", vbCritical, gstrNombreEmpresa
        txtMontoRecibido.SetFocus
        Exit Function
    End If
    
    If Trim(txtObservacion.Text) = "" Then
        MsgBox "Debe ingresar una Observacion!", vbCritical, gstrNombreEmpresa
        txtObservacion.SetFocus
        Exit Function
    End If
    
    
    If Trim(txtNumVoucher.Text) = "" And optTipoLiquidacion(1).Value = True Then
        MsgBox "Debe ingresar el número de Voucher!", vbCritical, gstrNombreEmpresa
        txtNumVoucher.SetFocus
        Exit Function
    End If
    
    TodoOk = True
'    If cboFondo.ListIndex <= -1 Then
'        MsgBox "Debe seleccionar la SIV.", vbCritical, Me.Caption
'        If cboFondo.Enabled Then cboFondo.SetFocus
'        Exit Function
'    End If
    
    
End Function

Public Sub Cancelar()
    
    cmdOpcion.Visible = True
    cmdReservar.Visible = True
    Call Limpiar
    With tabTipoCambio
        .TabEnabled(1) = False
        .TabEnabled(0) = True
        .TabEnabled(2) = False
        .Tab = 0
    End With
    
    cmdReservar.Visible = True
        
End Sub

Private Sub LlenarFormulario(strModo As String)

    Dim adoRegistro               As ADODB.Recordset, adoTemporal       As ADODB.Recordset
    Dim strSQL                    As String, strOperacion               As String
    Dim strFechaOperacionM        As String, strDatoConsultaCli As String
    Dim strDatoConsultaValor      As String, strDatoTipo As String
    Dim strDatoCli                As String, intRegistro As Integer
    Dim strNumDoc                 As String, strFechaTipoCambio As String
    
      
    Select Case strModo
    
        Case Reg_Adicion
                     
                     
            dtpFechaCambio.Value = gdatFechaActual
            dtpFechaValor.Value = gdatFechaActual
            strFechaTipoCambio = Convertyyyymmdd(dtpFechaCambio.Value)
            'optTipo(0).Value = True ' por defecto
            Call optTipo_Click(0)
          
                        txtCodParticipe.SetFocus
          
          'strNumDoc = ObtenerSecuencialInversionOperacion(strCodFondo, Valor_Documento)
          'txtNDocumento.Text = strNumDoc
          'txtNDocumento.Enabled = False
          
'            adoComm.CommandText = "SELECT FechaTipoCambio,ValorTipoCambioCompra,ValorTipoCambioVenta FROM TipoCambioFondoTemporal " & _
'                "WHERE FechaTipoCambio = '" & strFechaTipoCambio & "'"
                
            Set adoRegistro = adoComm.Execute
            
'            If Not adoRegistro.EOF Then
'                txtValorTipoCambio.Text = IIf(CStr(adoRegistro("ValorTipoCambioCompra")) = 0, 1, CStr(adoRegistro("ValorTipoCambioCompra")))
'                txtValorTipoCambio.SetFocus
'            Else
'                txtValorTipoCambio.Text = 1
'                txtValorTipoCambio.SetFocus
'            End If
            
            txtValorTipoCambioCosto.Text = txtValorTipoCambio.Value

            adoRegistro.Close: Set adoRegistro = Nothing
            txtMontoRecibido.Text = "0"
            txtMontoEntregado.Text = "0"
            optTipo(0).Value = vbChecked
            optTipo(1).Value = vbUnchecked
            
            
        Case Reg_Edicion
            
              strSQL = "SELECT  * FROM TesoreriaOperacion  " _
                 & "WHERE CodFondo= '" & gstrCodFondoContable & "'   AND CodAdministradora   = '" & gstrCodAdministradora & "' AND NumOperacion='" & Trim(tdgConsulta.Columns(0).Value) & "'"
                              
                Set adoRegistro = New ADODB.Recordset
                With adoComm
                    .CommandText = strSQL
                    Set adoRegistro = .Execute
                   
                    txtCodParticipe.Text = adoRegistro("CodParticipe").Value
                    Call txtCodParticipe_LostFocus
                    
                    lblNumOperacion.Caption = adoRegistro("NumOperacion").Value
                    dtpFechaCambio.Value = adoRegistro("FechaOperacion").Value
                    dtpFechaValor.Value = adoRegistro("FechaObligacion").Value
                    strFechaTipoCambio = Convertyyyymmdd(dtpFechaCambio.Value)
                    
                    If adoRegistro("CodMoneda").Value = "01" Then
                        
                        optTipo(0).Value = vbChecked
                        optTipo(1).Value = vbUnchecked
                        
                    Else
                    
                        optTipo(0).Value = vbUnchecked
                        optTipo(1).Value = vbChecked
                    
                    End If
                    
                    If adoRegistro("ModalidadOperacion").Value = "01" Then
                        
                        optTipoLiquidacion(0).Value = vbChecked
                        optTipoLiquidacion(1).Value = vbUnchecked
                        
                    Else
                    
                        optTipoLiquidacion(0).Value = vbUnchecked
                        optTipoLiquidacion(1).Value = vbChecked
                    
                    End If
                    
                    txtValorTipoCambio.Text = adoRegistro("ValorTipoCambioCliente").Value
                    txtMontoRecibido.Text = adoRegistro("MontoOperacion").Value
                    txtMontoEntregado.Text = adoRegistro("MontoOperacionCambio").Value
                    
                    txtValorTipoCambioCosto.Text = adoRegistro("ValorTipoCambioTesoreria").Value
                    txtMontoEntregadoBanco.Text = adoRegistro("MontoOperacionTesoreria").Value
                    txtMontoRecibidoBanco.Text = adoRegistro("MontoOperacionCambioTesoreria").Value
                    
                    txtObservacion.Text = adoRegistro("DescripObservacion").Value
                   
                    adoRegistro.Close: Set adoRegistro = Nothing
                End With
            
            
             
    End Select
    
 End Sub

Private Function TodoOkReversar() As Boolean

    TodoOkReversar = False
        
    If tdgConsulta.SelBookmarks.Count = 0 Or tdgConsulta.SelBookmarks.Count > 1 Then
        MsgBox "Debe seleccionar un registro para reversar", vbCritical, gstrNombreEmpresa
        Exit Function
    End If
        
    If cboEstadoOperacionBusqueda.ListIndex > -1 Then
        If strCodEstadoBusqueda <> "02" Then
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

Public Sub Reversar()
    
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
        
        If gdatFechaActual > tdgConsulta.Columns(1).Value Then 'cambiar la condicion por la fecha
            str_msg = str_msg + " Para continuar se requiere la Autorización, continuar?"
            If MsgBox(str_msg, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                Exit Sub
            Else
                'Call inputbox_Password(frmOrdenRentaVariableCliente, "*")
                str_pwd = InputBox(" Ingrese la contraseña de Autorización", App.Title)
                
                strSQL = "SELECT  * FROM AuxiliarParametro WHERE CodTipoParametro = 'PWDORD' " _
                 & "AND CodParametro= '01' and ValorParametro = '" & str_pwd & "'"
                              
                Set adoRegistro = New ADODB.Recordset
                With adoComm
                    .CommandText = strSQL
                    Set adoRegistro = .Execute
                    If adoRegistro.EOF Then
                        MsgBox "La contraseña ingresada es incorrecta...", vbCritical, Me.Caption
                        Exit Sub
                    End If
                    adoRegistro.Close: Set adoRegistro = Nothing
                End With
                
            End If
        End If
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
            Valor_NumOpeTesoreria & "','" & strNumOperacion & "','" & Space(1) + motivo + Space(1) + "') }"
            
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

Private Sub Form_Resize()
    Call AutoAjustarGrillas
End Sub

Private Sub cmdReservar_Click()
    Call Reversar
End Sub


Private Sub optTipo_Click(index As Integer)

'    If Index = 0 Then
'
'    Else
'
'    End If

    strTipoMov = IIf(optTipo(0), "C", "V")
    
    If strTipoMov = "C" Then
        strValorCambio = gstrValorTipoCambioOperacion
    Else
        strValorCambio = gstrValorTipoCambioLiquidacionRC
    End If
    
    
    
    Call cboMoneda_Click
    
    Call txtValorTipoCambio_LostFocus
    
End Sub

Private Sub optTipoLiquidacion_Click(index As Integer)


    If index = 0 Then
        linDirecto.Visible = False
        lblFlechaTransitoria.Visible = True
        lblTransitoriaComitente.ForeColor = &H800000
        lblFechaComitente.ForeColor = &H800000
        fraLiquidacion.Visible = False
        
        cboCuentaBancoResultado.Visible = False
        LinUtilidad.Visible = False
        lblDescrip(22).Visible = False
        lblUtilidad.Visible = False
        
        lblDescrip(29).Caption = "Tipo de Cambio Tesoreria"
        
        lblDescrip(36).Left = 6630
        lblDescrip(36).Top = 4590
        
        txtMontoUtilidad.Left = 8880
        txtMontoUtilidad.Top = 4560
       
        lblSignoMonedaUtilidad.Left = 11280
        lblSignoMonedaUtilidad.Top = 4620
        
    Else
        linDirecto.Visible = True
        lblFlechaTransitoria.Visible = False
        lblTransitoriaComitente.ForeColor = &H80&
        lblFechaComitente.ForeColor = &H80&
        fraLiquidacion.Visible = True
        
        cboCuentaBancoResultado.Visible = True
        LinUtilidad.Visible = True
        lblDescrip(22).Visible = True
        lblUtilidad.Visible = True

        
        lblDescrip(29).Caption = "Tipo de Cambio Banco"
        
        lblDescrip(36).Left = 6630
        lblDescrip(36).Top = 6150
        
        txtMontoUtilidad.Left = 8880
        txtMontoUtilidad.Top = 6120
       
        lblSignoMonedaUtilidad.Left = 11280
        lblSignoMonedaUtilidad.Top = 6180
        
    End If
    
    

End Sub

Private Sub tdgConsulta_HeadClick(ByVal ColIndex As Integer)
    Static numColindex As Integer

    tdgConsulta.Splits(0).Columns(numColindex).HeadingStyle.ForegroundPicture = Null

    Call OrdenarDBGrid(ColIndex, adoConsulta, tdgConsulta)
    
    numColindex = ColIndex
End Sub

Private Sub txtCodParticipeBusqueda_LostFocus()
    Dim rst As Boolean
    Dim adoRegistro As ADODB.Recordset
    Dim strSQL As String

    If Trim(txtCodParticipeBusqueda.Text) <> Valor_Caracter Then
        rst = False

        txtCodParticipeBusqueda.Text = Format(txtCodParticipeBusqueda.Text, "00000000000000000000")

        With adoComm
            Set adoRegistro = New ADODB.Recordset

            .CommandText = ""
            strSQL = "SELECT CodParticipe,AP1.DescripParametro TipoIdentidad,NumIdentidad,DescripParticipe,FechaIngreso,TipoIdentidad CodTipoIdentidad,AP2.DescripParametro TipoMancomuno "
            strSQL = strSQL & "FROM ParticipeContrato JOIN AuxiliarParametro AP1 ON(AP1.CodParametro=ParticipeContrato.TipoIdentidad AND AP1.CodTipoParametro='TIPIDE') "
            strSQL = strSQL & "JOIN AuxiliarParametro AP2 ON(AP2.CodParametro=ParticipeContrato.TipoMancomuno AND AP2.CodTipoParametro='TIPMAN') "
            strSQL = strSQL & "WHERE CodParticipe='" & Trim(txtCodParticipeBusqueda.Text) & "'"
            .CommandText = strSQL
            Set adoRegistro = .Execute

            Do Until adoRegistro.EOF
                rst = True
                lblDescripParticipeBusqueda.Caption = Trim(adoRegistro("DescripParticipe"))
                strCodParticipeBusqueda = txtCodParticipeBusqueda.Text
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

Private Sub txtMontoEntregado_LostFocus()

    txtMontoRecibidoBanco.Text = txtMontoEntregado.Text
    
    Call CalcularMontoRecibido

End Sub

Private Sub txtMontoEntregadoBanco_LostFocus()
    
    Call CalcularMontoRecibidoBanco
    
End Sub

Private Sub txtMontoRecibido_LostFocus()
 
   
    Call CalcularMontoEntregado
    
    txtMontoRecibidoBanco.Text = txtMontoEntregado.Text
'
'    Call CalcularMontoRecibido
    

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
                lblNombreCliente = Trim(adoRegistro("DescripParticipe"))
                strCodParticipe = txtCodParticipe
                adoRegistro.MoveNext
                
            Loop
            adoRegistro.Close: Set adoRegistro = Nothing
            
            If Not rst Then
                strCodParticipe = ""
                txtCodParticipe = ""
                lblNombreCliente = ""
                MsgBox "Codigo de Cliente Incorrecto", vbCritical, Me.Caption
            End If
            
        End With
    Else
        strCodParticipe = ""
        txtCodParticipe = ""
        lblNombreCliente = ""
    End If

End Sub

Public Sub Limpiar()

    Me.txtCodParticipe.Text = Valor_Caracter
    Me.lblNombreCliente.Caption = Valor_Caracter
    Me.txtValorTipoCambio.Text = Valor_Caracter
    Me.txtMontoRecibido.Text = Valor_Caracter
    Me.txtMontoEntregado.Text = Valor_Caracter
    Me.txtObservacion.Text = Valor_Caracter
    Me.txtCodParticipeBusqueda.Text = Valor_Caracter
    Me.txtNumVoucher.Text = Valor_Caracter
    txtMontoEntregadoBanco.Text = Valor_Caracter
    txtMontoRecibidoBanco.Text = Valor_Caracter
    txtMontoUtilidad.Text = Valor_Caracter

    If cboMoneda.ListCount > 0 Then cboMoneda.ListIndex = 0
    
    optTipo(0).Value = True
    optTipo(1).Value = False
    optTipoLiquidacion(0).Value = True
    optTipoLiquidacion(1).Value = False
    
    Call txtCodParticipe_LostFocus
    Call txtCodParticipeBusqueda_LostFocus
    
End Sub

Private Sub txtMontoRecibidoBanco_LostFocus()

    Call CalcularMontoEntregadoBanco

End Sub

Private Sub txtValorTipoCambio_LostFocus()

'    Call txtMontoRecibido_LostFocus
    Call CalcularMontoEntregado
    
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

Private Function TodoOkProceso() As Boolean
        
    TodoOkProceso = False
        
    If tdgConsulta.SelBookmarks.Count - 1 = -1 Then
        MsgBox "Debe seleccionar registros para procesar!", vbCritical, gstrNombreEmpresa
        Exit Function
    End If
    
    If cboEstadoOperacionBusqueda.ListIndex >= 0 Then
        If tdgConsulta.Columns("EstadoOperacion") <> Estado_Acuerdo_Ingresado Then
            MsgBox "No se pueden procesar las operaciones seleccionadas!", vbCritical, gstrNombreEmpresa
            Exit Function
        End If
    End If
        
    '*** Si todo paso OK ***
    TodoOkProceso = True
  
End Function

Private Sub txtValorTipoCambioCosto_LostFocus()

    Call CalcularMontoEntregadoBanco
    
End Sub
Private Sub CalcularUtilidad()
    
   
    'txtMontoUtilidad.Text = CStr(CDbl(txtMontoRecibido.Text) - CDbl(txtMontoEntregadoBanco.Text))



    If strTipoMov = "C" Then

       txtMontoUtilidad.Text = CStr(CDbl(txtMontoRecibido.Text) - CDbl(txtMontoEntregadoBanco.Text))
    Else
       
     txtMontoUtilidad.Text = CStr(CDbl(txtMontoRecibidoBanco.Text) - CDbl(txtMontoEntregado.Text))

    End If



End Sub
Private Sub CalcularMontoRecibido()

    If strCodMonedaRecibido <> Mid(strCodMonedaParPorDefecto, 3, 2) Then
        txtMontoRecibido.Text = CStr(CDbl(txtMontoEntregado.Value) * CDbl(txtValorTipoCambio.Text))
    Else
        txtMontoRecibido.Text = CStr(CDbl(txtMontoEntregado.Value) / CDbl(txtValorTipoCambio.Text))
    End If

    'Call CalcularMontoRecibidoBanco
    
End Sub
Private Sub CalcularMontoEntregado()

    If strCodMonedaRecibido <> Mid(strCodMonedaParPorDefecto, 1, 2) Then
        txtMontoEntregado.Text = CStr(CDbl(txtMontoRecibido.Value) / CDbl(txtValorTipoCambio.Text))
    Else
        txtMontoEntregado.Text = CStr(CDbl(txtMontoRecibido.Value) * CDbl(txtValorTipoCambio.Text))
    End If

    'Call CalcularMontoEntregadoBanco
    
End Sub

Private Sub CalcularMontoEntregadoBanco()
    
    'txtMontoEntregadoBanco.Text = txtMontoRecibido.Value
    
    If strCodMonedaEntregado <> Mid(strCodMonedaParPorDefecto, 1, 2) Then
        txtMontoEntregadoBanco.Text = CStr(CDbl(txtMontoRecibidoBanco.Value) / CDbl(txtValorTipoCambioCosto.Value))
    Else
        txtMontoEntregadoBanco.Text = CStr(CDbl(txtMontoRecibidoBanco.Value) * CDbl(txtValorTipoCambioCosto.Value))
    End If

    Call CalcularUtilidad

End Sub

Private Sub CalcularMontoRecibidoBanco()
    
    'txtMontoRecibidoBanco.Text = txtMontoEntregado.Value
    
    If strCodMonedaEntregado <> Mid(strCodMonedaParPorDefecto, 1, 2) Then
        txtMontoRecibidoBanco.Text = CStr(CDbl(txtMontoEntregadoBanco.Value) * CDbl(txtValorTipoCambioCosto.Value))
    Else
        txtMontoRecibidoBanco.Text = CStr(CDbl(txtMontoEntregadoBanco.Value) / CDbl(txtValorTipoCambioCosto.Value))
    End If

    Call CalcularUtilidad

End Sub

Public Sub SubImprimir(index As Integer)

    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    Dim strFechaDesde           As String, strFechaHasta        As String
    Dim strSeleccionRegistro    As String
 
    
    Select Case index
        Case 1
        
            gstrNameRepo = "OperacionCambioCliente"
            
            strSeleccionRegistro = "{MovimientoFondo.FechaRegistro} IN 'Fch1' TO 'Fch2'"
            gstrSelFrml = strSeleccionRegistro
            frmRangoFecha.Show vbModal
                        
            If gstrSelFrml = "0" Then
                Set frmReporte = New frmVisorReporte
    
                ReDim aReportParamS(4)
                ReDim aReportParamFn(2)
                ReDim aReportParamF(2)
                            
                aReportParamFn(0) = "Fondo"
                aReportParamFn(1) = "NombreEmpresa"
                aReportParamFn(2) = "Hora"
                           
                aReportParamF(0) = Trim(cboFondo.Text)
                aReportParamF(1) = gstrNombreEmpresa & Space(1)
                aReportParamF(1) = Time
                            
                aReportParamS(0) = strCodFondo
                aReportParamS(1) = gstrCodAdministradora
                'aReportParamS(2) = Convertyyyymmdd(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10))
                aReportParamS(2) = Convertyyyymmdd(gstrFchDel)
                'aReportParamS(3) = Convertyyyymmdd(CVDate(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10)))
                aReportParamS(3) = Convertyyyymmdd(gstrFchAl)
                aReportParamS(4) = strCodEstadoBusqueda
                
            End If
                                     
    End Select

    If gstrSelFrml <> "0" Then Exit Sub
    
    gstrSelFrml = Valor_Caracter
    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"

    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())

    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
    frmReporte.Show vbModal

    Set frmReporte = Nothing

    Screen.MousePointer = vbNormal
    
End Sub
Public Sub SubImprimir2(index As Integer)
    
     Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    Dim strFechaMovimBCDesde           As String, strFechaMovimBCHasta        As String
    Dim datFechaSiguiente              As String
    Dim strSeleccionRegistro           As String
 
     
     
 
     
      If Not IsNull(dtpFechaDesde.Value) Or Not IsNull(dtpFechaHasta.Value) Then
        strFechaMovimBCDesde = Convertyyyymmdd(dtpFechaDesde.Value)
        datFechaSiguiente = DateAdd("d", 1, dtpFechaHasta.Value)
        strFechaMovimBCHasta = Convertyyyymmdd(datFechaSiguiente)
     End If
     
            gstrNameRepo = "OperacionCambioClienteGrilla"
            
            gstrSelFrml = strSeleccionRegistro
                        
           
                Set frmReporte = New frmVisorReporte
    
                ReDim aReportParamS(5)
                ReDim aReportParamFn(4)
                ReDim aReportParamF(4)
                            
                aReportParamFn(0) = "Fondo"
                aReportParamFn(1) = "NombreEmpresa"
                aReportParamFn(2) = "Hora"
                aReportParamFn(3) = "Usuario"
                aReportParamFn(4) = "Estado"
                           
                aReportParamF(0) = Trim(cboFondo.Text)
                aReportParamF(1) = gstrNombreEmpresa & Space(1)
                aReportParamF(2) = Time
                aReportParamF(3) = gstrLogin
                aReportParamF(4) = cboEstadoOperacionBusqueda.Text
                            
                aReportParamS(0) = strCodFondo
                aReportParamS(1) = gstrCodAdministradora
                aReportParamS(2) = strFechaMovimBCDesde
                aReportParamS(3) = strFechaMovimBCHasta
                aReportParamS(4) = strCodEstadoBusqueda
                aReportParamS(5) = strCodParticipeBusqueda
    
    gstrSelFrml = Valor_Caracter
    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"

    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())

    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
    frmReporte.Show vbModal

    Set frmReporte = Nothing

    Screen.MousePointer = vbNormal
    


End Sub

Private Sub AutoAjustarGrillas()
    
    Dim i As Integer
    
    If Not adoConsulta.EOF Then
        If adoConsulta.RecordCount > 0 Then
            For i = 1 To tdgConsulta.Columns.Count - 1
                tdgConsulta.Columns(i).AutoSize
            Next
            
            tdgConsulta.Columns(0).AutoSize
            tdgConsulta.Columns(14).AutoSize
        End If
    End If
    
    tdgConsulta.Refresh

End Sub
