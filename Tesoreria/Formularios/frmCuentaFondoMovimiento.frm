VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{830D5A06-9B70-4F8C-98B6-7A19C4E7760B}#1.0#0"; "TAMControls.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmCuentaFondoMovimiento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transferencias entre Cuentas"
   ClientHeight    =   9045
   ClientLeft      =   240
   ClientTop       =   510
   ClientWidth     =   12705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   12705
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
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
      Picture         =   "frmCuentaFondoMovimiento.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   8250
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
      Left            =   6360
      Picture         =   "frmCuentaFondoMovimiento.frx":0671
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   8250
      Width           =   1200
   End
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   10560
      TabIndex        =   56
      Top             =   8250
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
      Left            =   360
      TabIndex        =   55
      Top             =   8250
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   1296
      Buttons         =   4
      Caption0        =   "&Nuevo"
      Tag0            =   "0"
      Visible0        =   0   'False
      ToolTipText0    =   "Nuevo"
      Caption1        =   "&Consultar"
      Tag1            =   "1"
      Visible1        =   0   'False
      ToolTipText1    =   "Consultar"
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
   Begin TabDlg.SSTab tabCuenta 
      Height          =   8235
      Left            =   0
      TabIndex        =   0
      Top             =   -30
      Width           =   12645
      _ExtentX        =   22304
      _ExtentY        =   14526
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
      TabPicture(0)   =   "frmCuentaFondoMovimiento.frx":0C96
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraCriterioBusqueda"
      Tab(0).Control(1)=   "tdgConsulta"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Registro de Movimiento"
      TabPicture(1)   =   "frmCuentaFondoMovimiento.frx":0CB2
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraCuenta(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdCancelarReg"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdRegistrar"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cboFormaPago"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin VB.ComboBox cboFormaPago 
         Height          =   315
         Left            =   7890
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   6840
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.CommandButton cmdRegistrar 
         Caption         =   "&Registrar"
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
         Picture         =   "frmCuentaFondoMovimiento.frx":0CCE
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Registrar el Movimiento"
         Top             =   7350
         Width           =   1200
      End
      Begin VB.CommandButton cmdCancelarReg 
         Caption         =   "&Cancelar"
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
         Left            =   10560
         Picture         =   "frmCuentaFondoMovimiento.frx":1217
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Registrar el Movimiento"
         Top             =   7350
         Width           =   1200
      End
      Begin VB.Frame fraCriterioBusqueda 
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
         Left            =   -74670
         TabIndex        =   31
         Top             =   540
         Width           =   11985
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
            Left            =   10200
            Picture         =   "frmCuentaFondoMovimiento.frx":1779
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   1050
            Width           =   1200
         End
         Begin VB.ComboBox cboFondo 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   600
            Width           =   6270
         End
         Begin VB.ComboBox cboEstado 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   1050
            Width           =   2565
         End
         Begin MSComCtl2.DTPicker dtpFechaMovimBCDesde 
            Height          =   315
            Left            =   1680
            TabIndex        =   34
            Top             =   1500
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
            Format          =   207159297
            CurrentDate     =   38785
         End
         Begin MSComCtl2.DTPicker dtpFechaMovimBCHasta 
            Height          =   315
            Left            =   4365
            TabIndex        =   35
            Top             =   1500
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
            Format          =   207159297
            CurrentDate     =   38785
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
            Left            =   360
            TabIndex        =   39
            Top             =   1530
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
            Index           =   25
            Left            =   3690
            TabIndex        =   38
            Top             =   1560
            Width           =   510
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
            Index           =   29
            Left            =   360
            TabIndex        =   37
            Top             =   1080
            Width           =   600
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
            Index           =   10
            Left            =   360
            TabIndex        =   36
            Top             =   630
            Width           =   540
         End
      End
      Begin VB.Frame fraCuenta 
         Caption         =   "Registro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6735
         Index           =   1
         Left            =   360
         TabIndex        =   4
         Top             =   570
         Width           =   11850
         Begin VB.CheckBox chkModificaTC2 
            Caption         =   "FijarT/C"
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
            Left            =   6240
            TabIndex        =   57
            Top             =   5880
            Value           =   1  'Checked
            Width           =   1575
         End
         Begin VB.ComboBox cboCuentaDestino 
            Height          =   315
            Left            =   3330
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   4320
            Width           =   5385
         End
         Begin VB.ComboBox cboBancoDestino 
            Height          =   315
            Left            =   3330
            Style           =   2  'Dropdown List
            TabIndex        =   51
            Top             =   3870
            Width           =   5385
         End
         Begin VB.ComboBox cboCuentaOrigen 
            Height          =   315
            Left            =   3330
            Style           =   2  'Dropdown List
            TabIndex        =   50
            Top             =   1890
            Width           =   5385
         End
         Begin VB.ComboBox cboBancoOrigen 
            Height          =   315
            Left            =   3330
            Style           =   2  'Dropdown List
            TabIndex        =   49
            Top             =   1440
            Width           =   5385
         End
         Begin VB.ComboBox cboTipoMovimiento 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1500
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   420
            Visible         =   0   'False
            Width           =   2955
         End
         Begin VB.CheckBox chkModificaTC 
            Caption         =   "FijarT/C"
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
            Left            =   10200
            TabIndex        =   6
            Top             =   7200
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CheckBox ChkMonDiferente 
            Caption         =   "Transacción en distinta moneda"
            Enabled         =   0   'False
            Height          =   255
            Left            =   840
            TabIndex        =   5
            Top             =   7260
            Width           =   3255
         End
         Begin TAMControls.TAMTextBox txtMontoMovimientoOrigen 
            Height          =   315
            Left            =   7560
            TabIndex        =   27
            Top             =   2400
            Width           =   1935
            _ExtentX        =   3413
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
            Container       =   "frmCuentaFondoMovimiento.frx":1CE1
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
         Begin TAMControls.TAMTextBox txtMontoMovimientoDestino 
            Height          =   315
            Left            =   7530
            TabIndex        =   28
            Top             =   4830
            Width           =   1935
            _ExtentX        =   3413
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
            Container       =   "frmCuentaFondoMovimiento.frx":1CFD
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
         Begin TAMControls.TAMTextBox txtTipoCambio 
            Height          =   315
            Left            =   3300
            TabIndex        =   29
            Top             =   5820
            Width           =   1575
            _ExtentX        =   2778
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
            Container       =   "frmCuentaFondoMovimiento.frx":1D19
            Text            =   "0.00000000"
            Decimales       =   8
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   12648447
            EnterTab        =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   8
         End
         Begin TAMControls.TAMTextBox txtNumDocumento 
            Height          =   315
            Left            =   3300
            TabIndex        =   45
            Top             =   6270
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
            Container       =   "frmCuentaFondoMovimiento.frx":1D35
            Decimales       =   2
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   12648447
            EnterTab        =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   2
         End
         Begin VB.Label lblMonedaDestino 
            AutoSize        =   -1  'True
            Caption         =   "(USD)"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   9540
            TabIndex        =   54
            Top             =   5250
            Width           =   510
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblMonedaOrigen 
            AutoSize        =   -1  'True
            Caption         =   "(USD)"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   9540
            TabIndex        =   53
            Top             =   2850
            Width           =   525
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblNumOperacion 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "GENERADO"
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
            Height          =   315
            Left            =   8970
            TabIndex        =   48
            Top             =   420
            Width           =   1755
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
            Index           =   16
            Left            =   7380
            TabIndex        =   47
            Top             =   450
            Width           =   1290
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Referencia Banco"
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
            Left            =   690
            TabIndex        =   46
            Top             =   6300
            Width           =   1965
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Transferencia"
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
            Left            =   5640
            TabIndex        =   30
            Top             =   6300
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Banco"
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
            Left            =   750
            TabIndex        =   26
            Top             =   3870
            Width           =   555
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Cuenta Transferencia de"
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
            Left            =   720
            TabIndex        =   25
            Top             =   4320
            Width           =   1995
         End
         Begin VB.Label lblDescrip 
            Alignment       =   2  'Center
            BackColor       =   &H8000000B&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Destino"
            ForeColor       =   &H00800000&
            Height          =   300
            Index           =   2
            Left            =   705
            TabIndex        =   24
            Top             =   3330
            Width           =   10140
         End
         Begin VB.Label lblDescrip 
            Alignment       =   2  'Center
            BackColor       =   &H8000000B&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Origen"
            ForeColor       =   &H00800000&
            Height          =   300
            Index           =   1
            Left            =   720
            TabIndex        =   23
            Top             =   960
            Width           =   10020
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Banco"
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
            Left            =   780
            TabIndex        =   22
            Top             =   1440
            Width           =   555
         End
         Begin VB.Label lblDescripDestino 
            Caption         =   "Monto Transferencia"
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
            Left            =   5610
            TabIndex        =   21
            Top             =   4890
            Width           =   1785
         End
         Begin VB.Label lblMonedaDestino 
            AutoSize        =   -1  'True
            Caption         =   "(USD)"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   9540
            TabIndex        =   20
            Top             =   4860
            Width           =   510
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblDescripOrigen 
            Caption         =   "Monto Transferencia"
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
            Height          =   270
            Left            =   5610
            TabIndex        =   19
            Top             =   2460
            Width           =   1830
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblMonedaOrigen 
            AutoSize        =   -1  'True
            Caption         =   "(USD)"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   9540
            TabIndex        =   18
            Top             =   2430
            Width           =   525
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "XXXXX"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   13
            Left            =   210
            TabIndex        =   17
            Top             =   7230
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "T/C Arbitraje "
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
            Left            =   690
            TabIndex        =   16
            Top             =   5880
            Width           =   1170
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   0
            X1              =   750
            X2              =   10890
            Y1              =   5610
            Y2              =   5610
         End
         Begin VB.Label lblDescripTCArbitraje 
            AutoSize        =   -1  'True
            Caption         =   "(USD/PEN)"
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
            Left            =   4920
            TabIndex        =   15
            Top             =   5880
            Width           =   1080
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Cuenta Transferencia de"
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
            Index           =   15
            Left            =   750
            TabIndex        =   14
            Top             =   1950
            Width           =   2205
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   1
            X1              =   750
            X2              =   10890
            Y1              =   5640
            Y2              =   5640
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
            Index           =   23
            Left            =   5610
            TabIndex        =   13
            Top             =   5280
            Width           =   1440
         End
         Begin VB.Label lblSaldoCuentaTransferencia 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            Height          =   315
            Left            =   7530
            TabIndex        =   12
            Top             =   5250
            Width           =   1935
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo"
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
            Left            =   720
            TabIndex        =   11
            Top             =   480
            Visible         =   0   'False
            Width           =   390
         End
         Begin VB.Label lblFechaMovimiento 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   5640
            TabIndex        =   10
            Top             =   420
            Width           =   1335
         End
         Begin VB.Label lblSaldoDisponible 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            Height          =   315
            Left            =   7560
            TabIndex        =   9
            Top             =   2820
            Width           =   1935
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
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
            Left            =   4860
            TabIndex        =   8
            Top             =   465
            Width           =   540
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
            Left            =   5610
            TabIndex        =   7
            Top             =   2850
            Width           =   1440
         End
      End
      Begin TAMControls.ucBotonEdicion cmdAccion 
         Height          =   390
         Left            =   -67770
         TabIndex        =   1
         Top             =   7170
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
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Height          =   4965
         Left            =   -74670
         OleObjectBlob   =   "frmCuentaFondoMovimiento.frx":1D51
         TabIndex        =   60
         Top             =   2820
         Width           =   11955
      End
   End
   Begin TAMControls.ucBotonEdicion cmdSalir2 
      Height          =   390
      Left            =   10440
      TabIndex        =   2
      Top             =   9960
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
      Left            =   840
      TabIndex        =   3
      Top             =   9960
      Visible         =   0   'False
      Width           =   4200
      _ExtentX        =   7408
      _ExtentY        =   688
      Buttons         =   3
      Caption0        =   "&Nuevo"
      Tag0            =   "0"
      Visible0        =   0   'False
      ToolTipText0    =   "Nuevo"
      Caption1        =   "&Consultar"
      Tag1            =   "1"
      Visible1        =   0   'False
      ToolTipText1    =   "Consultar"
      Caption2        =   "&Buscar"
      Tag2            =   "5"
      Visible2        =   0   'False
      ToolTipText2    =   "Buscar"
      UserControlHeight=   390
      UserControlWidth=   4200
   End
End
Attribute VB_Name = "frmCuentaFondoMovimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrFondo()              As String, arrEstado()      As String
Dim arrMoneda1()            As String, arrBanco1()          As String
Dim arrMoneda2()            As String, arrBanco2()          As String
Dim arrCuentaActivo()       As String, arrTipoRemunerada()  As String
Dim arrBancoOrigen()        As String, arrBancoDestino()    As String
Dim arrTasa()               As String, arrPorMonto()        As String
Dim arrTipoMovimiento()     As String, arrCalculo()         As String
Dim arrCuentaFondo()        As String
Dim arrCuentaOrigen()       As String, arrCuentaDestino()   As String
Dim arrFormaPago()          As String
Dim strCodFondo             As String, strCodTipoCuenta     As String
Dim strCodCuentaActivo      As String, strCodTipoRemunerada As String
Dim strCodTasa              As String, strCodPorMonto       As String
Dim strCodTipoMovimiento    As String, strCodCalculo        As String
Dim strCodFile              As String, strCodAnalitica      As String
Dim strCodCuentaFondo       As String, strCodFileCuenta     As String
Dim strCodAnaliticaCuenta   As String, strSignoMoneda       As String
Dim strCodMoneda            As String
Dim strCodMonedaCuenta      As String, strSQL               As String
Dim strSignoMonedaCuenta    As String, strCodBancoCuenta    As String
Dim strCodSignoMonedaCuenta As String, strCodBanco          As String
Dim strCodSignoMoneda       As String
Dim strEstado               As String
Dim dblSaldoDisponible      As Double   'HMC
Dim strModalidadCambio      As String
Dim strCodContraparte       As String
Dim strEstadoOperacion      As String

Dim strTipoContraparte      As String
Dim strCodigoCCI            As String   'ACC 02/11/2009
Dim strSignoMoneda1         As String
Dim strSignoMoneda2         As String
Dim strCodSignoMoneda1      As String
Dim strCodSignoMoneda2      As String
Dim strFormaPago            As String
Dim adoRegistroAux          As ADODB.Recordset
Dim adoConsulta             As ADODB.Recordset
Dim strCodMonedaFondo       As String
Dim strIndSentidoTipoCambio As String
Dim dblTipoCambioOpera      As Double
Dim intDiasDesplazamiento   As Integer
Dim strCodMonedaParEvaluacion As String
Dim strCodMonedaParPorDefecto As String
Dim datFechaConsulta          As Date
Dim strCodBancoDestino        As String
Dim strCodOrigenCuentaDestino   As String
Dim strCodOrigenCuentaOrigen   As String
Dim blnModifica                 As Boolean
Dim strCodEstado As String


Private Sub LlenarFormulario(strModo As String)

    Dim adoRecord As ADODB.Recordset
    Dim intRegistro As Integer
    
    Select Case strModo
        Case Reg_Adicion
        
            strIndSentidoTipoCambio = "M"

            lblFechaMovimiento.Caption = gdatFechaActual
            txtTipoCambio.Text = CStr(gdblTipoCambio)
            txtMontoMovimientoDestino.Text = "0"
            txtMontoMovimientoOrigen.Text = "0"
            lblSaldoDisponible.Caption = "0"
            lblSaldoCuentaTransferencia.Caption = "0"
            
            chkModificaTC.Value = vbChecked

            cboBancoOrigen.ListIndex = -1
            If cboBancoOrigen.ListCount > 0 Then cboBancoOrigen.ListIndex = 0

            cboCuentaOrigen.ListIndex = -1
            If cboCuentaOrigen.ListCount > 0 Then cboCuentaOrigen.ListIndex = 0

            cboBancoDestino.ListIndex = -1
            If cboBancoDestino.ListCount > 0 Then cboBancoDestino.ListIndex = 0

            cboCuentaDestino.ListIndex = -1
            If cboCuentaDestino.ListCount > 0 Then cboCuentaDestino.ListIndex = 0

            'Por defecto
            intRegistro = ObtenerItemLista(arrTipoMovimiento(), Codigo_Movimiento_Retiro)
            If intRegistro > 0 Then cboTipoMovimiento.ListIndex = intRegistro
            
            cboFormaPago.ListIndex = -1
            If cboFormaPago.ListCount > 0 Then cboFormaPago.ListIndex = 0
            
            txtNumDocumento.Enabled = True
            txtNumDocumento.Text = ""
            
                        
        Case Reg_Consulta, Reg_Edicion
        
            If strModo = Reg_Consulta Then
                cmdOpcion.Visible = False
                cmdRegistrar.Enabled = False
                txtMontoMovimientoOrigen.Enabled = False
                txtMontoMovimientoDestino.Enabled = False
                txtTipoCambio.Enabled = False
                txtNumDocumento.Enabled = False
            Else
                cmdOpcion.Visible = True
                cmdRegistrar.Enabled = True
                txtMontoMovimientoOrigen.Enabled = True
                txtMontoMovimientoDestino.Enabled = True
                txtTipoCambio.Enabled = True
                txtNumDocumento.Enabled = True
            End If
        
            Dim adoTemporal As ADODB.Recordset

            Set adoRecord = New ADODB.Recordset

            'ACR aqui va query de la tabla BancoCuentaMovimiento... accceder por NumOrdenCobroPago viene de MovimientoFondo ordenado por secuencial

            adoComm.CommandText = "{ call up_TEObtenerTransferenciaBancaria ('" & strCodFondo & "','" & gstrCodAdministradora & "','" & Trim(tdgConsulta.Columns("NumOperacion")) & "') }"
            Set adoRecord = adoComm.Execute

            If Not adoRecord.EOF Then
                
                lblNumOperacion.Alignment = vbCenter
                lblFechaMovimiento.Alignment = vbCenter
                
                lblNumOperacion.Caption = Trim(adoRecord("NumOperacion"))
                
                lblFechaMovimiento.Caption = Trim(adoRecord("FechaObligacion"))
                
                txtNumDocumento.Text = Trim(adoRecord("NumDocumento"))
                
                intRegistro = ObtenerItemLista(arrTipoMovimiento(), IIf(adoRecord("TipoMovimiento") = "E", Codigo_Movimiento_Deposito, Codigo_Movimiento_Retiro))
                If intRegistro > -1 Then cboTipoMovimiento.ListIndex = intRegistro
                
                '*** Origen ***
                cboBancoOrigen.ListIndex = -1
                intRegistro = ObtenerItemLista(arrBancoOrigen(), adoRecord("CodBancoClienteOrigen"))
                If intRegistro > -1 Then cboBancoOrigen.ListIndex = intRegistro
                
                cboCuentaOrigen.ListIndex = -1
                intRegistro = ObtenerItemLista(arrCuentaOrigen(), adoRecord("CodFileBancoClienteOrigen") + adoRecord("CodAnaliticaBancoClienteOrigen") + adoRecord("CodMoneda") + adoRecord("CodBancoClienteOrigen") + adoRecord("NumCuentaClienteOrigen") + adoRecord("CodCuentaBancoClienteOrigen"))
                If intRegistro > -1 Then cboCuentaOrigen.ListIndex = intRegistro
                
                lblDescrip(14).Enabled = False
                lblSaldoDisponible.Enabled = False
                lblMonedaOrigen(0).Caption = adoRecord("CodSignoMonedaOrigen")
                lblMonedaOrigen(1).Caption = adoRecord("CodSignoMonedaOrigen")
                
                txtMontoMovimientoOrigen.Text = IIf(CDbl(adoRecord("MontoOperacion")) < 0, CDbl(adoRecord("MontoOperacion")) * -1, adoRecord("MontoOperacion"))
                
                '*** Destino ***
                cboBancoDestino.ListIndex = -1
                intRegistro = ObtenerItemLista(arrBancoDestino(), adoRecord("CodBancoClienteDestino"))
                If intRegistro > -1 Then cboBancoDestino.ListIndex = intRegistro
                
                cboCuentaDestino.ListIndex = -1
                intRegistro = ObtenerItemLista(arrCuentaDestino(), adoRecord("CodFileBancoClienteDestino") + adoRecord("CodAnaliticaBancoClienteDestino") + adoRecord("CodMonedaCambio") + adoRecord("CodBancoClienteDestino") + adoRecord("NumCuentaClienteDestino") + adoRecord("CodCuentaBancoClienteDestino"))
                If intRegistro > -1 Then cboCuentaDestino.ListIndex = intRegistro
                
                lblSaldoCuentaTransferencia.Caption = "0"
                lblDescrip(23).Enabled = False
                lblSaldoCuentaTransferencia.Enabled = False
                lblMonedaDestino(0).Caption = adoRecord("CodSignoMonedaDestino")
                lblMonedaDestino(1).Caption = adoRecord("CodSignoMonedaDestino")
                
                txtMontoMovimientoDestino.Text = IIf(CDbl(adoRecord("MontoOperacionCambio")) < 0, CDbl(adoRecord("MontoOperacionCambio")) * -1, adoRecord("MontoOperacionCambio"))
                
                '-----
                txtTipoCambio.Text = adoRecord("ValorTipoCambioCliente")
                lblDescripTCArbitraje.Caption = "(" + Trim(adoRecord("CodSignoMonedaOrigen")) + "/" + Trim(adoRecord("CodSignoMonedaDestino")) + ")"

                intRegistro = ObtenerItemLista(arrFormaPago(), adoRecord("TipoMovimiento"))
                If intRegistro > 0 Then cboFormaPago.ListIndex = intRegistro
            
            End If

    End Select
    
    fraCuenta(1).Caption = Trim(cboFondo.Text)
    
End Sub

Public Sub Cancelar()

    cmdOpcion.Visible = True
    With tabCuenta
        .TabEnabled(0) = True
        .TabEnabled(1) = False
        .Tab = 0
    End With
   Call Buscar
    
    
End Sub

Private Sub CargarReportes()

    'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Transferencia entre cuentas"
'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Visible = True
'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Text = "Movimientos Caja/Cuenta Corriente"
    
End Sub

Public Sub Eliminar()

    
        If strEstado = Reg_Consulta Then
            
            strEstadoOperacion = "03"
            
            If Not TodoOkAnular() Then Exit Sub
            
            If MsgBox("Se procederá a anular la Transferencia Bancaria Nro. " & tdgConsulta.Columns("NumOperacion").Value & vbNewLine & vbNewLine & "Seguro de continuar ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            
                adoComm.CommandText = "UPDATE TesoreriaOperacion SET EstadoOperacion='" & strEstadoOperacion & "' " & _
                "WHERE NumOperacion='" & tdgConsulta.Columns("NumOperacion").Value & "' AND " & _
                "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' "
                
                adoConn.Execute adoComm.CommandText
                
                Call Buscar
                
            End If
            
        End If



End Sub

Private Function TodoOkAnular()

    TodoOkAnular = False

    If tdgConsulta.SelBookmarks.Count = 0 Or tdgConsulta.SelBookmarks.Count > 1 Then
        MsgBox "Debe seleccionar un registro para Anular", vbCritical + vbOKOnly, gstrNombreEmpresa
        Exit Function
    End If
        
    If cboEstado.ListIndex > -1 Then
        If strCodEstado = Estado_Caja_Confirmado Then
            MsgBox "No se puede anular un registro ya procesado", vbOKOnly + vbCritical, Me.Caption
            Exit Function
        ElseIf strCodEstado = Estado_Caja_Anulado Then
            MsgBox "Este registro ya esta anulado", vbOKOnly + vbCritical, Me.Caption
            Exit Function
         ElseIf strCodEstado = "05" Then 'falta la constante para reversado
            MsgBox "No se puede anular un registro ya Reversado", vbOKOnly + vbCritical, Me.Caption
            Exit Function
        End If
    Else
        MsgBox "Debe seleccionar algun estado de operacion", vbCritical + vbOKOnly, gstrNombreEmpresa
        If cboEstado.Enabled Then cboEstado.SetFocus
        Exit Function
    End If
        
    '*** Si todo paso OK ***
    TodoOkAnular = True

End Function


Public Sub Modificar()

    If strEstado = Reg_Consulta Then
        If tdgConsulta.Columns("EstadoOperacion") = Estado_Acuerdo_Ingresado Then strEstado = Reg_Edicion
        LlenarFormulario strEstado
        With tabCuenta
            .TabEnabled(0) = False
            '.TabEnabled(2) = True
            .Tab = 1
        End With
    End If
    
    
End Sub


Public Sub Salir()

    Unload Me
    
End Sub



Public Sub SubImprimir(index As Integer)

    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    Dim strFechaDesde           As String, strFechaHasta        As String


    Select Case index
        
        Case 1
        
        gstrNameRepo = "TransferenciaCuentaBancaria"
        frmRangoFecha.Show vbModal
            
        If gstrSelFrml = "0" Then
            
            Me.MousePointer = vbHourglass
            
            Set frmReporte = New frmVisorReporte
    
            ReDim aReportParamS(4)
            ReDim aReportParamFn(3)
            ReDim aReportParamF(3)
                        
            aReportParamFn(0) = "Usuario"
            aReportParamFn(1) = "Hora"
            aReportParamFn(2) = "NombreEmpresa"
            aReportParamFn(3) = "Fondo"
                        
            aReportParamF(0) = gstrLogin
            aReportParamF(1) = Format(Time(), "hh:mm:ss")
            aReportParamF(2) = gstrNombreEmpresa & Space(1)
            aReportParamF(3) = Trim(cboFondo.Text)
            
            strFechaDesde = Convertyyyymmdd(gstrFchDel)
            strFechaHasta = Convertyyyymmdd(gstrFchAl)
            
            aReportParamS(0) = strCodFondo
            aReportParamS(1) = gstrCodAdministradora
            aReportParamS(2) = strFechaDesde
            aReportParamS(3) = strFechaHasta
            aReportParamS(4) = strCodEstado
            
        End If
    
    End Select
   
    If gstrSelFrml <> "0" Then Exit Sub
   
    gstrSelFrml = Valor_Caracter
    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"

    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())

    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
    frmReporte.Show vbModal

    Set frmReporte = Nothing
    Me.MousePointer = vbNormal
    Screen.MousePointer = vbNormal

End Sub

Public Sub SubImprimir2(index As Integer)

    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    Dim strFechaMovimBCDesde           As String, strFechaMovimBCHasta        As String
    Dim datFechaSiguiente              As String
    
      If Not IsNull(dtpFechaMovimBCDesde.Value) Or Not IsNull(dtpFechaMovimBCHasta.Value) Then
        strFechaMovimBCDesde = Convertyyyymmdd(dtpFechaMovimBCDesde.Value)
        datFechaSiguiente = DateAdd("d", 1, dtpFechaMovimBCHasta.Value)
        strFechaMovimBCHasta = Convertyyyymmdd(datFechaSiguiente)
      End If

        
        gstrNameRepo = "TransferenciaCuentaBancariaGrilla"
            Me.MousePointer = vbHourglass
            
            Set frmReporte = New frmVisorReporte
    
            ReDim aReportParamS(4)
            ReDim aReportParamFn(3)
            ReDim aReportParamF(3)
                        
            aReportParamFn(0) = "Usuario"
            aReportParamFn(1) = "Hora"
            aReportParamFn(2) = "NombreEmpresa"
            aReportParamFn(3) = "Fondo"
                        
            aReportParamF(0) = gstrLogin
            aReportParamF(1) = Format(Time(), "hh:mm:ss")
            aReportParamF(2) = gstrNombreEmpresa & Space(1)
            aReportParamF(3) = Trim(cboFondo.Text)
            
            aReportParamS(0) = strCodFondo
            aReportParamS(1) = gstrCodAdministradora
            aReportParamS(2) = strFechaMovimBCDesde
            aReportParamS(3) = strFechaMovimBCHasta
            aReportParamS(4) = strCodEstado
            
        
    gstrSelFrml = Valor_Caracter
    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"

    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())

    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
    frmReporte.Show vbModal

    Set frmReporte = Nothing
    Me.MousePointer = vbNormal
    Screen.MousePointer = vbNormal

End Sub


Private Sub cboBancoDestino_Click()

    strCodBancoDestino = Valor_Caracter

    If cboBancoDestino.ListIndex < 0 Then Exit Sub
    strCodBancoDestino = Trim(arrBancoDestino(cboBancoDestino.ListIndex))

    '*** Cuentas del Fondo ***
    strSQL = "SELECT (CodFile + CodAnalitica + CodMoneda + CodBanco + (NumCuenta + replicate (' ',(30 - len(NumCuenta)))) + CodCuentaActivo ) CODIGO,(DescripCuenta + ' ' + NumCuenta) DESCRIP " & _
        "FROM BancoCuenta " & _
        "WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND CodBanco='" & strCodBancoDestino & "' "
        
    If cboBancoDestino.ListIndex > 0 Then
        strSQL = strSQL & " AND CodBanco = '" & arrBancoDestino(cboBancoDestino.ListIndex) & "'"
    End If
    
    CargarControlLista strSQL, cboCuentaDestino, arrCuentaDestino(), ""
    If cboCuentaDestino.ListCount > 0 Then cboCuentaDestino.ListIndex = 0

End Sub

Private Sub cboBancoOrigen_Click()

    strCodBanco = Valor_Caracter
    
    If cboBancoOrigen.ListIndex < 0 Then Exit Sub
    
    strCodBanco = Trim(arrBancoOrigen(cboBancoOrigen.ListIndex))

    '*** Cuentas del Fondo ***

    strSQL = "SELECT (CodFile + CodAnalitica + CodMoneda + CodBanco + (NumCuenta + replicate (' ',(30 - len(NumCuenta)))) + CodCuentaActivo ) CODIGO,(DescripCuenta + ' ' + NumCuenta) DESCRIP " & _
        "FROM BancoCuenta " & _
        "WHERE CodFondo = '" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND CodBanco='" & strCodBanco & "' "
        
    If cboBancoOrigen.ListIndex > 0 Then
        strSQL = strSQL & " AND CodBanco = '" & arrBancoOrigen(cboBancoOrigen.ListIndex) & "'"
    End If
    
    CargarControlLista strSQL, cboCuentaOrigen, arrCuentaOrigen(), ""
    If cboCuentaOrigen.ListCount > 0 Then cboCuentaOrigen.ListIndex = 0
       

End Sub

Private Sub cboCuentaDestino_Click()

    Dim curSaldoCuenta  As Currency
    Dim datFecha        As Date, datFechaSiguiente  As Date
    
    strCodFileCuenta = Valor_Caracter
    strCodAnaliticaCuenta = Valor_Caracter
    strCodCuentaFondo = Valor_Caracter
    strCodMonedaCuenta = Valor_Caracter
    lblSaldoCuentaTransferencia.Caption = "0"

    If cboCuentaDestino.ListIndex < 0 Then Exit Sub
    
    strCodFileCuenta = Trim(Mid(arrCuentaDestino(cboCuentaDestino.ListIndex), 1, 3))
    strCodAnaliticaCuenta = Trim(Mid(arrCuentaDestino(cboCuentaDestino.ListIndex), 4, 8))
    strCodCuentaFondo = Trim(Mid(arrCuentaDestino(cboCuentaDestino.ListIndex), 52, 10))
    strCodMonedaCuenta = Trim(Mid(arrCuentaDestino(cboCuentaDestino.ListIndex), 12, 2))
    'strCodOrigenCuentaDestino = Mid(Trim(arrCuentaDestino(cboCuentaDestino.ListIndex)), 47, 2)
    
    'strCodBancoCuenta = Trim(Mid(arrCuentaDestino(cboCuentaDestino.ListIndex), 14, 8))
    
    If lblFechaMovimiento.Caption <> "" Then
        datFecha = CVDate(lblFechaMovimiento.Caption)
        datFechaSiguiente = DateAdd("d", 1, datFecha)
    End If
    
    curSaldoCuenta = ObtenerSaldoFinalCuenta(strCodFondo, gstrCodAdministradora, strCodFileCuenta, strCodAnaliticaCuenta, Convertyyyymmdd(datFecha), Convertyyyymmdd(datFechaSiguiente), strCodCuentaFondo, strCodMonedaCuenta) ', strCodMonedaFondo)
    lblSaldoCuentaTransferencia.Caption = CStr(curSaldoCuenta)
    
    strSignoMonedaCuenta = ObtenerSignoMoneda(strCodMonedaCuenta)
    strCodSignoMonedaCuenta = ObtenerCodSignoMoneda(strCodMonedaCuenta)
    
'         'cargar el pop-up menu.
'    If strCodMonedaCuenta <> Valor_Caracter And strCodMoneda <> Valor_Caracter Then
'        If strCodMonedaCuenta = strCodMoneda Then
'            mnuSListPopup.Item(0).Enabled = False
'            mnuSListPopup.Item(2).Enabled = False
'            mnuSListPopup.Item(0).Caption = ""
'            mnuSListPopup.Item(2).Caption = ""
'            lblDescripTCArbitraje.Enabled = False
'            lblDescripTCArbitraje.Caption = strCodSignoMonedaCuenta
'        Else
'            strIndSentidoTipoCambio = "M"
'            mnuSListPopup.Item(0).Enabled = False
'            mnuSListPopup.Item(2).Enabled = True
'            lblDescripTCArbitraje.Enabled = True
'            mnuSListPopup.Item(0).Caption = "(" & strCodSignoMoneda & "/" & strCodSignoMonedaCuenta & ")"
'            mnuSListPopup.Item(2).Caption = "(" & strCodSignoMonedaCuenta & "/" & strCodSignoMoneda & ")"
'            lblDescripTCArbitraje.Caption = "(" & strCodSignoMoneda & "/" & strCodSignoMonedaCuenta & ")"
'        End If
'    Else
'        mnuSListPopup.Item(0).Caption = ""
'        mnuSListPopup.Item(2).Caption = ""
'        lblDescripTCArbitraje.Enabled = False
'    End If
    
    strCodMonedaParEvaluacion = strCodMoneda & strCodMonedaCuenta
    
    If strCodMoneda <> strCodMonedaCuenta Then
        strCodMonedaParPorDefecto = ObtenerMonedaParPorDefecto(gstrCodClaseTipoCambioOperacionFondo, strCodMonedaParEvaluacion)
    Else
        strCodMonedaParPorDefecto = strCodMonedaParEvaluacion
    End If

    datFechaConsulta = DateAdd("d", intDiasDesplazamiento, datFecha)
   
    If strCodMoneda <> strCodMonedaCuenta Then
        ChkMonDiferente.Value = vbChecked
        txtTipoCambio.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, datFechaConsulta, Mid(strCodMonedaParPorDefecto, 1, 2), Mid(strCodMonedaParPorDefecto, 3, 2)))
        txtTipoCambio.Enabled = True
        chkModificaTC2.Enabled = True
        chkModificaTC2.Value = vbUnchecked
        
        Call chkModificaTC2_Click
    Else
        ChkMonDiferente.Value = vbUnchecked
        txtTipoCambio.Text = "1"
        txtTipoCambio.Enabled = False
        chkModificaTC2.Enabled = False
        chkModificaTC2.Value = vbUnchecked
        
        Call chkModificaTC2_Click
    End If
    
    Call txtTipoCambio_KeyPress(vbKeyReturn)
    
    'lblDescrip(23).Caption = "Saldo Disponible" & Space(1) & strSignoMonedaCuenta
    
'    lblMonedaOrigen.Caption = strCodSignoMoneda
    lblMonedaDestino(0).Caption = strCodSignoMonedaCuenta
    lblMonedaDestino(1).Caption = strCodSignoMonedaCuenta
    

End Sub

'Private Sub cboCuentaFondo_Click()
'
'    Dim curSaldoCuenta  As Currency
'    Dim datFecha        As Date, datFechaSiguiente  As Date
'
'    strCodFileCuenta = Valor_Caracter
'    strCodAnaliticaCuenta = Valor_Caracter
'    strCodCuentaFondo = Valor_Caracter
'    strCodMonedaCuenta = Valor_Caracter
'    lblSaldoCuentaTransferencia.Caption = "0"
'
'    If cboCuentaFondo.ListIndex < 0 Then Exit Sub
'
'    strCodFileCuenta = Left(arrCuentaFondo(cboCuentaFondo.ListIndex), 3)
'    strCodAnaliticaCuenta = Mid(arrCuentaFondo(cboCuentaFondo.ListIndex), 4, 8)
'    strCodCuentaFondo = Trim(Mid(arrCuentaFondo(cboCuentaFondo.ListIndex), 12, 10))
'    strCodMonedaCuenta = Trim(Mid(arrCuentaFondo(cboCuentaFondo.ListIndex), 22, 2))
'    strCodBancoCuenta = Trim(Right(arrCuentaFondo(cboCuentaFondo.ListIndex), 8))
'
'    datFecha = CVDate(lblFechaMovimiento.Caption)
'    datFechaSiguiente = DateAdd("d", 1, datFecha)
'
'    curSaldoCuenta = ObtenerSaldoFinalCuenta(strCodFondo, gstrCodAdministradora, strCodFileCuenta, strCodAnaliticaCuenta, Convertyyyymmdd(datFecha), Convertyyyymmdd(datFechaSiguiente), strCodCuentaFondo, strCodMonedaCuenta, strCodMonedaFondo)
'    lblSaldoCuentaTransferencia.Caption = CStr(curSaldoCuenta)
'
'    strSignoMonedaCuenta = ObtenerSignoMoneda(strCodMonedaCuenta)
'    strCodSignoMonedaCuenta = ObtenerCodSignoMoneda(strCodMonedaCuenta)
'
'    lblDescripTCArbitraje.Caption = "(" & strCodSignoMoneda & "/" & strCodSignoMonedaCuenta & ")"
'
'    If strCodMoneda <> strCodMonedaCuenta Then
'        ChkMonDiferente.Value = vbChecked
'        txtTipoCambio.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, gdatFechaActual, strCodMoneda, strCodMonedaCuenta))
'        txtTipoCambio.Enabled = True
'    Else
'        ChkMonDiferente.Value = vbUnchecked
'        txtTipoCambio.Text = "1"
'        txtTipoCambio.Enabled = False
'    End If
'    Call txtTipoCambio_KeyPress(vbKeyReturn)
'
'
'    lblDescrip(23).Caption = "Saldo Disponible" & Space(1) & strSignoMonedaCuenta
'
'    lblMonedaOrigen.Caption = strCodSignoMoneda
'    lblMonedaDestino.Caption = strCodSignoMonedaCuenta
'
'End Sub


Private Sub cboCuentaOrigen_Click()
        
    Dim datFecha        As Date, datFechaSiguiente  As Date
    Dim curSaldoCuenta  As Double
        
    If cboCuentaOrigen.ListIndex < 0 Then Exit Sub

    strCodMoneda = Mid(Trim(arrCuentaOrigen(cboCuentaOrigen.ListIndex)), 12, 2)
       
    strCodCuentaActivo = Valor_Caracter
    strCodAnalitica = Valor_Caracter
    strCodFile = Valor_Caracter
    
    strCodCuentaActivo = Mid(Trim(arrCuentaOrigen(cboCuentaOrigen.ListIndex)), 52, 10)
    strCodAnalitica = Mid(Trim(arrCuentaOrigen(cboCuentaOrigen.ListIndex)), 4, 8)
    strCodFile = Mid(Trim(arrCuentaOrigen(cboCuentaOrigen.ListIndex)), 1, 3)
    
    'adoRecord("CodFileBancoClienteOrigen") + adoRecord("CodAnaliticaBancoClienteOrigen") + adoRecord("CodMoneda") + adoRecord("CodBancoClienteOrigen") + adoRecord("NumCuentaClienteOrigen") + adoRecord("CodCuentaBancoClienteOrigen")
    
    If lblFechaMovimiento.Caption <> "" Then
       datFecha = CVDate(lblFechaMovimiento.Caption)
       datFechaSiguiente = DateAdd("d", 1, datFecha)
    End If
    
                    
    '*** Obtener los saldos de la cuenta ***
    lblSaldoDisponible.Caption = "0"
    curSaldoCuenta = ObtenerSaldoFinalCuenta(strCodFondo, gstrCodAdministradora, strCodFile, strCodAnalitica, Convertyyyymmdd(datFecha), Convertyyyymmdd(datFechaSiguiente), strCodCuentaActivo, strCodMoneda) ', strCodMoneda)
    lblSaldoDisponible.Caption = curSaldoCuenta
     
     
    strSignoMoneda = ObtenerSignoMoneda(strCodMoneda)
    strCodSignoMoneda = ObtenerCodSignoMoneda(strCodMoneda)
    
    strCodMonedaParEvaluacion = strCodMoneda & strCodMonedaCuenta
    
    If strCodMoneda <> strCodMonedaCuenta Then
        strCodMonedaParPorDefecto = ObtenerMonedaParPorDefecto(gstrCodClaseTipoCambioOperacionFondo, strCodMonedaParEvaluacion)
    Else
        strCodMonedaParPorDefecto = strCodMonedaParEvaluacion
    End If
    
    datFechaConsulta = DateAdd("d", intDiasDesplazamiento, datFecha)
    
    If strCodMoneda <> strCodMonedaCuenta Then
       ChkMonDiferente.Value = vbChecked
       txtTipoCambio.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, datFechaConsulta, Mid(strCodMonedaParPorDefecto, 1, 2), Mid(strCodMonedaParPorDefecto, 3, 2)))
       txtTipoCambio.Enabled = True
       chkModificaTC2.Enabled = True
       chkModificaTC2.Value = vbUnchecked
       
       Call chkModificaTC2_Click
    Else
       ChkMonDiferente.Value = vbUnchecked
       txtTipoCambio.Text = "1"
       txtTipoCambio.Enabled = False
       chkModificaTC2.Enabled = False
       chkModificaTC2.Value = vbUnchecked
       
       Call chkModificaTC2_Click
    End If
    
    Call txtTipoCambio_KeyPress(vbKeyReturn)
         
    lblMonedaOrigen(0).Caption = strCodSignoMoneda
    lblMonedaOrigen(1).Caption = strCodSignoMoneda

End Sub

Private Sub cboFondo_Click()

    Dim adoRegistro     As ADODB.Recordset
    Dim strNumRucFondo  As String
        
    
    strCodFondo = Valor_Caracter
    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(arrFondo(cboFondo.ListIndex))
    
    strCodMonedaFondo = ""
    
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Fecha Vigente, Valor Cuota, Valor Cuota T-1, Moneda y Cantidad Inicial de Cuotas del Fondo ***
        .CommandText = "{ call up_ACSelDatosParametro(23,'" & strCodFondo & "','" & gstrCodAdministradora & "','000') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            lblFechaMovimiento.Caption = CStr(adoRegistro("FechaCuota"))
            gdatFechaActual = adoRegistro("FechaCuota")
            gstrFechaActual = Convertyyyymmdd(gdatFechaActual)
            
            strCodMonedaFondo = adoRegistro("CodMoneda")
            frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
         
            '*** Valores a las fechas de búsqueda
            dtpFechaMovimBCDesde.Value = gdatFechaActual
            dtpFechaMovimBCHasta.Value = dtpFechaMovimBCDesde.Value
            
            txtTipoCambio.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, gdatFechaActual, Trim(adoRegistro("CodMoneda")), Codigo_Moneda_Local))
            
            If CDbl(txtTipoCambio.Value) = 0 Then txtTipoCambio.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, DateAdd("d", -1, gdatFechaActual), Trim(adoRegistro("CodMoneda")), Codigo_Moneda_Local))
        End If
        
        adoRegistro.Close
        
        .CommandText = "{ call up_ACSelDatosParametro(24,'" & strCodFondo & "','" & gstrCodAdministradora & "') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            strNumRucFondo = CStr(adoRegistro("NumRucFondo"))
            strTipoContraparte = Codigo_Tipo_Persona_Portafolio
        End If
        
        adoRegistro.Close
        
        .CommandText = "SELECT CodPersona FROM InstitucionPersona"
        .CommandText = .CommandText & " WHERE "
        .CommandText = .CommandText & " TipoPersona   = '" & strTipoContraparte & "' AND "
        .CommandText = .CommandText & " TipoIdentidad = '" & Codigo_Tipo_Registro_Unico_Contribuyente & "' AND "
        .CommandText = .CommandText & " NumIdentidad  = '" & strNumRucFondo & "'"
    
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            strCodContraparte = adoRegistro("CodPersona")
        Else
            strCodContraparte = "00000000"
        End If
        
        adoRegistro.Close
        
    End With
    
   Call Buscar
    
    
End Sub


Private Sub cboFormaPago_Click()

    strFormaPago = Valor_Caracter
    
    If cboFormaPago.ListIndex <= 0 Then Exit Sub
    
    strFormaPago = "02" 'Trim(arrFormaPago(cboFormaPago.ListIndex))

End Sub


Private Sub cboEstado_Click()

    strCodEstado = Valor_Caracter
    
    If cboEstado.ListIndex < 0 Then Exit Sub

    strCodEstado = Trim(arrEstado(cboEstado.ListIndex))
    
  Call Buscar
    
    
End Sub


Private Sub cboTipoMovimiento_Click()

    Dim strSQL As String
    
    strCodTipoMovimiento = ""
    If cboTipoMovimiento.ListIndex < 0 Then Exit Sub
    
    'strCodTipoMovimiento = Trim(arrTipoMovimiento(cboTipoMovimiento.ListIndex))
        
    strCodTipoMovimiento = "02"
    If strCodTipoMovimiento = Codigo_Movimiento_Deposito Then
'        lblDescrip(22).Caption = "Transferir De ..."
'        lblDescrip(1).Caption = "Destino"
'        lblDescrip(2).Caption = "Origen"
        lblDescrip(15).Caption = "Transferencia hacia la cuenta ..."
        lblDescrip(3).Caption = "Transferencia desde la cuenta ..."
'        lblDescripOrigen.Caption = "Monto Transferencia Destino"
'        lblDescripDestino.Caption = "Monto Transferencia Origen"
    Else
'        lblDescrip(22).Caption = "Transferir A ..."
'        lblDescrip(1).Caption = "Origen"
'        lblDescrip(2).Caption = "Destino"
        lblDescrip(15).Caption = "Transferencia desde la cuenta ..."
        lblDescrip(3).Caption = "Transferencia hacia la cuenta ..."
        'fraInicio.Caption = "Origen"
        'fraFinal.Caption = "Destino"
 '       lblDescripOrigen.Caption = "Monto Transferencia Origen"
 '       lblDescripDestino.Caption = "Monto Transferencia Destino"
    End If
    
    
End Sub


Private Function TodoOk()

    TodoOk = False
    
    If cboTipoMovimiento.ListIndex = 0 Then
        MsgBox "Seleccione el tipo de movimiento.", vbCritical, gstrNombreEmpresa
        cboTipoMovimiento.SetFocus
        Exit Function
    End If
    
    If strCodTipoMovimiento = Codigo_Movimiento_Deposito Then
    
        If cboBancoOrigen.ListIndex = -1 Then
            MsgBox "Seleccione el Banco destino de la transferencia.", vbCritical, gstrNombreEmpresa
            cboBancoOrigen.SetFocus
            Exit Function
        End If
    
        If cboCuentaOrigen.ListIndex = -1 Then
            MsgBox "Seleccione la Cuenta destino de la transferencia.", vbCritical, gstrNombreEmpresa
            cboCuentaOrigen.SetFocus
            Exit Function
        End If
    
        If cboBancoDestino.ListIndex = -1 Then
            MsgBox "Seleccione el Banco origen de la transferencia.", vbCritical, gstrNombreEmpresa
            cboBancoDestino.SetFocus
            Exit Function
        End If
        
        If cboCuentaDestino.ListIndex = -1 Then
            MsgBox "Seleccione la Cuenta origen de la transferencia.", vbCritical, gstrNombreEmpresa
            cboCuentaDestino.SetFocus
            Exit Function
        End If
        
    
    Else
        
        If cboBancoOrigen.ListIndex = -1 Then
            MsgBox "Seleccione el Banco origen de la transferencia.", vbCritical, gstrNombreEmpresa
            cboBancoOrigen.SetFocus
            Exit Function
        End If
        
        If cboCuentaOrigen.ListIndex = -1 Then
            MsgBox "Seleccione la Cuenta origen de la transferencia.", vbCritical, gstrNombreEmpresa
            cboCuentaOrigen.SetFocus
            Exit Function
        End If
        
    
        If cboBancoDestino.ListIndex = -1 Then
            MsgBox "Seleccione el Banco destino de la transferencia.", vbCritical, gstrNombreEmpresa
            cboBancoOrigen.SetFocus
            Exit Function
        End If
        
        If cboCuentaDestino.ListIndex = -1 Then
            MsgBox "Seleccione la Cuenta destino de la transferencia.", vbCritical, gstrNombreEmpresa
            cboCuentaDestino.SetFocus
            Exit Function
        End If
   
    End If
    
    If CCur(txtTipoCambio.Value) = 0 Then
        MsgBox "Ingrese el tipo de cambio de la transferencia.", vbCritical, gstrNombreEmpresa
        txtTipoCambio.SetFocus
        Exit Function
    End If
    
  
    'Validando montos
    
    If CCur(txtMontoMovimientoDestino.Value) = 0 Or CCur(txtMontoMovimientoOrigen.Value) = 0 Then
        MsgBox "Monto del movimiento no puede ser cero", vbCritical, gstrNombreEmpresa
        txtMontoMovimientoOrigen.SetFocus
        Exit Function
    End If

'    If strCodTipoMovimiento = Codigo_Movimiento_Retiro Then 'Codigo_Movimiento_Deposito Then
'        If CCur(txtMontoMovimientoOrigen.Value) > CCur(lblSaldoDisponible.Caption) Then
'            MsgBox "Monto del movimiento no puede ser mayor al saldo de la cuenta", vbCritical, gstrNombreEmpresa
'            txtMontoMovimientoOrigen.SetFocus
'            Exit Function
'        End If
'    Else
'        If CCur(txtMontoMovimientoDestino.Value) > CCur(lblSaldoCuentaTransferencia.Caption) Then
'            MsgBox "Monto del movimiento no puede ser mayor al saldo disponible", vbCritical, gstrNombreEmpresa
'            txtMontoMovimientoDestino.SetFocus
'            Exit Function
'        End If
'    End If
    
    
    TodoOk = True

End Function

Private Sub chkModificaTC_Click()
    
'    If chkModificaTC.Value = vbChecked Then
'        txtTipoCambio.Enabled = True
'    Else
'        txtTipoCambio.Enabled = False
'    End If
        
End Sub

Private Sub chkModificaTC2_Click()
    If chkModificaTC2.Value = vbChecked Then
        txtTipoCambio.Enabled = True
    Else
        txtTipoCambio.Enabled = False
    End If
End Sub

Private Sub cmdCancelarReg_Click()

    Call Cancelar

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
    
        intContador = tdgConsulta.SelBookmarks.Count - 1
               
        strFechaGrabar = Convertyyyymmdd(lblFechaMovimiento.Caption) & Space(1) & Format(Time, "hh:mm")
                   
        Call ConfiguraRecordsetAuxiliar
            
        'Set adoRegistro = New ADODB.Recordset
        
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
                
            .CommandText = "{ call up_TEProcTransferenciaBancaria('" & _
                strCodFondo & "','" & gstrCodAdministradora & "','" & _
                strFechaGrabar & "','" & _
                strTesoreriaOperacionXML & "') }"
            
            .Execute .CommandText
                                             
        End With
        
        Me.MousePointer = vbDefault
        
        MsgBox Mensaje_Proceso_Exitoso, vbExclamation
        
        frmMainMdi.stbMdi.Panels(3).Text = "Acción"
        
        'cmdOpcion.Visible = True
        With tabCuenta
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

Private Sub cmdRegistrar_Click()
    
    Call Grabar
    
End Sub
Private Sub Grabar()

    Dim curMontoMovimiento              As Double, intRegistro              As Integer
    Dim strDescripMovimiento            As String, strNumOrdenCobroPago     As String
    Dim strTipoMovimiento               As String, strIndMovimiento         As String
    Dim strTipoOperacion                As String, strNumDocumento          As String
    Dim strAccion                       As String, strNumOperacion          As String
    Dim strFechaGrabar                  As String, strFechaObligacion       As String
    Dim strCodParticipe                 As String, strDescripObservacion    As String
        
    'Manejo de Transacciones
    Dim adoError                        As ADODB.Error
    Dim strErrMsg                       As String
    Dim intAccion                       As Integer
    Dim lngNumError                     As Long
    
    Dim strMsgError                     As String
    
    Dim objTipoCambioReemplazoXML       As DOMDocument60
    Dim strTipoCambioReemplazoXML       As String
    Dim strGlosaCuentaOrigen            As String
    Dim strGlosaCuentaDestino           As String
    Dim strCodOrigenCuenta              As String
    
    On Error GoTo ErrorHandler
    
    If Not TodoOk() Then Exit Sub
    
    Me.MousePointer = vbHourglass

    strTipoOperacion = "11"
    
    If strCodTipoMovimiento = Codigo_Movimiento_Deposito Then
        strDescripMovimiento = "Depósito en Cuenta " & strCodFile & "-" & Mid(arrCuentaDestino(cboCuentaDestino.ListIndex), 22, 30) 'strCodAnalitica   'Mid(arrNumCuenta(cboNumCuenta.ListIndex), 4, 8) 'lblAnalitica.Caption
        curMontoMovimiento = CCur(txtMontoMovimientoOrigen.Value)
        'strIndDebeHaber = "D"
        strIndMovimiento = "E"
    Else
        strDescripMovimiento = "Retiro de Cuenta " & strCodFile & "-" & Mid(arrCuentaOrigen(cboCuentaOrigen.ListIndex), 22, 30) '& strCodAnalitica   'Mid(arrNumCuenta(cboNumCuenta.ListIndex), 4, 8) 'lblAnalitica.Caption
        curMontoMovimiento = CCur(txtMontoMovimientoOrigen.Value) * -1
        'strIndDebeHaber = "H"
        strIndMovimiento = "S"
    End If
    

    strNumDocumento = Trim(txtNumDocumento.Text)
    
    strCodParticipe = ""
    
    strDescripObservacion = "" 'Trim(txtDescripObservaciones.Text)
    
    If strEstado = Reg_Adicion Then
        strAccion = "I"
    End If
    
    If strEstado = Reg_Edicion Then
        strAccion = "U"
    End If
    
    strFechaGrabar = Convertyyyymmdd(lblFechaMovimiento.Caption) & Space(1) & Format(Time, "hh:mm")
    strFechaObligacion = Convertyyyymmdd(lblFechaMovimiento.Caption)
    
    strNumOperacion = lblNumOperacion.Caption

            
    If ChkMonDiferente.Value Then
        strDescripMovimiento = "Transferencia con Operación de cambio"
    Else
        strDescripMovimiento = "Transferencia entre cuentas"
    End If
    
    
    ''------------------------------
    ''---Registro en BD-------------
            
    With adoComm
        

                .CommandText = "{ call up_TEManTesoreriaOperacion ('" & _
                             strCodFondo & "','" & gstrCodAdministradora & "','" & strNumOperacion & "'," & _
                             "'000', '00000000', '000'," & _
                             "'000','" & strTipoOperacion & "','','" & strCodTipoMovimiento & "','" & strCodParticipe & "','" & strTipoContraparte & "','" & strCodContraparte & "'," & _
                             "'" & strDescripMovimiento & "','" & strFechaGrabar & "','" & strFechaObligacion & "'," & _
                             "'19000101','" & strCodMoneda & "'," & CCur(txtMontoMovimientoOrigen.Value) & "," & _
                             "'" & strCodMonedaCuenta & "'," & CCur(txtMontoMovimientoDestino.Value) & "," & _
                             "'" & gstrCodClaseTipoCambioOperacionFondo & "','" & gstrValorTipoCambioOperacion & "'," & CDbl(txtTipoCambio.Text) & ",0,0,''," & _
                             "'', 0, '13','" & strNumDocumento & "','" & strDescripObservacion & "'," & _
                             "'" & strCodCuentaActivo & "','" & strCodFile & "','" & strCodAnalitica & "'," & _
                             "'" & strCodCuentaFondo & "','" & strCodFileCuenta & "','" & strCodAnaliticaCuenta & "'," & _
                             "'', '', ''," & _
                             "'', '', ''," & _
                             "'', '', '', '', ''," & _
                             "'01','" & strAccion & "') }"
    
        .Execute
                
    
    End With
        

    Me.MousePointer = vbDefault

    MsgBox Mensaje_Adicion_Exitosa, vbExclamation

    frmMainMdi.stbMdi.Panels(3).Text = "Acción"
   
    strEstado = Reg_Consulta
    cmdOpcion.Visible = True
    
    With tabCuenta
        .TabEnabled(0) = True
        .Tab = 0
    End With
    
    Call Buscar
    
    Exit Sub
    
ErrorHandler:
    
    If err.Number <> 0 Then
        MsgBox err.Number & " " & err.Description, vbCritical + vbOKOnly, Me.Caption
        Me.MousePointer = vbDefault
    End If


End Sub

Private Sub cmdReservar_Click()

    Dim strFechaGrabar  As String
    Dim strNumOperacion  As String 'se usa
    Dim motivo          As String
    Dim str_msg, str_pwd         As String
    Dim adoRegistro     As ADODB.Recordset

    If strEstado <> Reg_Edicion Then
        If strEstado <> Reg_Consulta Then Exit Sub
    End If
    
    If Not TodoOkReversar() Then Exit Sub

    If MsgBox("Desea reversar la Transferencia Bancaria. " & tdgConsulta.Columns("NumOperacion").Value & " ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
        
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
Private Function TodoOkReversar() As Boolean

    TodoOkReversar = False
        
    If tdgConsulta.SelBookmarks.Count = 0 Or tdgConsulta.SelBookmarks.Count > 1 Then
        MsgBox "Debe seleccionar un registro para reversar", vbCritical, gstrNombreEmpresa
        Exit Function
    End If
        
    If cboEstado.ListIndex > -1 Then
        If strCodEstado <> "02" Then
            MsgBox "Solo se puede reversar operaciones ya procesadas", vbOKOnly + vbCritical, Me.Caption
            Exit Function
        End If
    Else
        MsgBox "Debe seleccionar algun estado de operacion", vbCritical, gstrNombreEmpresa
        If cboEstado.Enabled Then cboEstado.SetFocus
        Exit Function
    End If
        
    '*** Si todo paso OK ***
    TodoOkReversar = True

End Function
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

Private Sub CargarListas()
    
    '*** Fondos ***
    strSQL = "{ call up_ACSelDatosParametro(74,'" & gstrCodAdministradora & "','" & gstrCodFondoContable & "') }"
    CargarControlLista strSQL, cboFondo, arrFondo(), ""
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
    
    '*** Estado ***
    strSQL = "SELECT CodParametro CODIGO, DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='ESTRAN' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboEstado, arrEstado(), Sel_Defecto
    If cboEstado.ListCount > 0 Then cboEstado.ListIndex = 0
        
    '*** Banco origen***
    'strSQL = "{ call up_ACSelDatos(22) }"
    strSQL = "SELECT CodPersona CODIGO,RazonSocial DESCRIP FROM InstitucionPersona I WHERE TipoPersona='02' AND IndBanco='X' AND IndVigente='X' "
    strSQL = strSQL + " AND EXISTS (SELECT * FROM BancoCuenta WHERE CodBanco = I.CodPersona) ORDER BY DescripPersona "
    CargarControlLista strSQL, cboBancoOrigen, arrBancoOrigen(), ""
    If cboBancoOrigen.ListCount > 0 Then cboBancoOrigen.ListIndex = 0

    '*** Banco destino***
    'strSQL = "{ call up_ACSelDatos(22) }"
    strSQL = "SELECT CodPersona CODIGO,RazonSocial DESCRIP FROM InstitucionPersona I WHERE TipoPersona='02' AND IndBanco='X' AND IndVigente='X' "
    strSQL = strSQL + " AND exists(SELECT * FROM BancoCuenta WHERE CodBanco = I.CodPersona) ORDER BY DescripPersona "
    CargarControlLista strSQL, cboBancoDestino, arrBancoDestino(), ""
    If cboBancoDestino.ListCount > 0 Then cboBancoDestino.ListIndex = 0

    '*** Formas de pago para caso de transferencias ***
    strSQL = "SELECT CodFormaPago CODIGO, DescripFormaPago DESCRIP FROM FormaPago ORDER BY DescripFormaPago"
    'strSQL = "SELECT CodFormaPago CODIGO, DescripFormaPago DESCRIP FROM FormaPago WHERE IndBCR = 'X' OR IndExterior ='X' OR IndTransferSimple = 'X' ORDER BY DescripFormaPago"
    ' "SELECT CodParametro CODIGO, DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='MEDPAG' AND CodParametro IN ('" & Forma_Pago_TransferBCR & "','" & Forma_Pago_TransferExterior & "','" & Forma_Pago_TransferSimple & "' ) ORDER BY DescripParametro"
    CargarControlLista strSQL, cboFormaPago, arrFormaPago(), Sel_Defecto
    If cboFormaPago.ListCount > 0 Then cboFormaPago.ListIndex = 0


    '*** Cuenta Activo ***
    'strSQL = "SELECT CodCuenta CODIGO, DescripCuenta DESCRIP FROM PlanContable WHERE IndMovimiento='X' AND (CodCuenta LIKE '101%' OR CodCuenta LIKE '104%' OR CodCuenta LIKE '181%') AND IndAuxiliar='X' AND CodAdministradora='" & gstrCodAdministradora & "' ORDER BY CodCuenta"
    'CargarControlLista strSQL, cboCuentaActivo, arrCuentaActivo(), Sel_Defecto
    
    '*** Tipo de Cálculo de Intereses ***
    'strSQL = "SELECT CodParametro CODIGO, DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='CALINT' ORDER BY DescripParametro"
    'CargarControlLista strSQL, cboCalculo, arrCalculo(), ""
    
    'If cboCalculo.ListCount > 0 Then cboCalculo.ListIndex = 0
    
    '*** Tipo Remunerada ***
    'strSQL = "SELECT CodParametro CODIGO, DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPREM' ORDER BY DescripParametro"
    'CargarControlLista strSQL, cboTipoRemunerada, arrTipoRemunerada(), Sel_Defecto
        
    '*** Por Monto ***
    'strSQL = "SELECT CodParametro CODIGO, DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='REMMON' ORDER BY DescripParametro"
   'CargarControlLista strSQL, cboPorMonto, arrPorMonto(), Sel_Defecto
    
    Dim adoRegistro As ADODB.Recordset
    
    Set adoRegistro = New ADODB.Recordset
    
    '*** Tipo Movimiento ***
    strSQL = "SELECT CodParametro CODIGO, DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='MOVCTA' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboTipoMovimiento, arrTipoMovimiento(), Sel_Defecto
   
    '*** Obtener dias de desplazamiento
'    adoComm.CommandText = "SELECT CONVERT(int,ValorParametro) AS DiasDesplaza FROM ParametroGeneral WHERE CodParametro = '19'"
'    Set adoRegistro = adoComm.Execute
'
'    If Not (adoRegistro.EOF) Then
'        intDiasDesplazamiento = adoRegistro("DiasDesplaza")
'    End If
'
'    If intDiasDesplazamiento = Null Then
'        intDiasDesplazamiento = 0
'    End If
    
    Set adoRegistro = Nothing
   
        
End Sub

Private Sub InicializarValores()

    '*** Valores Iniciales ***
    strEstado = Reg_Defecto
    tabCuenta.Tab = 0
     
    tabCuenta.TabEnabled(1) = False
    '*** Valores a las fechas de búsqueda
    dtpFechaMovimBCDesde.Value = gdatFechaActual
    dtpFechaMovimBCHasta.Value = dtpFechaMovimBCDesde.Value
    
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgConsulta.Columns(0).Width = tdgConsulta.Width * 0.01 * 10
    tdgConsulta.Columns(1).Width = tdgConsulta.Width * 0.01 * 10
    tdgConsulta.Columns(2).Width = tdgConsulta.Width * 0.01 * 30
    tdgConsulta.Columns(3).Width = tdgConsulta.Width * 0.01 * 5
    tdgConsulta.Columns(4).Width = tdgConsulta.Width * 0.01 * 10
    tdgConsulta.Columns(5).Width = tdgConsulta.Width * 0.01 * 5
    tdgConsulta.Columns(6).Width = tdgConsulta.Width * 0.01 * 10
    tdgConsulta.Columns(7).Width = tdgConsulta.Width * 0.01 * 10
    tdgConsulta.Columns(8).Width = tdgConsulta.Width * 0.01 * 20
    tdgConsulta.Columns(10).Width = tdgConsulta.Width * 0.01 * 15
    tdgConsulta.Columns(11).Width = tdgConsulta.Width * 0.01 * 20
    tdgConsulta.Columns(12).Width = tdgConsulta.Width * 0.01 * 15
    
'    NumOperacion, --ampliar campo --centrar --
'FechaObligacion, --ampliar campo --centrar
'DescripOperacion, --ampliar campo
'DescripMonedaOrigen, --ampliar campo --centrar
'MontoOperacionOrigen, --ampliar campo --derecha --numero
'DescripMonedaDestino, -ampliar campo --centrar
'MontoOperacionDestino, --ampliar campo --derecha --numero
'NumDocumento , --reducir - -derercha
'DescripBancoOrigen , ok
'TipoCuentaOrigen,
'NumCuentaOrigen, ..reducir
'DescripBancoDestino
'
'TipoCuentaDestino,
'NumCuentaDestino ..reducir


    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    
End Sub

Private Sub DarFormato()

    Dim c As Object
    Dim elemento As Object
    
    For Each c In Me.Controls
        If TypeOf c Is Label Then
            Call FormatoEtiqueta(c, vbLeftJustify)
        End If
    Next
            
    For Each elemento In Me.Controls
    
        If TypeOf elemento Is TDBGrid Then
            Call FormatoGrilla(elemento)
        End If
    
    Next
            
    lblDescrip(1).Alignment = vbCenter
    lblDescrip(2).Alignment = vbCenter
    lblSaldoDisponible.Alignment = vbRightJustify
    lblSaldoCuentaTransferencia.Alignment = vbRightJustify
            
            
End Sub
Private Sub Form_Unload(Cancel As Integer)

    Call OcultarReportes
    Set frmCuentaFondo = Nothing
    
End Sub

Public Sub Buscar()

    Dim strFechaMovimBCDesde          As String
    Dim strFechaMovimBCHasta          As String
    Dim datFechaSiguiente             As Date
    Dim strSQL As String

    Me.MousePointer = vbHourglass

    If Not IsNull(dtpFechaMovimBCDesde.Value) Or Not IsNull(dtpFechaMovimBCHasta.Value) Then
        strFechaMovimBCDesde = Convertyyyymmdd(dtpFechaMovimBCDesde.Value)
        datFechaSiguiente = DateAdd("d", 1, dtpFechaMovimBCHasta.Value)
        strFechaMovimBCHasta = Convertyyyymmdd(datFechaSiguiente)
    End If



            strSQL = "{ call up_TEListarTransferenciaBancaria ('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                    strFechaMovimBCDesde & "','" & strFechaMovimBCHasta & "','" & strCodEstado & "') }"

            Set adoConsulta = New ADODB.Recordset

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

            Call AutoAjustarGrillas


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
            
            If cboEstado.ListIndex = 0 Then
        
                MsgBox "Seleccione un Estado.", vbCritical, gstrNombreEmpresa
                cboEstado.SetFocus
            Else
                Call Buscar
        End If


        Case vExit
            Call Salir
        
    End Select
    
End Sub

Public Sub Adicionar()

    frmMainMdi.stbMdi.Panels(3).Text = "Adicionar movimiento entre cuentas del fondo..."
    
    strEstado = Reg_Adicion
    LlenarFormulario strEstado
    cmdOpcion.Visible = False
    cmdRegistrar.Enabled = True
    'Habilitar la edición de campos
    'frmOrigen.Enabled = True
    'frmDestino.Enabled = True
    'frmTipoMovimiento.Enabled = True
    txtMontoMovimientoOrigen.Enabled = True
    txtMontoMovimientoDestino.Enabled = True
    txtTipoCambio.Enabled = True
    'frmFormaPago.Enabled = True
    '-----
    With tabCuenta
        .TabEnabled(0) = False
        .TabEnabled(1) = True
        .Tab = 1
        .Caption = "Registro de Movimiento"
    End With
        'Call Deshabilita
'    End If
    
End Sub



Private Sub lblDescripTCArbitraje_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button And vbRightButton Then
    ' User right-clicked the list box.
'    PopupMenu mnuListPopup
  
End If

End Sub

Private Sub lblSaldoCuentaTransferencia_Change()

    Call FormatoMillarEtiqueta(lblSaldoCuentaTransferencia, Decimales_Monto)
    
End Sub

Private Sub lblSaldoDisponible_Change()

    Call FormatoMillarEtiqueta(lblSaldoDisponible, Decimales_Monto)
    
End Sub

'Private Sub mnuSListPopup_Click(Index As Integer)
'
'    lblDescripTCArbitraje.Caption = mnuSListPopup(Index).Caption
'
'    If Index = 0 Then
'        strIndSentidoTipoCambio = "M"
'        dblTipoCambioOpera = CDbl(txtTipoCambio.Value)
'        txtTipoCambio.Text = 1 / CDbl(txtTipoCambio.Value)    'CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, DateAdd("d", -1, gdatFechaActual), strCodMoneda, strCodMonedaCuenta))
'        mnuSListPopup.Item(Index).Enabled = False
'        mnuSListPopup.Item(2).Enabled = True
'    End If
'
'    If Index = 2 Then
'        strIndSentidoTipoCambio = "D"
'        txtTipoCambio.Text = 1 / CDbl(txtTipoCambio.Value) 'CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, DateAdd("d", -1, gdatFechaActual), strCodMonedaCuenta, strCodMoneda))
'        mnuSListPopup.Item(Index).Enabled = False
'        mnuSListPopup.Item(0).Enabled = True
'        dblTipoCambioOpera = CDbl(txtTipoCambio.Value)
'    End If
'
'
'End Sub

Private Sub tabCuenta_Click(PreviousTab As Integer)

    Select Case tabCuenta.Tab
        Case 1, 2
            If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vQuery)
            If strEstado = Reg_Defecto Then tabCuenta.Tab = 0
                                
    End Select
    
End Sub


Private Sub tdgConsulta_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)

    If ColIndex = 4 Or ColIndex = 6 Then
        Call DarFormatoValor(Value, Decimales_Monto)
    End If


End Sub

Private Sub tdgConsulta_HeadClick(ByVal ColIndex As Integer)
    Static numColindex As Integer

    tdgConsulta.Splits(0).Columns(numColindex).HeadingStyle.ForegroundPicture = Null

    Call OrdenarDBGrid(ColIndex, adoConsulta, tdgConsulta)
    
    numColindex = ColIndex
End Sub

Private Sub tdgConsulta_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

'    On Error GoTo Error1            '/**/ HMC Habilitamos la rutina de Errores.
'
'    If Not IsNull(LastRow) Then     '/**/
'        If tdgConsulta.Row < 0 Then Exit Sub
'
'        If CStr(LastRow) = Valor_Caracter Then Exit Sub
'
'        tdgConsulta.ToolTipText = Valor_Caracter
'        If LastRow >= 1 Then
'            Call SaldoCuentaDineraria(tdgConsulta.Columns("CodCuentaActivo").Value, tdgConsulta.Columns("CodFile").Value, tdgConsulta.Columns("CodAnalitica").Value, tdgConsulta.Columns("CodMoneda").Value)
'        Else
'            tdgConsulta.ToolTipText = Valor_Caracter
'        End If
'
'    End If                           '/**/
'    Exit Sub
'
'Error1:
'    MsgBox DescripcionError & vbNewLine & DescripcionTecnica & err.Description, vbExclamation, TituloError ' Mostrar Error
    
End Sub


Private Sub txtMontoMovimientoDestino_KeyPress(KeyAscii As Integer)

'    Call ValidaCajaTexto(KeyAscii, "M", txtMontoMovimientoDestino, Decimales_Monto)
    
    blnModifica = True
    
    If KeyAscii = vbKeyReturn Then
        Call CalculoTotal(1)
        blnModifica = False
    End If
    
End Sub


Private Sub txtMontoMovimientoDestino_LostFocus()

    If blnModifica Then
        Call CalculoTotal(1)
        blnModifica = False
    End If
    

End Sub

Private Sub txtMontoMovimientoOrigen_KeyPress(KeyAscii As Integer)
    
    blnModifica = True
    
    If KeyAscii = vbKeyReturn Then
        Call CalculoTotal(0)
        blnModifica = False
    End If

End Sub



Private Sub txtMontoMovimientoOrigen_LostFocus()


    If blnModifica Then
        Call CalculoTotal(0)
        blnModifica = False
    End If


End Sub

Private Sub txtTipoCambio_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        Call CalculoTotal(0)
    End If
    
    
End Sub

Private Sub CalculoTotal(index As Integer)

    Dim curMonTotal As Currency
    Dim dblValorTC As Double
    
    If index = 0 Then ' actualiza desde el origen al destino
        strCodMonedaParEvaluacion = strCodMoneda & strCodMonedaCuenta
        If chkModificaTC.Value = vbUnchecked Then
            dblValorTC = 0
            If CDbl(txtMontoMovimientoDestino.Value) <> 0 Then
                If strCodMoneda <> strCodMonedaCuenta Then
                    dblValorTC = CDbl(txtMontoMovimientoOrigen.Value) / CDbl(txtMontoMovimientoDestino.Value)
                Else
                    dblValorTC = 1 'CDbl(txtMontoMovimientoOrigen.Value) * CDbl(txtMontoMovimientoDestino.Value)
                    curMonTotal = txtMontoMovimientoOrigen.Value
                    'NUEVO**
                    txtMontoMovimientoDestino.Text = curMonTotal
                End If
                If strCodMonedaParEvaluacion <> strCodMonedaParPorDefecto Then
                    dblValorTC = 1 / dblValorTC
                End If
            End If
            txtTipoCambio.Text = dblValorTC   'CStr(dblValorTC)
            
        Else
            If CDbl(txtTipoCambio.Value) <> 0 Then
                curMonTotal = Round(ObtenerMontoArbitraje(CDbl(txtMontoMovimientoOrigen.Value), CDbl(txtTipoCambio.Value), strCodMonedaParEvaluacion, strCodMonedaParPorDefecto), 2)
            Else
                curMonTotal = 0
            End If
            txtMontoMovimientoDestino.Text = curMonTotal  'CStr(curMonTotal)
        End If
    End If
    
    If index = 1 Then ' actualiza desde el destino al origen
        strCodMonedaParEvaluacion = strCodMonedaCuenta & strCodMoneda
        If chkModificaTC.Value = vbUnchecked Then
            dblValorTC = 0
            If CDbl(txtMontoMovimientoDestino.Value) <> 0 Then
                If strCodMoneda <> strCodMonedaCuenta Then
                    dblValorTC = CDbl(txtMontoMovimientoOrigen.Value) / CDbl(txtMontoMovimientoDestino.Value)
                Else
                    dblValorTC = 1 'CDbl(txtMontoMovimientoOrigen.Value) * CDbl(txtMontoMovimientoDestino.Value)
                End If
                If strCodMonedaParEvaluacion <> strCodMonedaParPorDefecto Then
                    dblValorTC = 1 / dblValorTC
                End If
            End If
            txtTipoCambio.Text = dblValorTC   'CStr(dblValorTC)
        Else
            If CDbl(txtTipoCambio.Value) <> 0 Then
                curMonTotal = Round(ObtenerMontoArbitraje(CDbl(txtMontoMovimientoDestino.Value), CDbl(txtTipoCambio.Value), strCodMonedaParEvaluacion, strCodMonedaParPorDefecto), 2)
            Else
                curMonTotal = 0
            End If
            txtMontoMovimientoOrigen.Text = curMonTotal   'CStr(curMonTotal)
        End If
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

Private Function TodoOkProceso() As Boolean
        
    TodoOkProceso = False
        
    If tdgConsulta.SelBookmarks.Count - 1 = -1 Then
        MsgBox "Debe seleccionar registros para procesar!", vbCritical, gstrNombreEmpresa
        Exit Function
    End If
    
    If cboEstado.ListIndex > 0 Then
        If tdgConsulta.Columns("EstadoOperacion") <> Estado_Acuerdo_Ingresado Then
            MsgBox "No se pueden procesar las operaciones seleccionadas!", vbCritical, gstrNombreEmpresa
            Exit Function
        End If
    End If
        
    '*** Si todo paso OK ***
    TodoOkProceso = True
  
End Function

Private Sub txtTipoCambio_LostFocus()
    Call txtTipoCambio_KeyPress(vbKeyReturn)
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

