VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{830D5A06-9B70-4F8C-98B6-7A19C4E7760B}#1.0#0"; "TAMControls.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmCuentaFondo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Definición de Cuentas del Fondo"
   ClientHeight    =   8985
   ClientLeft      =   240
   ClientTop       =   510
   ClientWidth     =   13605
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8985
   ScaleWidth      =   13605
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   11520
      TabIndex        =   4
      Top             =   8160
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
      Left            =   960
      TabIndex        =   3
      Top             =   8160
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
      Caption2        =   "&Eliminar"
      Tag2            =   "4"
      Visible2        =   0   'False
      ToolTipText2    =   "Eliminar"
      UserControlWidth=   4200
   End
   Begin TabDlg.SSTab tabCuenta 
      Height          =   7935
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   13545
      _ExtentX        =   23892
      _ExtentY        =   13996
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
      TabPicture(0)   =   "frmCuentaFondo.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDescrip(10)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cboFondo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "tdgConsulta"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "frmCuentaFondo.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdAccion"
      Tab(1).Control(1)=   "fraCuenta(0)"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Movimiento entre Cuentas"
      TabPicture(2)   =   "frmCuentaFondo.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraCuenta(1)"
      Tab(2).ControlCount=   1
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   -64920
         TabIndex        =   28
         Top             =   7080
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
         Bindings        =   "frmCuentaFondo.frx":0054
         Height          =   5925
         Left            =   360
         OleObjectBlob   =   "frmCuentaFondo.frx":006E
         TabIndex        =   57
         Top             =   1200
         Width           =   12555
      End
      Begin VB.Frame fraCuenta 
         Caption         =   "Definición"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6375
         Index           =   0
         Left            =   -74640
         TabIndex        =   5
         Top             =   540
         Width           =   12750
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
            Left            =   2220
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   5280
            Width           =   3705
         End
         Begin VB.TextBox txtMontoRemunerada 
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
            Left            =   9390
            TabIndex        =   26
            Top             =   5280
            Width           =   2865
         End
         Begin VB.ComboBox cboPorMonto 
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
            Left            =   9390
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   4800
            Width           =   2865
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   9390
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   4320
            Width           =   2865
         End
         Begin VB.ComboBox cboTipoRemunerada 
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
            Left            =   9390
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   3840
            Width           =   2865
         End
         Begin VB.CheckBox chkRemunerada 
            Caption         =   "Remunerada"
            Height          =   255
            Left            =   8130
            TabIndex        =   22
            Top             =   3360
            Width           =   1695
         End
         Begin VB.TextBox txtNumCuenta 
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
            Left            =   2220
            MaxLength       =   30
            TabIndex        =   17
            Top             =   4800
            Width           =   5505
         End
         Begin VB.TextBox txtResponsable 
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
            Left            =   2220
            MaxLength       =   40
            TabIndex        =   16
            Top             =   4320
            Width           =   5505
         End
         Begin VB.TextBox txtDescripCuenta 
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
            Left            =   2220
            MaxLength       =   80
            TabIndex        =   15
            Top             =   3840
            Width           =   5505
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
            Left            =   2220
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   960
            Width           =   5505
         End
         Begin VB.ComboBox cboTipoCuenta 
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
            Left            =   2220
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   480
            Width           =   5505
         End
         Begin VB.ComboBox cboCuentaActivo 
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
            Left            =   2220
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   1995
            Width           =   5505
         End
         Begin VB.ComboBox cboBanco 
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
            Left            =   2220
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   1500
            Width           =   5505
         End
         Begin MSComCtl2.DTPicker dtpFechaApertura 
            Height          =   315
            Left            =   10155
            TabIndex        =   20
            Top             =   960
            Width           =   2025
            _ExtentX        =   3572
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
         Begin MSComCtl2.DTPicker dtpFechaInicioCalculo 
            Height          =   315
            Left            =   10155
            TabIndex        =   21
            Top             =   1500
            Width           =   2025
            _ExtentX        =   3572
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
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Cálculo"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   20
            Left            =   360
            TabIndex        =   53
            Top             =   5295
            Width           =   645
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Inicio Cálculo"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   19
            Left            =   8130
            TabIndex        =   52
            Top             =   1515
            Width           =   1755
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tasa"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   18
            Left            =   8130
            TabIndex        =   51
            Top             =   4335
            Width           =   435
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   21
            Left            =   8130
            TabIndex        =   34
            Top             =   5295
            Width           =   540
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Por Monto"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   8
            Left            =   8130
            TabIndex        =   33
            Top             =   4815
            Width           =   885
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   7
            Left            =   8130
            TabIndex        =   32
            Top             =   3855
            Width           =   390
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Analítica"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   8100
            TabIndex        =   31
            Top             =   495
            Width           =   780
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Apertura"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   8100
            TabIndex        =   30
            Top             =   975
            Width           =   1590
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Descripción"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   4
            Left            =   360
            TabIndex        =   29
            Top             =   3840
            Width           =   1020
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Responsable"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   5
            Left            =   360
            TabIndex        =   27
            Top             =   4350
            Width           =   1110
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Número"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   6
            Left            =   360
            TabIndex        =   19
            Top             =   4815
            Width           =   660
         End
         Begin VB.Label lblAnalitica 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   10155
            TabIndex        =   13
            Top             =   480
            Width           =   2025
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   17
            Left            =   360
            TabIndex        =   9
            Top             =   975
            Width           =   690
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Banco"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   8
            Top             =   1515
            Width           =   555
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta Activo"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   7
            Top             =   2025
            Width           =   1215
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   16
            Left            =   360
            TabIndex        =   6
            Top             =   495
            Width           =   390
         End
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
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   660
         Width           =   6390
      End
      Begin VB.Frame fraCuenta 
         Caption         =   "Registro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5565
         Index           =   1
         Left            =   -74640
         TabIndex        =   35
         Top             =   600
         Width           =   12240
         Begin VB.TextBox txtDescripcionTransf 
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
            Left            =   1920
            TabIndex        =   44
            Top             =   2340
            Width           =   8115
         End
         Begin VB.TextBox txtNroTransaccion 
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
            Left            =   1920
            TabIndex        =   43
            Top             =   1950
            Width           =   1755
         End
         Begin VB.CheckBox chkModificaTC 
            Caption         =   "FijarT/C"
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
            Left            =   7770
            TabIndex        =   47
            Top             =   3210
            Value           =   1  'Checked
            Width           =   1575
         End
         Begin VB.Frame fraFinal 
            Caption         =   "Destino"
            ForeColor       =   &H00000080&
            Height          =   555
            Left            =   450
            TabIndex        =   65
            Top             =   5910
            Width           =   5115
            Begin VB.TextBox txtTipoCambioDestinoContable 
               Alignment       =   1  'Right Justify
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
               Height          =   285
               Left            =   1710
               TabIndex        =   74
               Top             =   900
               Width           =   2025
            End
            Begin VB.TextBox txtMontoContableDestino 
               Alignment       =   1  'Right Justify
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
               Height          =   285
               Left            =   1710
               TabIndex        =   66
               Top             =   1320
               Width           =   2025
            End
            Begin VB.Label lblMonedaDestinoContable 
               AutoSize        =   -1  'True
               Caption         =   "(USD/PEN)"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   3900
               TabIndex        =   72
               Top             =   960
               Width           =   990
               WordWrap        =   -1  'True
            End
            Begin VB.Label lblMonedaContable 
               AutoSize        =   -1  'True
               Caption         =   "(PEN)"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   1
               Left            =   4320
               TabIndex        =   71
               Top             =   1350
               Width           =   510
               WordWrap        =   -1  'True
            End
            Begin VB.Label lblDescripTCDestinoContable 
               AutoSize        =   -1  'True
               Caption         =   "T/C Contable"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   180
               TabIndex        =   68
               Top             =   960
               Width           =   2265
               WordWrap        =   -1  'True
            End
            Begin VB.Label lblDescripDestinoContable 
               AutoSize        =   -1  'True
               Caption         =   "Monto Contable"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   27
               Left            =   210
               TabIndex        =   67
               Top             =   1380
               Width           =   1350
            End
         End
         Begin VB.Frame fraInicio 
            Caption         =   "Origen"
            ForeColor       =   &H00000080&
            Height          =   645
            Left            =   570
            TabIndex        =   61
            Top             =   5790
            Width           =   5085
            Begin VB.TextBox txtTipoCambioOrigenContable 
               Alignment       =   1  'Right Justify
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
               Height          =   285
               Left            =   1710
               TabIndex        =   73
               Top             =   900
               Width           =   2025
            End
            Begin VB.TextBox txtMontoContableOrigen 
               Alignment       =   1  'Right Justify
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
               Height          =   285
               Left            =   1710
               TabIndex        =   63
               Top             =   1320
               Width           =   2025
            End
            Begin VB.Label lblMonedaOrigenContable 
               AutoSize        =   -1  'True
               Caption         =   "(USD/PEN)"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   3960
               TabIndex        =   70
               Top             =   960
               Width           =   1020
               WordWrap        =   -1  'True
            End
            Begin VB.Label lblMonedaContable 
               AutoSize        =   -1  'True
               Caption         =   "(PEN)"
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   0
               Left            =   4410
               TabIndex        =   69
               Top             =   1380
               Width           =   510
               WordWrap        =   -1  'True
            End
            Begin VB.Label lblDescripOrigenContable 
               AutoSize        =   -1  'True
               Caption         =   "Monto Contable"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   240
               TabIndex        =   64
               Top             =   1350
               Width           =   1350
            End
            Begin VB.Label lblDescripTCOrigenContable 
               AutoSize        =   -1  'True
               Caption         =   "T/C Contable "
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   210
               TabIndex        =   62
               Top             =   930
               Width           =   1935
               WordWrap        =   -1  'True
            End
         End
         Begin VB.CheckBox ChkMonDiferente 
            Caption         =   "Transacción en distinta moneda"
            Enabled         =   0   'False
            Height          =   255
            Left            =   360
            TabIndex        =   48
            Top             =   4170
            Width           =   3255
         End
         Begin VB.ComboBox cboCuentaFondo 
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
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   45
            Top             =   2760
            Width           =   8145
         End
         Begin VB.ComboBox cboTipoMovimiento 
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
            Top             =   1290
            Width           =   2145
         End
         Begin VB.CommandButton cmdRegistrar 
            Caption         =   "&Registrar"
            Height          =   735
            Left            =   8700
            Picture         =   "frmCuentaFondo.frx":58CA
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "Registrar el Movimiento"
            Top             =   4680
            Width           =   1200
         End
         Begin TAMControls.TAMTextBox txtMontoMovimientoOrigen 
            Height          =   315
            Left            =   7740
            TabIndex        =   42
            Tag             =   "0"
            Top             =   1290
            Width           =   2175
            _ExtentX        =   3836
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
            Container       =   "frmCuentaFondo.frx":5B85
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
         Begin TAMControls.TAMTextBox txtMontoMovimientoDestino 
            Height          =   315
            Left            =   7770
            TabIndex        =   49
            Top             =   4110
            Width           =   2175
            _ExtentX        =   3836
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
            Container       =   "frmCuentaFondo.frx":5BA1
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
         Begin TAMControls.TAMTextBox txtTipoCambio 
            Height          =   315
            Left            =   4560
            TabIndex        =   46
            Top             =   3180
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
            Container       =   "frmCuentaFondo.frx":5BBD
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
         Begin VB.Label lblDescripcionTransf 
            Caption         =   "Descripción"
            ForeColor       =   &H00800000&
            Height          =   345
            Left            =   300
            TabIndex        =   82
            Top             =   2400
            Width           =   1425
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Nro. Transacción"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   24
            Left            =   300
            TabIndex        =   81
            Top             =   2010
            Width           =   1485
         End
         Begin VB.Label lblDescripDestino 
            AutoSize        =   -1  'True
            Caption         =   "Monto Transferencia Destino"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   5160
            TabIndex        =   80
            Top             =   4140
            Width           =   2475
         End
         Begin VB.Label lblMonedaDestino 
            AutoSize        =   -1  'True
            Caption         =   "(USD)"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   9930
            TabIndex        =   79
            Top             =   4140
            Width           =   510
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblDescripOrigen 
            AutoSize        =   -1  'True
            Caption         =   "Monto Transferencia Origen"
            ForeColor       =   &H00800000&
            Height          =   390
            Left            =   5220
            TabIndex        =   78
            Top             =   1320
            Width           =   2520
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblMonedaOrigen 
            AutoSize        =   -1  'True
            Caption         =   "(USD)"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   9930
            TabIndex        =   77
            Top             =   1320
            Width           =   525
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "XXXXX"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   13
            Left            =   330
            TabIndex        =   76
            Top             =   6180
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "T/C Arbitraje "
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   12
            Left            =   3390
            TabIndex        =   75
            Top             =   3210
            Width           =   1170
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   0
            X1              =   300
            X2              =   10140
            Y1              =   4500
            Y2              =   4500
         End
         Begin VB.Label lblDescripTCArbitraje 
            AutoSize        =   -1  'True
            Caption         =   "(USD/SOL)"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   6600
            TabIndex        =   60
            Top             =   3210
            Width           =   1080
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblDescripCuentaOrigen 
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
            Left            =   1980
            TabIndex        =   59
            Top             =   360
            Width           =   8115
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   15
            Left            =   300
            TabIndex        =   58
            Top             =   360
            Width           =   615
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   1
            X1              =   270
            X2              =   10110
            Y1              =   1800
            Y2              =   1800
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Disponible"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   23
            Left            =   5190
            TabIndex        =   56
            Top             =   3750
            Width           =   2100
         End
         Begin VB.Label lblSaldoCuentaTransferencia 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            Height          =   315
            Left            =   7770
            TabIndex        =   55
            Top             =   3660
            Width           =   2145
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Transferir A ..."
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   22
            Left            =   300
            TabIndex        =   54
            Top             =   2790
            Width           =   1245
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   9
            Left            =   300
            TabIndex        =   40
            Top             =   1335
            Width           =   390
         End
         Begin VB.Label lblFechaMovimiento 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "99/99/9999"
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
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label lblSaldoDisponible 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            Height          =   315
            Left            =   7740
            TabIndex        =   38
            Top             =   855
            Width           =   2175
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   11
            Left            =   300
            TabIndex        =   37
            Top             =   855
            Width           =   540
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Disponible"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   14
            Left            =   5220
            TabIndex        =   36
            Top             =   870
            Width           =   2040
         End
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Fondo"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   10
         Left            =   510
         TabIndex        =   2
         Top             =   690
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmCuentaFondo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrFondo()              As String, arrTipoCuenta()      As String
Dim arrMoneda()             As String, arrBanco()           As String
Dim arrCuentaActivo()       As String, arrTipoRemunerada()  As String
Dim arrTasa()               As String, arrPorMonto()        As String
Dim arrTipoMovimiento()     As String, arrCalculo()         As String
Dim arrCuentaFondo()        As String

Dim strCodFondo             As String, strCodTipoCuenta     As String
Dim strCodMoneda            As String, strCodBanco          As String
Dim strCodCuentaActivo      As String, strCodTipoRemunerada As String
Dim strCodTasa              As String, strCodPorMonto       As String
Dim strCodTipoMovimiento    As String, strCodCalculo        As String
Dim strCodFile              As String, strCodAnalitica      As String
Dim strCodCuentaFondo       As String, strCodFileCuenta     As String
Dim strCodAnaliticaCuenta   As String, strSignoMoneda       As String
Dim strCodMonedaCuenta      As String, strSQL               As String
Dim strSignoMonedaCuenta    As String, strCodBancoCuenta    As String
Dim strCodSignoMonedaCuenta As String
Dim strCodSignoMoneda       As String
Dim strEstado               As String
Dim dblSaldoDisponible      As Double   'HMC
Dim strModalidadCambio      As String
Dim strCodContraparte       As String
Dim strTipoContraparte      As String
Dim strCodMonedaParEvaluacion As String, strCodMonedaParPorDefecto As String
Dim adoRegistroAux          As ADODB.Recordset
Dim adoConsulta             As ADODB.Recordset
Dim indSortAsc              As Boolean, indSortDesc         As Boolean

Private Sub SaldoCuenta()

    Dim adoRegistro As ADODB.Recordset
    Dim datFecha As Date, datFechaSiguiente As Date
    
    Set adoRegistro = New ADODB.Recordset
    
    datFecha = CVDate(lblFechaMovimiento.Caption)
    datFechaSiguiente = DateAdd("d", 1, datFecha)
    
    With adoComm
        If strCodMoneda = Codigo_Moneda_Local Then
            .CommandText = "SELECT (SaldoInicialContable + SaldoParcialContable) SaldoContable, (SaldoInicialMN + MontoDebeMN + MontoHaberMN) SaldoDisponible From PartidaContableSaldos "
        Else
            .CommandText = "SELECT (SaldoInicialContable + SaldoParcialContable) SaldoContable, (SaldoInicialME + MontoDebeME + MontoHaberME) SaldoDisponible From PartidaContableSaldos "
        End If
        .CommandText = .CommandText & "WHERE (FechaSaldo>='" & Convertyyyymmdd(datFecha) & "' AND FechaSaldo<'" & Convertyyyymmdd(datFechaSiguiente) & "') AND "
        .CommandText = .CommandText & "CodCuenta='" & strCodCuentaActivo & "' AND CodAnalitica='" & strCodAnalitica & "' AND "
        .CommandText = .CommandText & "CodFile='" & strCodFile & "' AND CodFondo='" & strCodFondo & "' AND "
        .CommandText = .CommandText & "CodAdministradora='" & gstrCodAdministradora & "'"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            lblSaldoDisponible.Caption = adoRegistro("SaldoDisponible")
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
End Sub

Private Sub SaldoCuentaDineraria(ByVal strpCodCuenta As String, ByVal strpCodFile As String, ByVal strpCodAnalitica As String, ByVal strpCodMoneda As String)

    Dim adoRegistro As ADODB.Recordset
    Dim datFecha As Date, datFechaSiguiente As Date
    
    Set adoRegistro = New ADODB.Recordset
    
    datFecha = CVDate(lblFechaMovimiento.Caption)
    datFechaSiguiente = DateAdd("d", 1, datFecha)
    
    With adoComm
        If strpCodMoneda = Codigo_Moneda_Local Then
            .CommandText = "SELECT (SaldoInicialContable + SaldoParcialContable) SaldoContable, (SaldoInicialMN + MontoDebeMN + MontoHaberMN) SaldoDisponible From PartidaContableSaldos "
        Else
            .CommandText = "SELECT (SaldoInicialContable + SaldoParcialContable) SaldoContable, (SaldoInicialME + MontoDebeME + MontoHaberME) SaldoDisponible From PartidaContableSaldos "
        End If
        .CommandText = .CommandText & "WHERE (FechaSaldo>='" & Convertyyyymmdd(datFecha) & "' AND FechaSaldo<'" & Convertyyyymmdd(datFechaSiguiente) & "') AND "
        .CommandText = .CommandText & "CodCuenta='" & strpCodCuenta & "' AND CodAnalitica='" & strpCodAnalitica & "' AND "
        .CommandText = .CommandText & "CodFile='" & strpCodFile & "' AND CodFondo='" & strCodFondo & "' AND "
        .CommandText = .CommandText & "CodAdministradora='" & gstrCodAdministradora & "'"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            tdgConsulta.ToolTipText = "Saldo Disponible:" & Space(1) & tdgConsulta.Columns(4).Value & Space(1) & FormatNumber(adoRegistro("SaldoDisponible"), 2)
            dblSaldoDisponible = FormatNumber(adoRegistro("SaldoDisponible"), 2)   'HMC - Variable creada para abilitar el Boton Eliminar
        Else
            tdgConsulta.ToolTipText = "Saldo Disponible:" & Space(1) & tdgConsulta.Columns(4).Value & Space(1) & FormatNumber(0, 2)
            dblSaldoDisponible = FormatNumber(0, 2)                                'HMC
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    Me.Refresh
    
End Sub

Private Function TodoOk() As Boolean

       
    TodoOk = False
    
    If cboTipoCuenta.ListIndex = 0 Then
        MsgBox "Seleccione el Tipo de Cuenta", vbCritical, gstrNombreEmpresa
        cboTipoCuenta.SetFocus
        Exit Function
    End If
    
    If cboMoneda.ListIndex = 0 Then
        MsgBox "Seleccione la moneda", vbCritical, gstrNombreEmpresa
        cboMoneda.SetFocus
        Exit Function
    End If
    
    'If cboBanco.ListIndex = 0 Then
     '   MsgBox "Seleccione el Banco", vbCritical, gstrNombreEmpresa
      '  cboBanco.SetFocus
       ' Exit Function
    'End If
    
    If cboCuentaActivo.ListIndex = 0 Then
        MsgBox "Seleccione la cuenta del activo", vbCritical, gstrNombreEmpresa
        cboCuentaActivo.SetFocus
        Exit Function
    End If
        
    If Trim(txtDescripCuenta.Text) = "" Then
        MsgBox "Descripción de cuenta no ingresada", vbCritical, gstrNombreEmpresa
        txtDescripCuenta.SetFocus
        Exit Function
    End If
    
    If Trim(txtResponsable.Text) = "" Then
        MsgBox "Responsable de cuenta no ingresado", vbCritical, gstrNombreEmpresa
        txtResponsable.SetFocus
        Exit Function
    End If
    
    If Trim(txtNumCuenta.Text) = "" Then
        MsgBox "Número de cuenta no ingresado", vbCritical, gstrNombreEmpresa
        txtNumCuenta.SetFocus
        Exit Function
    End If
        
    If chkRemunerada.Value Then
        If cboTipoRemunerada.ListIndex = 0 Then
            MsgBox "Seleccione el tipo de cuenta remunerada", vbCritical, gstrNombreEmpresa
            cboTipoRemunerada.SetFocus
            Exit Function
        End If
        
'        If cboTasa.ListIndex = 0 Then
'            MsgBox "Seleccione la tasa", vbCritical, gstrNombreEmpresa
'            cboTasa.SetFocus
'            Exit Function
'        End If
        
        If cboPorMonto.Enabled Then
            If cboPorMonto.ListIndex = 0 Then
                MsgBox "Seleccione el tipo de monto", vbCritical, gstrNombreEmpresa
                cboPorMonto.SetFocus
                Exit Function
            End If
            
            If CDec(txtMontoRemunerada.Text) = 0 Then
                MsgBox "Monto no ingresado", vbCritical, gstrNombreEmpresa
                txtMontoRemunerada.SetFocus
                Exit Function
            End If
        End If
    Else
        If strCodTipoCuenta = Codigo_Tipo_Cuenta_Ahorro Then
'            If cboTasa.ListIndex = 0 Then
'                MsgBox "Seleccione la tasa", vbCritical, gstrNombreEmpresa
'                cboTasa.SetFocus
'                Exit Function
'            End If
        End If
    End If
        
    '*** Si todo paso OK ***
    TodoOk = True
  
End Function
Public Sub Abrir()

End Sub


Public Sub Adicionar()

    frmMainMdi.stbMdi.Panels(3).Text = "Adicionar cuentas del fondo..."
    
    strEstado = Reg_Adicion
    LlenarFormulario strEstado
    cmdOpcion.Visible = False
    With tabCuenta
        .TabEnabled(0) = False
        .TabEnabled(2) = False
        .Tab = 1
    End With
    Call Deshabilita
    
End Sub

Private Sub Deshabilita()
    
    chkRemunerada.Enabled = False
    cboTipoRemunerada.Enabled = False
    cboTasa.Enabled = False
    cboPorMonto.Enabled = False
    txtMontoRemunerada.Enabled = False
    
End Sub

Private Sub Habilita()
    
    chkRemunerada.Value = vbUnchecked
        
    If strCodTipoCuenta = Codigo_Tipo_Cuenta_Corriente Then
        cboTasa.Enabled = False
        chkRemunerada.Enabled = True
    Else
        chkRemunerada.Enabled = False
        cboTasa.Enabled = False
        If strCodTipoCuenta = Codigo_Tipo_Cuenta_Ahorro Then cboTasa.Enabled = True
    End If
    
End Sub

Private Sub LlenarFormulario(strModo As String)

    Dim adoRecord As ADODB.Recordset
    
    Select Case strModo
        Case Reg_Adicion
            txtDescripCuenta.Text = Valor_Caracter: txtResponsable.Text = Valor_Caracter
            txtNumCuenta.Text = Valor_Caracter
            txtMontoRemunerada.Text = "0": txtTipoCambio.Text = "0"
            txtMontoMovimientoDestino.Text = "0"
            txtMontoMovimientoOrigen.Text = "0"
            
            cboTipoCuenta.ListIndex = -1
            If cboTipoCuenta.ListCount > 0 Then cboTipoCuenta.ListIndex = 0
            
            cboMoneda.ListIndex = -1
            If cboMoneda.ListCount > 0 Then cboMoneda.ListIndex = 0
            
            cboBanco.ListIndex = -1
            If cboBanco.ListCount > 0 Then cboBanco.ListIndex = 0
            
            cboCuentaActivo.ListIndex = -1
            If cboCuentaActivo.ListCount > 0 Then cboCuentaActivo.ListIndex = 0
            
            cboTipoRemunerada.ListIndex = -1
            If cboTipoRemunerada.ListCount > 0 Then cboTipoRemunerada.ListIndex = 0
            
            cboTasa.ListIndex = -1
            If cboTasa.ListCount > 0 Then cboTasa.ListIndex = 0
            
            cboPorMonto.ListIndex = -1
            If cboPorMonto.ListCount > 0 Then cboPorMonto.ListIndex = 0
            
            dtpFechaApertura.Value = gdatFechaActual
            dtpFechaInicioCalculo.Value = gdatFechaActual
            chkRemunerada.Value = vbUnchecked
                        
            cboTipoCuenta.SetFocus
                        
        Case Reg_Edicion
            Dim intRegistro As Integer
            Dim adoTemporal As ADODB.Recordset
            
            Set adoRecord = New ADODB.Recordset
               
            strCodFile = Trim(tdgConsulta.Columns(0))
            strCodAnalitica = Trim(tdgConsulta.Columns(1))
            
            adoComm.CommandText = "{ call up_ACSelDatosParametro(22,'" & strCodAnalitica & "','" & strCodFile & "','" & strCodFondo & "','" & gstrCodAdministradora & "') }"
            Set adoRecord = adoComm.Execute
            
            If Not adoRecord.EOF Then
                txtDescripCuenta.Text = Trim(adoRecord("DescripCuenta"))
                txtResponsable.Text = Trim(adoRecord("ResponsableCuenta"))
                txtNumCuenta.Text = Trim(adoRecord("NumCuenta"))
                lblSaldoCuentaTransferencia.Caption = "0"
                cboCuentaFondo.Clear
                                                                
                cboTipoCuenta.ListIndex = 0
                intRegistro = ObtenerItemLista(arrTipoCuenta(), adoRecord("TipoCuenta"))
                If intRegistro > 0 Then cboTipoCuenta.ListIndex = intRegistro

                intRegistro = ObtenerItemLista(arrMoneda(), adoRecord("CodMoneda"))
                If intRegistro > 0 Then cboMoneda.ListIndex = intRegistro

                intRegistro = ObtenerItemLista(arrBanco(), adoRecord("CodBanco"))
                If intRegistro > 0 Then cboBanco.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrCuentaActivo(), adoRecord("CodCuentaActivo"))
                If intRegistro > 0 Then cboCuentaActivo.ListIndex = intRegistro
                
                dtpFechaApertura.Value = adoRecord("FechaApertura")
                dtpFechaApertura.Value = adoRecord("FechaInicioCalculo")
                
                intRegistro = ObtenerItemLista(arrCalculo(), adoRecord("CodCalculo"))
                If intRegistro > 0 Then cboCalculo.ListIndex = intRegistro
                
                If Trim(adoRecord("IndRemunerada")) = "X" Then
                    chkRemunerada.Value = vbChecked
                    
                    intRegistro = ObtenerItemLista(arrTipoRemunerada(), adoRecord("TipoRemunerada"))
                    If intRegistro > 0 Then cboTipoRemunerada.ListIndex = intRegistro
                    
                    intRegistro = ObtenerItemLista(arrTasa(), adoRecord("CodTasa"))
                    If intRegistro > 0 Then cboTasa.ListIndex = intRegistro
                    
                    If cboPorMonto.Enabled Then
                        intRegistro = ObtenerItemLista(arrPorMonto(), adoRecord("TipoMontoRemunerada"))
                        If intRegistro > 0 Then cboPorMonto.ListIndex = intRegistro
                        'txtMontoRemunerada.Text = adoRecord("MontoBaseRemunerada")
                    End If
                    
                Else
                    If cboTasa.Enabled Then
                        intRegistro = ObtenerItemLista(arrTasa(), adoRecord("CodTasa"))
                        If intRegistro > 0 Then cboTasa.ListIndex = intRegistro
                    End If
                End If
                
                txtMontoRemunerada.Text = adoRecord("MontoBaseRemunerada")
                                
                lblDescripCuentaOrigen.Caption = Trim(txtDescripCuenta.Text) & Space(1) & Trim(txtNumCuenta.Text)
                'lblFechaMovimiento.Caption = CStr(gdatFechaActual)
            
                cboTipoMovimiento.ListIndex = -1
                If cboTipoMovimiento.ListCount > 0 Then cboTipoMovimiento.ListIndex = 0
            
'                txtTipoCambio.Text = CStr(gdblTipoCambio)
                txtMontoMovimientoDestino.Text = "0"
                txtMontoMovimientoOrigen.Text = "0"
                lblSaldoDisponible.Caption = "0"
                               
                '*** Obtener los saldos de la cuenta ***
                Call SaldoCuenta
            End If
            adoRecord.Close: Set adoRecord = Nothing
    End Select
    
    fraCuenta(0).Caption = "Fondo : " & Trim(cboFondo.Text)
    fraCuenta(1).Caption = "Fondo : " & Trim(cboFondo.Text)
    
End Sub
Public Sub Anterior()

End Sub

Public Sub Ayuda()

End Sub

Public Sub Cancelar()

    cmdOpcion.Visible = True
    With tabCuenta
        .TabEnabled(0) = True
        .TabEnabled(2) = True
        .Tab = 0
    End With
    Call Buscar
    
End Sub

Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Cuentas Bancarias"
'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Visible = True
'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Text = "Movimientos Caja/Cuenta Corriente"
    
End Sub

Public Sub Eliminar()

'/**/ HMC
If dblSaldoDisponible = 0 And Trim(tdgConsulta.Columns(0).Text) <> "" Then
    If dblSaldoDisponible = 0 Then
        If MsgBox("Esta Seguro que desea Cerrar esta Cuenta", vbQuestion + vbYesNo, "Observación") = vbYes Then
            MousePointer = vbHourglass
            
            strCodFile = Trim(tdgConsulta.Columns(0))
            strCodAnalitica = Trim(tdgConsulta.Columns(1))
            
            adoComm.CommandText = "UPDATE BancoCuenta SET IndVigente = '' " & _
                "WHERE (CodFondo = '" & strCodFondo & "' AND CodAdministradora = '" & gstrCodAdministradora & "' AND " & _
                "CodFile = '" & strCodFile & "' AND CodAnalitica = '" & strCodAnalitica & "')"
            adoConn.Execute adoComm.CommandText
            
            MousePointer = vbDefault
            tabCuenta.Tab = 0
            Call Buscar
        Else
            Exit Sub
        End If
    Else
        MsgBox "No se Puede Cerrar esta Cuenta ya que tiene un Saldo Vigente", vbCritical, "Observación"
    End If
End If
'/**/ HMC

End Sub


Public Sub Exportar()

End Sub

Public Sub Grabar()

    On Error GoTo Error1            '/**/ HMC Habilitamos la rutina de Errores.

    Dim adoRegistro As ADODB.Recordset
    
    If strEstado = Reg_Consulta Then Exit Sub
    
    If strEstado = Reg_Adicion Then
        If TodoOk() Then
            Me.MousePointer = vbHourglass
            '*** Guardar Cuentas del Fondo ***
            With adoComm
                .CommandText = "{ call up_TSManBancoCuenta('"
                .CommandText = .CommandText & strCodFondo & "','"
                .CommandText = .CommandText & gstrCodAdministradora & "','"
                .CommandText = .CommandText & strCodFile & "','"
                .CommandText = .CommandText & strCodAnalitica & "','"
                .CommandText = .CommandText & strCodTipoCuenta & "','"
                .CommandText = .CommandText & strCodBanco & "','"
                .CommandText = .CommandText & Trim(txtNumCuenta.Text) & "','"
                .CommandText = .CommandText & strCodMoneda & "','"
                .CommandText = .CommandText & Trim(txtDescripCuenta.Text) & "','"
                .CommandText = .CommandText & Trim(txtResponsable.Text) & "','"
                .CommandText = .CommandText & strCodCuentaActivo & "','"
                .CommandText = .CommandText & Convertyyyymmdd(dtpFechaApertura.Value) & "','"
                .CommandText = .CommandText & Convertyyyymmdd(dtpFechaInicioCalculo.Value) & "','"
                .CommandText = .CommandText & "X','"
                .CommandText = .CommandText & strCodTasa & "','"
                .CommandText = .CommandText & strCodCalculo & "','"
                If chkRemunerada.Value Then
                    .CommandText = .CommandText & "X','"
                Else
                    .CommandText = .CommandText & "','"
                End If
                .CommandText = .CommandText & strCodTipoRemunerada & "','"
                .CommandText = .CommandText & CDec(txtMontoRemunerada.Text) & "','"
                .CommandText = .CommandText & strCodPorMonto & "','"
                .CommandText = .CommandText & gstrLogin & "','"
                .CommandText = .CommandText & Convertyyyymmdd(gdatFechaActual) & "','"
                .CommandText = .CommandText & gstrLogin & "','"
                .CommandText = .CommandText & Convertyyyymmdd(gdatFechaActual) & "','"
                .CommandText = .CommandText & "I') }"
                
                adoConn.Execute .CommandText
                
'                Set adoRegistro = New ADODB.Recordset
'
'                .CommandText = "SELECT COUNT(*) NumReg FROM FondoCuenta WHERE TipoCuenta='" & Codigo_Tipo_Cuenta_Corriente & "' AND "
'                .CommandText = .CommandText & "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
'                Set adoRegistro = .Execute
'
'                If Not adoRegistro.EOF Then
'                    If IsNull(adoRegistro("NumReg")) Or CInt(adoRegistro("NumReg")) = 0 Then
                        .CommandText = "{ call up_TSManFondoCuenta('"
                        .CommandText = .CommandText & strCodFondo & "','"
                        .CommandText = .CommandText & gstrCodAdministradora & "','"
                        .CommandText = .CommandText & strCodBanco & "','"
                        .CommandText = .CommandText & Codigo_Operacion_Suscripcion & "','"
                        .CommandText = .CommandText & strCodTipoCuenta & "','"
                        .CommandText = .CommandText & Trim(txtNumCuenta.Text) & "','"
                        .CommandText = .CommandText & strCodCuentaActivo & "','"
                        .CommandText = .CommandText & strCodFile & "','"
                        .CommandText = .CommandText & strCodAnalitica & "','"
                        .CommandText = .CommandText & strCodCuentaActivo & "','"
                        .CommandText = .CommandText & "I') }"
                        
                        adoConn.Execute .CommandText
                        
                        .CommandText = "{ call up_TSManFondoCuenta('"
                        .CommandText = .CommandText & strCodFondo & "','"
                        .CommandText = .CommandText & gstrCodAdministradora & "','"
                        .CommandText = .CommandText & strCodBanco & "','"
                        .CommandText = .CommandText & Codigo_Operacion_Rescate & "','"
                        .CommandText = .CommandText & strCodTipoCuenta & "','"
                        .CommandText = .CommandText & Trim(txtNumCuenta.Text) & "','"
                        .CommandText = .CommandText & strCodCuentaActivo & "','"
                        .CommandText = .CommandText & strCodFile & "','"
                        .CommandText = .CommandText & strCodAnalitica & "','"
                        .CommandText = .CommandText & strCodCuentaActivo & "','"
                        .CommandText = .CommandText & "I') }"
                        
                        adoConn.Execute .CommandText
'                    End If
'                End If
'                adoRegistro.Close: Set adoRegistro = Nothing
                                                
            End With
                                                                                    
            Me.MousePointer = vbDefault
        
            MsgBox Mensaje_Adicion_Exitosa, vbExclamation
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
            cmdOpcion.Visible = True
            With tabCuenta
                .TabEnabled(0) = True
                .Tab = 0
            End With
            Call Buscar
        End If
    End If

    If strEstado = Reg_Edicion Then
        If TodoOk() Then
        
            Me.MousePointer = vbHourglass
            '*** Actualizar Cuentas del Fondo ***
            With adoComm
                .CommandText = "{ call up_TSManBancoCuenta('"
                .CommandText = .CommandText & strCodFondo & "','"
                .CommandText = .CommandText & gstrCodAdministradora & "','"
                .CommandText = .CommandText & strCodFile & "','"
                .CommandText = .CommandText & strCodAnalitica & "','"
                .CommandText = .CommandText & strCodTipoCuenta & "','"
                .CommandText = .CommandText & strCodBanco & "','"
                .CommandText = .CommandText & Trim(txtNumCuenta.Text) & "','"
                .CommandText = .CommandText & strCodMoneda & "','"
                .CommandText = .CommandText & Trim(txtDescripCuenta.Text) & "','"
                .CommandText = .CommandText & Trim(txtResponsable.Text) & "','"
                .CommandText = .CommandText & strCodCuentaActivo & "','"
                .CommandText = .CommandText & Convertyyyymmdd(dtpFechaApertura.Value) & "','"
                .CommandText = .CommandText & Convertyyyymmdd(dtpFechaInicioCalculo.Value) & "','"
                .CommandText = .CommandText & "X','"
                .CommandText = .CommandText & strCodTasa & "','"
                .CommandText = .CommandText & strCodCalculo & "','"
                If chkRemunerada.Value Then
                    .CommandText = .CommandText & "X','"
                Else
                    .CommandText = .CommandText & "','"
                End If
                .CommandText = .CommandText & strCodTipoRemunerada & "','"
                .CommandText = .CommandText & CDec(txtMontoRemunerada.Text) & "','"
                .CommandText = .CommandText & strCodPorMonto & "','"
                .CommandText = .CommandText & gstrLogin & "','"
                .CommandText = .CommandText & Convertyyyymmdd(gdatFechaActual) & "','"
                .CommandText = .CommandText & gstrLogin & "','"
                .CommandText = .CommandText & Convertyyyymmdd(gdatFechaActual) & "','"
                .CommandText = .CommandText & "U') }"
                
                adoConn.Execute .CommandText
                
                .CommandText = "{ call up_TSManFondoCuenta('"
                .CommandText = .CommandText & strCodFondo & "','"
                .CommandText = .CommandText & gstrCodAdministradora & "','"
                .CommandText = .CommandText & strCodBanco & "','"
                .CommandText = .CommandText & Codigo_Operacion_Suscripcion & "','"
                .CommandText = .CommandText & strCodTipoCuenta & "','"
                .CommandText = .CommandText & Trim(txtNumCuenta.Text) & "','"
                .CommandText = .CommandText & strCodCuentaActivo & "','"
                .CommandText = .CommandText & strCodFile & "','"
                .CommandText = .CommandText & strCodAnalitica & "','"
                .CommandText = .CommandText & strCodCuentaActivo & "','"
                .CommandText = .CommandText & "U') }"
                        
                adoConn.Execute .CommandText
                
            End With

            Me.MousePointer = vbDefault
            
            MsgBox Mensaje_Edicion_Exitosa, vbExclamation
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
            cmdOpcion.Visible = True
            With tabCuenta
                .TabEnabled(0) = True
                .Tab = 0
            End With
            Call Buscar
        End If
    End If
    
On Error GoTo 0                  '/**/
Exit Sub                         '/**/
Error1:     MsgBox DescripcionError & vbNewLine & DescripcionTecnica & err.Description, vbExclamation, TituloError ' Mostrar Error
Me.MousePointer = vbDefault      '/**/
   
End Sub

Public Sub Importar()

End Sub

Public Sub Imprimir()

End Sub

Public Sub Modificar()

    If strEstado = Reg_Consulta Then
        strEstado = Reg_Edicion
        Call Deshabilita
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabCuenta
            .TabEnabled(0) = False
            .TabEnabled(2) = True
            .Tab = 1
        End With
    End If
    
End Sub



Public Sub Primero()

End Sub

Public Sub Refrescar()

End Sub

Public Sub Salir()

    Unload Me
    
End Sub




Public Sub Seguridad()

End Sub

Public Sub Siguiente()

End Sub

Public Sub SubImprimir(index As Integer)

    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    Dim strFechaDesde           As String, strFechaHasta        As String

    If tabCuenta.Tab = 1 Then Exit Sub
    
    Select Case index
        Case 1
            gstrNameRepo = "CuentasBancarias"
                        
            Set frmReporte = New frmVisorReporte

            ReDim aReportParamS(1)
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
            
    End Select

    gstrSelFrml = Valor_Caracter
    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"

    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())

    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
    frmReporte.Show vbModal

    Set frmReporte = Nothing

    Screen.MousePointer = vbNormal
    
End Sub

Public Sub Ultimo()

End Sub

Private Sub cboBanco_Click()

    strCodBanco = Valor_Caracter
    If cboBanco.ListIndex < 0 Then Exit Sub
    
    strCodBanco = Trim(arrBanco(cboBanco.ListIndex))
    txtDescripCuenta.Text = Trim(cboTipoCuenta.Text) & Space(1) & strSignoMoneda & Space(1) & Trim(cboBanco.Text)
    
End Sub


Private Sub cboCalculo_Click()

    strCodCalculo = ""
    If cboCalculo.ListIndex < 0 Then Exit Sub
    
    strCodCalculo = Trim(arrCalculo(cboCalculo.ListIndex))
    
End Sub


Private Sub cboCuentaActivo_Click()

    strCodCuentaActivo = Valor_Caracter
    If cboCuentaActivo.ListIndex < 0 Then Exit Sub
    
    strCodCuentaActivo = Trim(arrCuentaActivo(cboCuentaActivo.ListIndex))
    
End Sub


Private Sub cboCuentaFondo_Click()

    Dim curSaldoCuenta  As Currency
    Dim datFecha        As Date, datFechaSiguiente  As Date

    
    strCodFileCuenta = Valor_Caracter
    strCodAnaliticaCuenta = Valor_Caracter
    strCodCuentaFondo = Valor_Caracter
    strCodMonedaCuenta = Valor_Caracter
    lblSaldoCuentaTransferencia.Caption = "0"
        
    If cboCuentaFondo.ListIndex < 0 Then Exit Sub
    
    strCodFileCuenta = Left(arrCuentaFondo(cboCuentaFondo.ListIndex), 3)
    strCodAnaliticaCuenta = Mid(arrCuentaFondo(cboCuentaFondo.ListIndex), 4, 8)
    strCodCuentaFondo = Trim(Mid(arrCuentaFondo(cboCuentaFondo.ListIndex), 12, 10))
    strCodMonedaCuenta = Trim(Mid(arrCuentaFondo(cboCuentaFondo.ListIndex), 22, 2))
    strCodBancoCuenta = Trim(Right(arrCuentaFondo(cboCuentaFondo.ListIndex), 8))
    
    datFecha = CVDate(lblFechaMovimiento.Caption)
    datFechaSiguiente = DateAdd("d", 1, datFecha)
    
    curSaldoCuenta = ObtenerSaldoFinalCuenta(strCodFondo, gstrCodAdministradora, strCodFileCuenta, strCodAnaliticaCuenta, Convertyyyymmdd(datFecha), Convertyyyymmdd(datFechaSiguiente), strCodCuentaFondo, strCodMonedaCuenta)
    lblSaldoCuentaTransferencia.Caption = CStr(curSaldoCuenta)
    
    strSignoMonedaCuenta = ObtenerSignoMoneda(strCodMonedaCuenta)
    strCodSignoMonedaCuenta = ObtenerCodSignoMoneda(strCodMonedaCuenta)
    
    strCodMonedaParEvaluacion = strCodMoneda & strCodMonedaCuenta
    
    If strCodMoneda <> strCodMonedaCuenta Then
        strCodMonedaParPorDefecto = ObtenerMonedaParPorDefecto(gstrCodClaseTipoCambioOperacionFondo, strCodMonedaParEvaluacion)
    Else
        strCodMonedaParPorDefecto = strCodMonedaParEvaluacion
    End If
    
    If strCodMonedaParPorDefecto = "0000" Then strCodMonedaParPorDefecto = strCodMonedaParEvaluacion
    
    lblDescripTCArbitraje.Caption = "(" + Trim(ObtenerCodSignoMoneda(Mid(strCodMonedaParPorDefecto, 3, 2))) + "/" + Trim(ObtenerCodSignoMoneda(Mid(strCodMonedaParPorDefecto, 1, 2))) + ")"
    
    If strCodMoneda <> strCodMonedaCuenta Then
        ChkMonDiferente.Value = vbChecked
        txtTipoCambio.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, gdatFechaActual, Mid(strCodMonedaParPorDefecto, 1, 2), Mid(strCodMonedaParPorDefecto, 3, 2)))
        txtTipoCambio.Enabled = True
    Else
        ChkMonDiferente.Value = vbUnchecked
        txtTipoCambio.Text = "1"
        txtTipoCambio.Enabled = False
    End If
    Call txtTipoCambio_KeyPress(vbKeyReturn)
    
    
    lblDescrip(23).Caption = "Saldo Disponible" & Space(1) & strSignoMonedaCuenta
    
    lblMonedaOrigen.Caption = strCodSignoMoneda
    lblMonedaDestino.Caption = strCodSignoMonedaCuenta
            
End Sub


Private Sub cboFondo_Click()

    Dim adoRegistro     As ADODB.Recordset
    Dim strNumRucFondo  As String
    
    
    strCodFondo = Valor_Caracter
    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(arrFondo(cboFondo.ListIndex))
    
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Fecha Vigente, Valor Cuota, Valor Cuota T-1, Moneda y Cantidad Inicial de Cuotas del Fondo ***
        .CommandText = "{ call up_ACSelDatosParametro(23,'" & strCodFondo & "','" & gstrCodAdministradora & "') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            lblFechaMovimiento.Caption = CStr(adoRegistro("FechaCuota"))
            gdatFechaActual = adoRegistro("FechaCuota")
            frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)

            txtTipoCambio.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, gdatFechaActual, Trim(adoRegistro("CodMoneda")), Codigo_Moneda_Local))
            If CDbl(txtTipoCambio.Text) = 0 Then txtTipoCambio.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, DateAdd("d", -1, gdatFechaActual), Trim(adoRegistro("CodMoneda")), Codigo_Moneda_Local))
        End If
        
        adoRegistro.Close
        
        'ACTUALIZA PARAMETROS GLOBALES POR FONDO
        If Not CargarParametrosGlobales(strCodFondo) Then Exit Sub
        
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


Private Sub cboMoneda_Click()
        
    strCodMoneda = Valor_Caracter: strSignoMoneda = Valor_Caracter
    If cboMoneda.ListIndex < 0 Then Exit Sub
    
    strCodMoneda = Trim(arrMoneda(cboMoneda.ListIndex))
    
    strSignoMoneda = ObtenerSignoMoneda(strCodMoneda)
    strCodSignoMoneda = ObtenerCodSignoMoneda(strCodMoneda)
    
    txtDescripCuenta.Text = Trim(cboTipoCuenta.Text) & Space(1) & strSignoMoneda & Space(1) & Trim(cboBanco.Text)
    lblDescrip(14).Caption = "Saldo Disponible" & Space(1) & strSignoMoneda
    
    If strCodMoneda <> Valor_Caracter Then
        '*** Tasa ***
        strSQL = "SELECT TasaInteresBancaria.CodTasa CODIGO, DescripTasa DESCRIP FROM TasaInteresBancaria,TasaInteresBancariaDetalle " & _
            "WHERE TasaInteresBancaria.CodTasa = TasaInteresBancariaDetalle.CodTasa AND IndVigente='X' AND TasaInteresBancaria.CodMoneda='" & strCodMoneda & "'"
        CargarControlLista strSQL, cboTasa, arrTasa(), Sel_Defecto
        
        If cboTasa.ListCount > 0 Then cboTasa.ListIndex = 0
    End If
    
    '*** Cuenta Activo ***
    strSQL = "{ call up_TELstCuentaActivoBanco('" & gstrCodAdministradora & "','" & strCodTipoCuenta & "','" & strCodMoneda & "') }"
    CargarControlLista strSQL, cboCuentaActivo, arrCuentaActivo(), Sel_Defecto
    If cboCuentaActivo.ListCount > 0 Then cboCuentaActivo.ListIndex = 0
    
End Sub


Private Sub cboPorMonto_Click()

    strCodPorMonto = ""
    If cboPorMonto.ListIndex < 0 Then Exit Sub
    
    strCodPorMonto = Trim(arrPorMonto(cboPorMonto.ListIndex))
    
End Sub


Private Sub cboTasa_Click()

    strCodTasa = ""
    If cboTasa.ListIndex < 0 Then Exit Sub
    
    strCodTasa = Trim(arrTasa(cboTasa.ListIndex))
    
End Sub


Private Sub cboTipoCuenta_Click()

    Dim adoRecord As ADODB.Recordset
    
    strCodTipoCuenta = Valor_Caracter
    If cboTipoCuenta.ListIndex < 0 Then Exit Sub
    
    strCodTipoCuenta = Trim(arrTipoCuenta(cboTipoCuenta.ListIndex))
    
    If strCodTipoCuenta = Codigo_Tipo_Cuenta_Corriente Or strCodTipoCuenta = Codigo_Tipo_Cuenta_Ahorro Then
        Call Habilita
    Else
        Call Deshabilita
    End If
    
    Set adoRecord = New ADODB.Recordset
            
    If strEstado = Reg_Adicion Then
        adoComm.CommandText = "SELECT MAX(CodAnalitica) CodAnalitica FROM BancoCuenta WHERE TipoCuenta='" & strCodTipoCuenta & "' AND CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
        Set adoRecord = adoComm.Execute
        
        If Not adoRecord.EOF Then
            If IsNull(adoRecord("CodAnalitica")) Then
                strCodAnalitica = "00000001"
            Else
                strCodAnalitica = Format(adoRecord("CodAnalitica") + 1, "00000000")
            End If
        Else
            strCodAnalitica = "00000001"
        End If
        adoRecord.Close: Set adoRecord = Nothing
    End If
    
    strCodFile = "0" & strCodTipoCuenta
    lblAnalitica.Caption = strCodFile & "-" & strCodAnalitica
    txtDescripCuenta.Text = Trim(cboTipoCuenta.Text) & Space(1) & strSignoMoneda & Space(1) & Trim(cboBanco.Text)
    
    '*** Cuenta Activo ***
    strSQL = "{ call up_TELstCuentaActivoBanco('" & gstrCodAdministradora & "','" & strCodTipoCuenta & "','" & strCodMoneda & "') }"
    CargarControlLista strSQL, cboCuentaActivo, arrCuentaActivo(), Sel_Defecto
    If cboCuentaActivo.ListCount > 0 Then cboCuentaActivo.ListIndex = 0
    
End Sub


Private Sub cboTipoMovimiento_Click()

    Dim strSQL As String
    
    strCodTipoMovimiento = ""
    If cboTipoMovimiento.ListIndex < 0 Then Exit Sub
    
    strCodTipoMovimiento = Trim(arrTipoMovimiento(cboTipoMovimiento.ListIndex))
        
    If strCodTipoMovimiento = Codigo_Movimiento_Deposito Then
        lblDescrip(22).Caption = "Transferir De ..."
        lblDescripOrigen.Caption = "Monto Transferencia Destino"
        lblDescripDestino.Caption = "Monto Transferencia Origen"
    Else
        lblDescrip(22).Caption = "Transferir A ..."
        fraInicio.Caption = "Origen"
        fraFinal.Caption = "Destino"
        lblDescripOrigen.Caption = "Monto Transferencia Origen"
        lblDescripDestino.Caption = "Monto Transferencia Destino"
    End If
    
    If strEstado = Reg_Edicion And strCodTipoMovimiento <> Valor_Caracter Then
        strSQL = "SELECT (CodFile + CodAnalitica + CodCuentaActivo + CodMoneda + CodBanco) CODIGO,(DescripCuenta + space(1) + NumCuenta) DESCRIP " & _
            "FROM BancoCuenta " & _
            "WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
            "(CodFile + CodAnalitica)<>'" & strCodFile & strCodAnalitica & "'" ' AND CodMoneda='" & strCodMoneda & "'"
        
        CargarControlLista strSQL, cboCuentaFondo, arrCuentaFondo(), ""
    
        If cboCuentaFondo.ListCount > 0 Then
            cboCuentaFondo.ListIndex = 0
        Else
            MsgBox "El Fondo no tiene otras cuentas definidas...", vbCritical, gstrNombreEmpresa
        End If
    End If
    
End Sub


Private Sub cboTipoRemunerada_Click()

    strCodTipoRemunerada = ""
    If cboTipoRemunerada.ListIndex < 0 Then Exit Sub
    
    strCodTipoRemunerada = Trim(arrTipoRemunerada(cboTipoRemunerada.ListIndex))
    
    If strCodTipoRemunerada = Codigo_Tipo_Remunerada_Monto Then
        cboPorMonto.Enabled = True
        txtMontoRemunerada.Enabled = True
    Else
        cboPorMonto.Enabled = False
        txtMontoRemunerada.Enabled = False
    End If
    
End Sub


Private Sub chkModificaTC_Click()

'If chkModificaTC.Value = vbChecked Then
'    txtTipoCambio.Enabled = True
'End If
'
'If chkModificaTC.Value = vbUnchecked Then
'    txtTipoCambio.Enabled = False
'End If



End Sub

Private Sub chkRemunerada_Click()

    If cboMoneda.ListIndex <= 0 Then
        MsgBox "Seleccione la moneda primero", vbCritical, gstrNombreEmpresa
    Else

        cboTipoRemunerada.ListIndex = 0
        cboTasa.ListIndex = 0
        cboPorMonto.ListIndex = 0
        txtMontoRemunerada = "0"
        
        If chkRemunerada.Value Then
            cboTipoRemunerada.Enabled = True
        Else
            cboTipoRemunerada.Enabled = False
            If strCodTipoCuenta = Codigo_Tipo_Cuenta_Corriente Then cboTasa.Enabled = False
        End If
    
    End If
    
End Sub
Private Function ValidaOk()

    ValidaOk = False
    
    If cboTipoMovimiento.ListIndex = 0 Then
        MsgBox "Seleccione el tipo de movimiento", vbCritical, gstrNombreEmpresa
        cboTipoMovimiento.SetFocus
        Exit Function
    End If
    
    If CCur(txtMontoMovimientoDestino.Text) = 0 Or CCur(txtMontoMovimientoOrigen.Text) = 0 Then
        MsgBox "Monto del movimientos no puede ser cero", vbCritical, gstrNombreEmpresa
        txtMontoMovimientoDestino.SetFocus
        Exit Function
    End If

    If strCodTipoMovimiento = Codigo_Movimiento_Retiro Then 'Codigo_Movimiento_Deposito Then
        If CCur(txtMontoMovimientoOrigen.Text) > CCur(lblSaldoDisponible.Caption) Then
            If MsgBox("Monto del movimiento origen es mayor al saldo de la cuenta. Va a sobregirar la cuenta." & vbNewLine & vbNewLine & _
                "Seguro de Continuar ?", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                txtMontoMovimientoOrigen.SetFocus
                Exit Function
            End If
        End If
    Else
        If CCur(txtMontoMovimientoDestino.Text) > CCur(lblSaldoCuentaTransferencia.Caption) Then
            If MsgBox("Monto del movimiento destino es mayor al saldo de la cuenta. Va a sobregirar la cuenta." & vbNewLine & vbNewLine & _
                "Seguro de Continuar ?", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                txtMontoMovimientoDestino.SetFocus
                Exit Function
            End If
        End If
    End If
    
    'Verificar existencia de tipos de cambio
    If strCodMoneda <> Codigo_Moneda_Local Then
                       
        If CDbl(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, gdatFechaActual, strCodMoneda, Codigo_Moneda_Local)) = 0 Then
            MsgBox "No existe tipo de cambio disponible!", vbCritical, gstrNombreEmpresa
            Exit Function
        End If
        
'        If CDbl(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, gdatFechaActual, strCodMoneda, Codigo_Moneda_Local, Tipo_Busqueda_Tipo_Cambio_Iterativo_Directo)) = 0 Then
'            If CDbl(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, gdatFechaActual, strCodMoneda, Codigo_Moneda_Local, Tipo_Busqueda_Tipo_Cambio_Iterativo_Inverso)) = 0 Then
'                MsgBox "No existe tipo de cambio disponible!", vbCritical, gstrNombreEmpresa
'                Exit Function
'            End If
'        End If
    
    End If
    
    If Trim(txtNroTransaccion.Text) = Valor_Caracter Then
        MsgBox "Número de Transacción no ingresado.", vbCritical, gstrNombreEmpresa
        txtNroTransaccion.SetFocus
        Exit Function
    End If
    
    ValidaOk = True

End Function
Private Sub ConfiguraRecordsetAuxiliar()

    Set adoRegistroAux = New ADODB.Recordset

    With adoRegistroAux
       .CursorLocation = adUseClient
       .Fields.Append "CodMonedaOrigen", adChar, 2
       .Fields.Append "CodMonedaCambio", adChar, 2
       .Fields.Append "ValorTipoCambio", adDecimal
       .Fields.Item("ValorTipoCambio").Precision = 20
       .Fields.Item("ValorTipoCambio").NumericScale = 12
       .LockType = adLockBatchOptimistic
    End With

    adoRegistroAux.Open

End Sub

Private Sub cmdRegistrar_Click()
    
    Dim curMontoMovimiento      As Double, intRegistro          As Integer
    Dim strDescripMovimiento    As String, strNumOrdenCobroPago As String
    Dim strIndDebeHaber         As String
    Dim strFechaAsiento         As String
   
    Dim strNumAsiento           As String, strCodCuentaMovimiento         As String
    Dim strCodFileMovimiento    As String, strCodAnaliticaMovimiento    As String
    Dim strCodMonedaMovimiento  As String, strCodBancoMovimiento        As String
    Dim intCantMovAsiento       As Integer, intContador                 As Integer
    Dim curMontoMN              As Currency, curMontoME                 As Currency
    Dim curMontoContable        As Currency
    Dim dblTipoCambio           As Double
    Dim dblTipoCambioAjuste     As Double
    
    Dim strBancoCuentaMovimiento As String
    Dim strOrdenCobroPago       As String
    Dim dblMontoPago            As Double
    Dim strTipoMovimiento       As String
    
    Dim strCodCuentaContraparte     As String
    Dim strCodFileContraparte       As String
    Dim strCodAnaliticaContraparte  As String
    
    'Manejo de Transacciones
    Dim adoError                    As ADODB.Error
    Dim strErrMsg                   As String
    Dim blnClientTran               As Boolean
    Dim intAccion                   As Integer
    Dim lngNumError                 As Long
    Dim curMontoContableAcum        As Currency
    Dim strTipoAuxiliar             As String
    Dim strCodAuxiliar              As String
    
    Dim strCodMonedaParEvaluacionOrigen As String
    Dim strCodMonedaParEvaluacionDestino As String
    Dim strCodMonedaParPorDefectoOrigen As String
    Dim strCodMonedaParPorDefectoDestino As String
    
    Dim dblTipoCambioOrigen As Double
    Dim dblTipoCambioDestino As Double
    
    Dim strMsgError                     As String
    Dim objTipoCambioReemplazoXML       As DOMDocument60
    Dim strTipoCambioReemplazoXML       As String
    Dim strIndUltimoMovimiento          As String
    Dim strIndSoloMovimientoContable    As String

    Dim dblValorTipoCambio              As Double, strTipoDocumento As String
    Dim strNumDocumento                 As String
    Dim strIndContracuenta              As String
    Dim strTipoPersonaContraparte       As String, strCodPersonaContraparte As String

    Dim strCodContracuenta              As String
    Dim strCodFileContracuenta          As String
    Dim strCodAnaliticaContracuenta     As String

    On Error GoTo CtrlError
    
    If Not ValidaOk() Then Exit Sub
    
    strTipoAuxiliar = Valor_Caracter
    strCodAuxiliar = Valor_Caracter
    strIndUltimoMovimiento = Valor_Caracter
    strIndSoloMovimientoContable = Valor_Caracter
    
    If strCodTipoMovimiento = Codigo_Movimiento_Deposito Then
        strDescripMovimiento = "DEPOSITO EN CUENTA " & lblAnalitica.Caption
        curMontoMovimiento = CCur(txtMontoMovimientoOrigen.Text)
        strIndDebeHaber = "D"
    Else
        strDescripMovimiento = "RETIRO DE CUENTA " & lblAnalitica.Caption
        curMontoMovimiento = CCur(txtMontoMovimientoOrigen.Text) * -1
        strIndDebeHaber = "H"
    End If
    
    'MonedaOrigen <> MonedaContable -->TipoCambio MonedaOrigen-MonedaContable
    'MonedaOrigen = MonedaContable  -->No usa Tipo de Cambio
    'MonedaDestino <> MonedaContable -->TipoCambio MonedaDestino-MonedaContable
    'MonedaDestino = MonedaContable -->No usa Tipo de Cambio

    'PRIMERO OBTENEMOS T/C DE LA OPERACION DE CAMBIO PROPIA
    If strCodMoneda <> strCodMonedaCuenta Then
        Call ConfiguraRecordsetAuxiliar
        
        adoRegistroAux.AddNew
        
        adoRegistroAux.Fields("CodMonedaOrigen") = Mid(strCodMonedaParPorDefecto, 1, 2) 'strCodMoneda
        adoRegistroAux.Fields("CodMonedaCambio") = Mid(strCodMonedaParPorDefecto, 3, 2) 'strCodMonedaCuenta
        adoRegistroAux.Fields("ValorTipoCambio") = CDbl(txtTipoCambio.Text)
        
        Call XMLADORecordset(objTipoCambioReemplazoXML, "TipoCambioReemplazo", "MonedaTipoCambio", adoRegistroAux, strMsgError)
        strTipoCambioReemplazoXML = objTipoCambioReemplazoXML.xml 'CrearXMLDetalle(objTipoCambioReemplazoXML)
        
        adoRegistroAux.Close
    Else
        strTipoCambioReemplazoXML = XML_TipoCambioReemplazo
    End If
    
    'ORIGEN
    strCodMonedaParEvaluacionOrigen = strCodMoneda & Codigo_Moneda_Local
    
    If strCodMoneda <> Codigo_Moneda_Local Then
        strCodMonedaParPorDefectoOrigen = ObtenerMonedaParPorDefecto(gstrCodClaseTipoCambioOperacionFondo, strCodMonedaParEvaluacionOrigen)
    Else
        strCodMonedaParPorDefectoOrigen = strCodMonedaParEvaluacionOrigen
    End If
    
    If strCodMonedaParPorDefectoOrigen = "0000" Then strCodMonedaParPorDefectoOrigen = strCodMonedaParEvaluacionOrigen
    
    If strCodMoneda <> Codigo_Moneda_Local Then
        dblTipoCambioOrigen = ObtenerTipoCambioMonedaXML(strCodMoneda, Codigo_Moneda_Local, Convertyyyymmdd(CVDate(lblFechaMovimiento.Caption)), strTipoCambioReemplazoXML, gstrCodClaseTipoCambioOperacionFondo)
    Else
        dblTipoCambioOrigen = 1
    End If
    
    'DESTINO
    strCodMonedaParEvaluacionDestino = strCodMonedaCuenta & Codigo_Moneda_Local
    
    If strCodMonedaCuenta <> Codigo_Moneda_Local Then
        strCodMonedaParPorDefectoDestino = ObtenerMonedaParPorDefecto(gstrCodClaseTipoCambioOperacionFondo, strCodMonedaParEvaluacionDestino)
    Else
        strCodMonedaParPorDefectoDestino = strCodMonedaParEvaluacionDestino
    End If
    
    If strCodMonedaParPorDefectoDestino = "0000" Then strCodMonedaParPorDefectoDestino = strCodMonedaParEvaluacionDestino
    
    If strCodMonedaCuenta <> Codigo_Moneda_Local Then
        dblTipoCambioDestino = ObtenerTipoCambioMonedaXML(strCodMonedaCuenta, Codigo_Moneda_Local, Convertyyyymmdd(CVDate(lblFechaMovimiento.Caption)), strTipoCambioReemplazoXML, gstrCodClaseTipoCambioOperacionFondo)
    Else
        dblTipoCambioDestino = 1
    End If
    
    curMontoContableAcum = 0
            
    With adoComm
        
        frmMainMdi.stbMdi.Panels(3).Text = "Procesando Transferencia..."
        
        .CommandText = "BEGIN TRAN OrdTrans"
        .Execute
        
        blnClientTran = True
        
        If ChkMonDiferente.Value Then
            strDescripMovimiento = "OPERACION DE CAMBIO"
        Else
            strDescripMovimiento = "TRANSFERENCIA ENTRE CUENTAS"
        End If
        intCantMovAsiento = 2
        
        .CommandType = adCmdStoredProc
        '*** Obtener el número del parámetro **
        .CommandText = "up_ACObtenerUltNumero"
        .Parameters.Append .CreateParameter("CodFondo", adChar, adParamInput, 3, strCodFondo)
        .Parameters.Append .CreateParameter("CodAdministradora", adChar, adParamInput, 3, gstrCodAdministradora)
        .Parameters.Append .CreateParameter("CodParametro", adChar, adParamInput, 6, Valor_NumComprobante)
        .Parameters.Append .CreateParameter("NuevoNumero", adChar, adParamOutput, 10, "")
        .Execute
            
        If Not .Parameters("NuevoNumero") Then
            strNumAsiento = .Parameters("NuevoNumero").Value
            .Parameters.Delete ("CodFondo")
            .Parameters.Delete ("CodAdministradora")
            .Parameters.Delete ("CodParametro")
            .Parameters.Delete ("NuevoNumero")
        End If
        
        .CommandType = adCmdText
                               
        strFechaAsiento = Convertyyyymmdd(CVDate(lblFechaMovimiento.Caption)) & Space(1) & Format(Time, "hh:ss")
                                                       
        '-----
        If Trim(txtDescripcionTransf.Text) <> Valor_Caracter Then
            strDescripMovimiento = Trim(txtDescripcionTransf.Text)
        End If
        '-----
                                       
        '*** Cabecera Asiento Contable***
        .CommandText = "{ call up_ACAdicAsientoContable('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
            strNumAsiento & "','" & strFechaAsiento & "','" & _
            gstrPeriodoActual & "','" & gstrMesActual & "','','" & _
            strDescripMovimiento & "','" & strCodMoneda & "','" & _
            Codigo_Moneda_Local & "','',''," & _
            CDec(txtMontoMovimientoOrigen.Text) & ",'" & Estado_Activo & "'," & _
            intCantMovAsiento & ",'" & _
            Convertyyyymmdd(CVDate(lblFechaMovimiento.Caption)) & Space(1) & Format(Time, "hh:ss") & "','" & _
            frmMainMdi.Tag & "',''," & _
            CDec(txtTipoCambio.Text) & ",'','','" & _
            strDescripMovimiento & "','','X','') }"
            
        'CDec(txtTipoCambio.Text)
            
        adoConn.Execute .CommandText
       
        strBancoCuentaMovimiento = ObtenerSecuencialInversionOperacion(strCodFondo, Valor_NumOpeCajaBancos)
                
        For intContador = 1 To intCantMovAsiento
            
            dblTipoCambio = 0
            
            If intContador = 1 Then
                'Cuenta Movimiento
                strCodCuentaMovimiento = strCodCuentaActivo
                strCodFileMovimiento = strCodFile
                strCodAnaliticaMovimiento = strCodAnalitica
                strCodMonedaMovimiento = strCodMoneda
                strCodBancoMovimiento = strCodBanco
                'Cuenta Contraparte
                strCodContracuenta = strCodCuentaFondo
                strCodFileContracuenta = strCodFileCuenta
                strCodAnaliticaContracuenta = strCodAnaliticaCuenta
                'Movimiento
                If strCodTipoMovimiento = Codigo_Movimiento_Deposito Then
                    curMontoContable = CCur(txtMontoMovimientoOrigen.Text)
                Else
                    curMontoContable = CCur(txtMontoMovimientoOrigen.Text) * -1
                End If
            
                dblTipoCambio = dblTipoCambioOrigen 'ObtenerTipoCambioArbitraje(dblTipoCambioOrigen, strCodMonedaParEvaluacionOrigen, strCodMonedaParPorDefectoOrigen)

                strIndUltimoMovimiento = Valor_Caracter
            
            End If
            
            If intContador = 2 Then
                'Cuenta Movimiento
                strCodCuentaMovimiento = strCodCuentaFondo
                strCodFileMovimiento = strCodFileCuenta
                strCodAnaliticaMovimiento = strCodAnaliticaCuenta
                strCodMonedaMovimiento = strCodMonedaCuenta
                strCodBancoMovimiento = strCodBancoCuenta
                'Cuenta Contraparte
                strCodContracuenta = strCodCuentaActivo
                strCodFileContracuenta = strCodFile
                strCodAnaliticaContracuenta = strCodAnalitica
                'Movimiento
                If strCodTipoMovimiento = Codigo_Movimiento_Deposito Then
                    curMontoContable = CCur(txtMontoMovimientoDestino.Text) * -1
                Else
                    curMontoContable = CCur(txtMontoMovimientoDestino.Text)
                End If
            
                dblTipoCambio = dblTipoCambioDestino 'ObtenerTipoCambioArbitraje(dblTipoCambioDestino, strCodMonedaParEvaluacionDestino, strCodMonedaParPorDefectoDestino)
            
                strIndUltimoMovimiento = Valor_Indicador
            
            End If
            
            If curMontoContable > 0 Then
                strDescripMovimiento = "DEPOSITO EN CUENTA"
                strIndDebeHaber = "D"
            Else
                strDescripMovimiento = "RETIRO DE CUENTA"
                strIndDebeHaber = "H"
            End If
            
            dblValorTipoCambio = 1
            
            If strCodMonedaMovimiento <> Codigo_Moneda_Local Then
                curMontoME = curMontoContable
                curMontoContable = Round(curMontoContable * dblTipoCambio, 2)
                curMontoMN = 0
                dblValorTipoCambio = dblTipoCambio
            Else
                curMontoME = 0
                curMontoMN = curMontoContable
            End If
            
            curMontoContableAcum = curMontoContableAcum + curMontoContable
       
            If curMontoContableAcum <> 0 And curMontoContableAcum <> curMontoContable Then
                curMontoContable = curMontoContable - curMontoContableAcum
            End If
         
            '-----
            If Trim(txtDescripcionTransf.Text) <> Valor_Caracter Then
                strDescripMovimiento = Trim(txtDescripcionTransf.Text)
            End If
            '-----
            
            dblValorTipoCambio = dblTipoCambio
            strTipoDocumento = "13"
            strNumDocumento = Trim(txtNroTransaccion.Text)
            strTipoPersonaContraparte = strTipoContraparte
            strCodPersonaContraparte = strCodContraparte
            strIndContracuenta = Valor_Indicador
            
            '*** Movimiento ***
            .CommandText = "{ call up_ACAdicAsientoContableDetalle('" & _
                strNumAsiento & "','" & strCodFondo & "','" & _
                gstrCodAdministradora & "'," & _
                intContador & ",'" & _
                strFechaAsiento & "','" & _
                gstrPeriodoActual & "','" & gstrMesActual & "','" & _
                strDescripMovimiento & "','" & strIndDebeHaber & "','" & _
                strCodCuentaMovimiento & "','" & strCodMonedaMovimiento & "'," & _
                CDec(curMontoMN) & "," & _
                CDec(curMontoME) & "," & _
                CDec(curMontoContable) & "," & _
                dblValorTipoCambio & ",'" & _
                strCodFileMovimiento & "','" & strCodAnaliticaMovimiento & "','" & _
                strTipoDocumento & "','" & _
                strNumDocumento & "','" & _
                strTipoPersonaContraparte & "','" & _
                strCodPersonaContraparte & "','" & _
                strIndContracuenta & "','" & _
                strCodContracuenta & "','" & _
                strCodFileContracuenta & "','" & _
                strCodAnaliticaContracuenta & "','" & _
                strIndUltimoMovimiento & "','','','" & _
                strTipoCambioReemplazoXML & "') }"
                
            adoConn.Execute .CommandText
               
            strOrdenCobroPago = "0000000000"
                          
            'Verificar la moneda
            If strCodMonedaMovimiento <> Codigo_Moneda_Local Then
                dblMontoPago = curMontoME
            Else
                dblMontoPago = curMontoMN
            End If
            
            If dblMontoPago > 0 Then
                strTipoMovimiento = "E"
            Else
                strTipoMovimiento = "S"
            End If
                         
            'Graba en tabla de BancoCuentaMovimiento
            .CommandText = "{ call up_ACAdicBancoCuentaMovimiento('" & strCodFondo & "','" & gstrCodAdministradora & "','" & strBancoCuentaMovimiento & "'," & _
                intContador & ",'" & strOrdenCobroPago & "','" & strNumAsiento & "'," & intContador & ",'" & strFechaAsiento & "','" & Trim(frmMainMdi.Tag) & "','" & strDescripMovimiento & "','" & _
                strTipoMovimiento & "','03','" & strCodBancoMovimiento & "','" & Trim(txtNroTransaccion.Text) & "','" & strCodCuentaMovimiento & "'," & _
                dblMontoPago & ",'" & strCodFileMovimiento & "','" & strCodAnaliticaMovimiento & "','" & strCodMonedaMovimiento & "'," & curMontoContable & ",'" & Codigo_Moneda_Local & "'," & dblTipoCambio & ",'" & gstrCodClaseTipoCambioOperacionFondo & "','" & _
                strCodContraparte & "','" & strTipoContraparte & "','" & _
                strCodCuentaContraparte & "','" & strCodFileContraparte & "','" & strCodAnaliticaContraparte & "') }"
            adoConn.Execute .CommandText
           
        Next
            
            
        '-- Verifica y ajusta posibles descuadres
'        .CommandText = "{ call up_ACProcAsientoContableAjuste('" & _
'                strCodFondo & "','" & _
'                gstrCodAdministradora & "','" & _
'                strNumAsiento & "') }"
'        adoConn.Execute .CommandText
        
        '*** Actualizar el número del parámetro **
        .CommandText = "{ call up_ACActUltNumero('"
        .CommandText = .CommandText & strCodFondo & "','"
        .CommandText = .CommandText & gstrCodAdministradora & "','"
        .CommandText = .CommandText & Valor_NumOpeCajaBancos & "','"
        .CommandText = .CommandText & strBancoCuentaMovimiento & "') }"
        adoConn.Execute .CommandText
        
        '*** Actualizar el número del parámetro **
        .CommandText = "{ call up_ACActUltNumero('"
        .CommandText = .CommandText & strCodFondo & "','"
        .CommandText = .CommandText & gstrCodAdministradora & "','"
        .CommandText = .CommandText & Valor_NumComprobante & "','"
        .CommandText = .CommandText & strNumAsiento & "') }"
        adoConn.Execute .CommandText
        
        adoComm.CommandText = "COMMIT TRAN OrdTrans"
        adoComm.Execute
        
        blnClientTran = False
        
        MsgBox "Transferencia procesada exitosamente", vbExclamation, Me.Caption
                
        frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
        strEstado = Reg_Consulta
        cmdOpcion.Visible = True
        With tabCuenta
            .TabEnabled(0) = True
            .Tab = 0
        End With
        
    End With
    
    Exit Sub
    
CtrlError:
    If adoConn.Errors.Count > 0 Then
        For Each adoError In adoConn.Errors
            strErrMsg = strErrMsg & adoError.Description & " (" & adoError.NativeError & ") "
        Next
        
        If blnClientTran Then
            adoComm.CommandText = "ROLLBACK TRAN OrdTrans"
            adoComm.Execute
        End If

        Me.MousePointer = vbDefault
        MsgBox strErrMsg, vbCritical + vbOKOnly, Me.Caption
    Else
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
    End If
    
    
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
    
    tabCuenta.TabVisible(2) = False
        
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
    
    '*** Tipo Cuenta ***
    strSQL = "SELECT CodParametro CODIGO, DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='CTAFON' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboTipoCuenta, arrTipoCuenta(), Sel_Defecto
    
    '*** Moneda ***
    strSQL = "{ call up_ACSelDatos(2) }"
    CargarControlLista strSQL, cboMoneda, arrMoneda(), Sel_Defecto
    
    '*** Banco ***
    strSQL = "{ call up_ACSelDatos(22) }"
    CargarControlLista strSQL, cboBanco, arrBanco(), Sel_Defecto

    '*** Cuenta Activo ***
    strSQL = "{ call up_TELstCuentaActivoBanco('" & gstrCodAdministradora & "','" & strCodTipoCuenta & "','" & strCodMoneda & "') }"
    CargarControlLista strSQL, cboCuentaActivo, arrCuentaActivo(), Sel_Defecto
    
    '*** Tipo de Cálculo de Intereses ***
    strSQL = "SELECT CodParametro CODIGO, DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='CALINT' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboCalculo, arrCalculo(), ""
    
    If cboCalculo.ListCount > 0 Then cboCalculo.ListIndex = 0
    
    '*** Tipo Remunerada ***
    strSQL = "SELECT CodParametro CODIGO, DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPREM' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboTipoRemunerada, arrTipoRemunerada(), Sel_Defecto
        
    '*** Por Monto ***
    strSQL = "SELECT CodParametro CODIGO, DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='REMMON' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboPorMonto, arrPorMonto(), Sel_Defecto
    
    '*** Tipo Movimiento ***
    strSQL = "SELECT CodParametro CODIGO, DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='MOVCTA' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboTipoMovimiento, arrTipoMovimiento(), Sel_Defecto
        
End Sub

Private Sub InicializarValores()

    '*** Valores Iniciales ***
    strEstado = Reg_Defecto
    tabCuenta.Tab = 0
    
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgConsulta.Columns(0).Width = tdgConsulta.Width * 0.01 * 5
    tdgConsulta.Columns(1).Width = tdgConsulta.Width * 0.01 * 10
    tdgConsulta.Columns(3).Width = tdgConsulta.Width * 0.01 * 16
    tdgConsulta.Columns(5).Width = tdgConsulta.Width * 0.01 * 9
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

    Call OcultarReportes
    Set frmCuentaFondo = Nothing
    
End Sub


Public Sub Buscar()

    Dim strSQL As String
    
    strSQL = "SELECT Tabla2.DescripParametro TipoCuenta, DescripPersona CodBanco,NumCuenta, Tabla1.CodSigno, CodFile, CodAnalitica, DescripCuenta," & _
        "ResponsableCuenta, CodCuentaActivo, FechaApertura , CodTasa, IndRemunerada,TipoRemunerada, MontoBaseRemunerada,  TipoMontoRemunerada, BancoCuenta.CodMoneda " & _
        "FROM BancoCuenta JOIN Moneda Tabla1 ON(Tabla1.CodMoneda=BancoCuenta.CodMoneda) " & _
        "JOIN AuxiliarParametro Tabla2 ON(Tabla2.CodParametro=BancoCuenta.TipoCuenta AND Tabla2.CodTipoParametro='CTAFON') " & _
        "JOIN InstitucionPersona ON (InstitucionPersona.CodPersona=BancoCuenta.CodBanco AND InstitucionPersona.TipoPersona='02') " & _
        "WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND BancoCuenta.IndVigente='X' " & _
        "ORDER BY TipoCuenta"
                        
    Set adoConsulta = New ADODB.Recordset
                        
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

'
Private Sub lblSaldoCuentaTransferencia_Change()

    Call FormatoMillarEtiqueta(lblSaldoCuentaTransferencia, Decimales_Monto)
    
End Sub

Private Sub lblSaldoDisponible_Change()

    Call FormatoMillarEtiqueta(lblSaldoDisponible, Decimales_Monto)
    
End Sub

Private Sub tabCuenta_Click(PreviousTab As Integer)

    Select Case tabCuenta.Tab
        Case 1
            If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vModify)
            If strEstado = Reg_Defecto Then tabCuenta.Tab = 0
        Case 2
            tabCuenta.Tab = PreviousTab
    End Select
    
End Sub

Private Sub tdgConsulta_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    On Error GoTo Error1            '/**/ HMC Habilitamos la rutina de Errores.

    If Not IsNull(LastRow) Then     '/**/
        If tdgConsulta.Row < 0 Then Exit Sub

        If CStr(LastRow) = Valor_Caracter Then Exit Sub
    
        tdgConsulta.ToolTipText = Valor_Caracter
        If LastRow >= 1 Then
            Call SaldoCuentaDineraria(tdgConsulta.Columns("CodCuentaActivo").Value, tdgConsulta.Columns("CodFile").Value, tdgConsulta.Columns("CodAnalitica").Value, tdgConsulta.Columns("CodMoneda").Value)
        Else
            tdgConsulta.ToolTipText = Valor_Caracter
        End If
    
    End If                           '/**/
    Exit Sub

Error1:
    MsgBox DescripcionError & vbNewLine & DescripcionTecnica & err.Description, vbExclamation, TituloError ' Mostrar Error
    
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

Private Sub txtMontoContableDestino_Change()

Call FormatoCajaTexto(txtMontoContableDestino, Decimales_Monto)

End Sub

Private Sub txtMontoContableOrigen_Change()

 Call FormatoCajaTexto(txtMontoContableOrigen, Decimales_Monto)

End Sub


Private Sub txtMontoMovimientoDestino_Change()

    If txtMontoMovimientoDestino.Tag = "0" Then Call CalculoTotal(1)

End Sub

Private Sub txtMontoMovimientoDestino_GotFocus()

    txtMontoMovimientoDestino.Tag = "0"

End Sub

Private Sub txtMontoMovimientoDestino_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        Call txtMontoMovimientoDestino_Change 'CalculoTotal(1)
    End If
    
End Sub



Private Sub txtMontoMovimientoOrigen_Change()

    If txtMontoMovimientoOrigen.Tag = "0" Then Call CalculoTotal(0)

End Sub

Private Sub txtMontoMovimientoOrigen_GotFocus()

    txtMontoMovimientoOrigen.Tag = "0"
        
End Sub

Private Sub txtMontoMovimientoOrigen_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        Call txtMontoMovimientoOrigen_Change 'Call CalculoTotal(0)
    End If

End Sub

Private Sub txtMontoRemunerada_Change()

    Call FormatoCajaTexto(txtMontoRemunerada, Decimales_Monto)
    
End Sub

Private Sub txtMontoRemunerada_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtMontoRemunerada, Decimales_Monto)
    
End Sub


Private Sub txtTipoCambio_Change()

    Call CalculoTotal(0)

End Sub

Private Sub txtTipoCambio_GotFocus()

    txtMontoMovimientoOrigen.Tag = "0"

End Sub

Private Sub txtTipoCambio_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        Call txtTipoCambio_Change 'CalculoTotal(0)
    End If
    
End Sub

Private Sub CalculoTotal(index As Integer)

    Dim curMonTotal As Currency
    Dim dblValorTC As Double

    'If Not IsNumeric(txtComisionAgente(Index).Text) And Not IsNumeric(txtComisionBolsa(Index).Text) And Not IsNumeric(txtComisionConasev(Index).Text) And Not IsNumeric(txtComisionCavali(Index).Text) And Not IsNumeric(txtComisionFondo(Index).Text) And Not IsNumeric(txtComisionFondoG(Index).Text) Then Exit Sub
    
    If index = 0 Then ' actualiza desde el origen al destino
'        txtMontoMovimientoOrigen.Tag = "0"
'        txtMontoMovimientoDestino.Tag = ""
        strCodMonedaParEvaluacion = strCodMoneda & strCodMonedaCuenta
        If chkModificaTC.Value = vbUnchecked Then
            dblValorTC = 0
            If CDbl(txtMontoMovimientoDestino.Text) <> 0 Then
                strCodMonedaParEvaluacion = strCodMonedaCuenta & strCodMoneda
                dblValorTC = txtMontoMovimientoOrigen.Value / txtMontoMovimientoDestino.Value
                dblValorTC = ObtenerTipoCambioArbitraje(dblValorTC, strCodMonedaParEvaluacion, strCodMonedaParPorDefecto)
            End If
            txtTipoCambio.Text = CStr(dblValorTC)
        Else
            If CDbl(txtTipoCambio.Text) <> 0 Then
                curMonTotal = Round(ObtenerMontoArbitraje(txtMontoMovimientoOrigen.Value, CDbl(txtTipoCambio.Text), strCodMonedaParEvaluacion, strCodMonedaParPorDefecto), 2)
                txtMontoMovimientoDestino.Text = CStr(curMonTotal)
            Else
                txtMontoMovimientoDestino.Text = "0"
            End If
        End If
    End If
    
    If index = 1 Then ' actualiza desde el destino al origen
'        txtMontoMovimientoOrigen.Tag = ""
'        txtMontoMovimientoDestino.Tag = "0"
        strCodMonedaParEvaluacion = strCodMonedaCuenta & strCodMoneda
        If chkModificaTC.Value = vbUnchecked Then
            dblValorTC = 0
            If CDbl(txtMontoMovimientoDestino.Text) <> 0 Then
                strCodMonedaParEvaluacion = strCodMoneda & strCodMonedaCuenta
                dblValorTC = txtMontoMovimientoOrigen.Value / txtMontoMovimientoDestino.Value
                dblValorTC = ObtenerTipoCambioArbitraje(dblValorTC, strCodMonedaParEvaluacion, strCodMonedaParPorDefecto)
            End If
            txtTipoCambio.Text = CStr(dblValorTC)
        Else
            If CDbl(txtTipoCambio.Text) <> 0 Then
                curMonTotal = Round(ObtenerMontoArbitraje(txtMontoMovimientoDestino.Value, CDbl(txtTipoCambio.Text), strCodMonedaParEvaluacion, strCodMonedaParPorDefecto), 2)
                txtMontoMovimientoOrigen.Text = CStr(curMonTotal)
            Else
                txtMontoMovimientoOrigen.Text = "0"
            End If
        End If
    End If
        
        
End Sub

