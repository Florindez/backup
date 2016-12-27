VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{830D5A06-9B70-4F8C-98B6-7A19C4E7760B}#1.0#0"; "TAMControls.ocx"
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmMovimientoBancario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimiento Bancario"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8865
   ScaleWidth      =   11370
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
      Left            =   8010
      Picture         =   "frmMovimientoBancario.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   8070
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
      Left            =   6600
      Picture         =   "frmMovimientoBancario.frx":0671
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   8070
      Width           =   1200
   End
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   9420
      TabIndex        =   68
      Top             =   8070
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
      Left            =   660
      TabIndex        =   67
      Top             =   8070
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
   Begin TAMControls.ucBotonEdicion cmdSalir2 
      Height          =   390
      Left            =   8820
      TabIndex        =   2
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
      Left            =   600
      TabIndex        =   1
      Top             =   9840
      Visible         =   0   'False
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   688
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
      Caption3        =   "Anular"
      Tag3            =   "4"
      Visible3        =   0   'False
      ToolTipText3    =   "Anular"
      UserControlHeight=   390
      UserControlWidth=   5700
   End
   Begin TabDlg.SSTab tabPagos 
      Height          =   7905
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   13944
      _Version        =   393216
      Style           =   1
      Tabs            =   2
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
      TabPicture(0)   =   "frmMovimientoBancario.frx":0C96
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDescrip(20)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDescrip(19)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblDescrip(17)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblDescrip(16)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Line1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "tdgConsulta"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fraCriterio"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtTotalSeleccionadoME"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtTotalME"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtTotalSeleccionado"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtTotal"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "dxDBGrid1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "frmMovimientoBancario.frx":0CB2
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraDatos"
      Tab(1).Control(1)=   "cmdAccion"
      Tab(1).ControlCount=   2
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   -67140
         TabIndex        =   66
         Top             =   6960
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
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
         Height          =   30
         Left            =   7170
         OleObjectBlob   =   "frmMovimientoBancario.frx":0CCE
         TabIndex        =   63
         Top             =   5670
         Width           =   30
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   7740
         TabIndex        =   19
         Top             =   7050
         Width           =   2000
      End
      Begin VB.TextBox txtTotalSeleccionado 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   3060
         TabIndex        =   18
         Top             =   7050
         Width           =   2000
      End
      Begin VB.TextBox txtTotalME 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   7740
         TabIndex        =   17
         Top             =   7440
         Width           =   2000
      End
      Begin VB.TextBox txtTotalSeleccionadoME 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   3060
         TabIndex        =   16
         Top             =   7440
         Width           =   2000
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
         Height          =   2385
         Left            =   360
         TabIndex        =   5
         Top             =   480
         Width           =   10455
         Begin VB.ComboBox cboEstadoOperacionBusqueda 
            Height          =   315
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   64
            Top             =   1860
            Width           =   2745
         End
         Begin VB.ComboBox cboCuentasBancarias 
            Height          =   315
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   1380
            Width           =   4530
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
            Left            =   8760
            Picture         =   "frmMovimientoBancario.frx":1949
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   1380
            Width           =   1200
         End
         Begin VB.ComboBox cboFondo 
            Height          =   315
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   510
            Width           =   6465
         End
         Begin MSComCtl2.DTPicker dtpFechaMovimBCDesde 
            Height          =   315
            Left            =   2160
            TabIndex        =   10
            Top             =   930
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
         Begin MSComCtl2.DTPicker dtpFechaMovimBCHasta 
            Height          =   315
            Left            =   5100
            TabIndex        =   11
            Top             =   930
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
            Index           =   22
            Left            =   360
            TabIndex        =   65
            Top             =   1890
            Width           =   1515
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
            Index           =   5
            Left            =   360
            TabIndex        =   15
            Top             =   1470
            Width           =   615
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
            Left            =   4170
            TabIndex        =   13
            Top             =   990
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
            Left            =   360
            TabIndex        =   12
            Top             =   1020
            Width           =   555
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
            TabIndex        =   6
            Top             =   540
            Width           =   540
         End
      End
      Begin VB.Frame fraDatos 
         Caption         =   "Movimientos"
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
         Height          =   6495
         Left            =   -74580
         TabIndex        =   4
         Top             =   390
         Width           =   10455
         Begin VB.TextBox txtTipoCambioPago 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   2340
            TabIndex        =   60
            Top             =   6930
            Visible         =   0   'False
            Width           =   1755
         End
         Begin VB.ComboBox cboCreditoFiscal 
            Enabled         =   0   'False
            Height          =   315
            Left            =   7470
            Style           =   2  'Dropdown List
            TabIndex        =   58
            Top             =   4680
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.ComboBox cboModalidadPago 
            Enabled         =   0   'False
            Height          =   315
            Left            =   5010
            Style           =   2  'Dropdown List
            TabIndex        =   56
            Top             =   7290
            Visible         =   0   'False
            Width           =   1785
         End
         Begin VB.ComboBox cboEstado 
            Enabled         =   0   'False
            Height          =   315
            Left            =   7440
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   5700
            Visible         =   0   'False
            Width           =   2115
         End
         Begin VB.ComboBox cboTipoGasto 
            Enabled         =   0   'False
            Height          =   315
            Left            =   5040
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   7710
            Visible         =   0   'False
            Width           =   1785
         End
         Begin VB.TextBox txtMontoPago 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1890
            MaxLength       =   40
            TabIndex        =   47
            Text            =   " "
            Top             =   5280
            Width           =   1800
         End
         Begin VB.TextBox txtNroVoucher 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1890
            MaxLength       =   12
            TabIndex        =   46
            Text            =   " "
            Top             =   5880
            Width           =   1830
         End
         Begin VB.ComboBox cboCuentas 
            Height          =   315
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   4170
            Width           =   4350
         End
         Begin VB.ComboBox cboMoneda 
            Height          =   315
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   3690
            Width           =   2565
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
            Left            =   9690
            TabIndex        =   32
            ToolTipText     =   "Buscar Proveedor"
            Top             =   2610
            Width           =   375
         End
         Begin VB.ComboBox cboFormaPago 
            Height          =   315
            Left            =   7230
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   2040
            Width           =   2835
         End
         Begin VB.ComboBox cboTipoMov 
            Height          =   315
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   2040
            Width           =   3375
         End
         Begin VB.TextBox txtDescripIngreso 
            Height          =   315
            Left            =   1890
            MaxLength       =   60
            TabIndex        =   25
            Top             =   1470
            Width           =   8115
         End
         Begin VB.ComboBox cboGasto 
            Height          =   315
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   960
            Width           =   4395
         End
         Begin MSComCtl2.DTPicker dtpFechaActual 
            Height          =   315
            Left            =   1890
            TabIndex        =   38
            Top             =   4710
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   206766081
            CurrentDate     =   38949
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Cambio"
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
            Index           =   21
            Left            =   900
            TabIndex        =   61
            Top             =   6990
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Crédito Fiscal"
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
            Left            =   3210
            TabIndex        =   59
            Top             =   7020
            Visible         =   0   'False
            Width           =   1605
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Modalidad Pago"
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
            Index           =   31
            Left            =   3540
            TabIndex        =   57
            Top             =   7410
            Visible         =   0   'False
            Width           =   1380
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
            Index           =   15
            Left            =   6360
            TabIndex        =   55
            Top             =   5730
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Gasto"
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
            Index           =   27
            Left            =   3960
            TabIndex        =   53
            Top             =   7800
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.Label lblCodProveedor 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   7440
            TabIndex        =   51
            Top             =   3120
            Width           =   2655
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
            Index           =   12
            Left            =   360
            TabIndex        =   50
            Top             =   5280
            Width           =   540
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
            Index           =   10
            Left            =   360
            TabIndex        =   49
            Top             =   5880
            Width           =   1140
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
            Left            =   3780
            TabIndex        =   48
            Top             =   5280
            Width           =   465
         End
         Begin VB.Label lblSaldoCuenta 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7470
            TabIndex        =   45
            Top             =   4170
            Width           =   2625
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Disponible"
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
            Left            =   6360
            TabIndex        =   44
            Top             =   4170
            Width           =   900
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
            Index           =   4
            Left            =   360
            TabIndex        =   43
            Top             =   4170
            Width           =   615
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
            TabIndex        =   41
            Top             =   3690
            Width           =   690
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
            Left            =   360
            TabIndex        =   39
            Top             =   4740
            Width           =   1305
         End
         Begin VB.Label lblProveedor 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1890
            TabIndex        =   37
            Top             =   2610
            Width           =   7710
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
            Index           =   13
            Left            =   360
            TabIndex        =   36
            Top             =   2610
            Width           =   555
         End
         Begin VB.Label lblTipoDocID 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1890
            TabIndex        =   35
            Top             =   3120
            Width           =   2535
         End
         Begin VB.Label lblNumDocID 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   4590
            TabIndex        =   34
            Top             =   3120
            Width           =   2655
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Documento ID"
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
            Left            =   360
            TabIndex        =   33
            Top             =   3120
            Width           =   1230
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
            Index           =   11
            Left            =   5700
            TabIndex        =   31
            Top             =   2040
            Width           =   1365
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
            Left            =   360
            TabIndex        =   29
            Top             =   2040
            Width           =   1410
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Detalle"
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
            TabIndex        =   27
            Top             =   1470
            Width           =   615
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Concepto"
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
            TabIndex        =   26
            Top             =   960
            Width           =   825
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
            TabIndex        =   8
            Top             =   450
            Width           =   540
         End
         Begin VB.Label lblDescripFondo 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1890
            TabIndex        =   7
            Top             =   450
            Width           =   8085
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmMovimientoBancario.frx":1EB1
         Height          =   3765
         Left            =   360
         OleObjectBlob   =   "frmMovimientoBancario.frx":1ECB
         TabIndex        =   62
         Top             =   2970
         Width           =   10455
      End
      Begin VB.Line Line1 
         X1              =   300
         X2              =   10740
         Y1              =   6930
         Y2              =   6930
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Monto Total MN"
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
         Index           =   16
         Left            =   5910
         TabIndex        =   23
         Top             =   7050
         Width           =   1455
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Monto Seleccionado MN"
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
         Left            =   510
         TabIndex        =   22
         Top             =   7050
         Width           =   2295
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Monto Total ME"
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
         Left            =   5910
         TabIndex        =   21
         Top             =   7440
         Width           =   1455
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Monto Seleccionado ME"
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
         Index           =   20
         Left            =   510
         TabIndex        =   20
         Top             =   7440
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmMovimientoBancario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strCodFondo         As String, strCodParticipe      As String
Dim strEstado           As String, strSQL               As String
Dim strTipMov           As String
Dim strFormPago         As String
Dim curMontoEmitido     As Currency, strNroVocuher      As String
Dim arrMoneda()         As String, strCodMoneda         As String, strSignoMoneda   As String, strCodSignoMoneda    As String
Dim arrCuenta()         As String, strCodCuenta         As String, strCodGasto      As String, strCodTipoGasto      As String
Dim arrCuentas()        As String
Dim arrCuentasBancarias() As String
Dim arrCuentasControl() As String
Dim arrGasto()          As String
Dim arrTipoGasto()      As String
Dim arrEstadoGasto()    As String
Dim arrModalidadPago()  As String
Dim arrCreditoFiscal()  As String
Dim arrTipMov()         As String, strEstadoOperacion   As String, strCodPeriodoTasa   As String
Dim strCodTipoValor     As String, strCodTipoTasa       As String, strCodBaseCalculo   As String
Dim strCodFileGasto     As String, strCodFileBanco  As String, strCodBanco  As String, strCodModalidadPago         As String
Dim strEstadoGasto      As String, strCodAnaliticaGasto As String, strCodAnaliticaBanco As String
Dim strCodCreditoFiscal As String, strCodCuentaBancaria As String, strCodMonedaOrden As String
Dim intSecuencialGasto  As Integer, dblTipoCambioOrden  As Double, strIndSeleccionMultiple  As String
Dim adoConsulta         As ADODB.Recordset
Dim adoRegistroAux      As ADODB.Recordset
Dim arrEstadoOperacionBusqueda()  As String
Dim arrFormaPago() As String
Dim strEstadoOperacionBusqueda As String
Dim strCodParticipeBusqueda As String
Dim strNumOperacion As String
Dim strCodProveedor As String
Dim strCodTipoProveedor As String

Private Sub cboCreditoFiscal_Click()

    strCodCreditoFiscal = Valor_Caracter
    If cboCreditoFiscal.ListIndex < 0 Then Exit Sub
    
    strCodCreditoFiscal = arrCreditoFiscal(cboCreditoFiscal.ListIndex)
    
End Sub

Private Sub cboCuentas_Click()
    
    Dim strFecha As String, strFechaMas1Dia As String
    
    strCodFileBanco = Valor_Caracter: strCodAnaliticaBanco = Valor_Caracter
    strCodBanco = Valor_Caracter: strCodCuenta = Valor_Caracter
    If cboCuentas.ListIndex < 0 Then Exit Sub
   
   '*******BMM NUEVO
   
    strCodFileBanco = Left(Trim(arrCuentas(cboCuentas.ListIndex)), 3)
    strCodAnaliticaBanco = Mid(Trim(arrCuentas(cboCuentas.ListIndex)), 4, 8)
    strCodBanco = Mid(Trim(arrCuentas(cboCuentas.ListIndex)), 12, 8)
    strCodCuenta = Trim(Right(arrCuentas(cboCuentas.ListIndex), 10))
   
'    strCodFileBanco = Left(Trim(arrCuentas(cboCuentas.ListIndex)), 3)
'    strCodAnaliticaBanco = Mid(Trim(arrCuentas(cboCuentas.ListIndex)), 4, 8)
'    strCodBanco = Mid(Trim(arrCuentas(cboCuentas.ListIndex)), 12, 8)
'    strCodCuenta = Trim(Right(arrCuentasBancarias(cboCuentas.ListIndex), 10))
   
   '*******
   
'    lblMonedaCuentaLiquidacion(0).Caption = ObtenerCodSignoMoneda(strCodMonedaPago)
'    lblMonedaCuentaLiquidacion(1).Caption = ObtenerCodSignoMoneda(strCodMonedaPago)
   
'    If Trim(txtTipoCambio.Text) = "" Then txtTipoCambio.Text = 0#
'
'    lblMovCuenta.Caption = Round(ObtenerMontoArbitraje(CDbl(txtMonto.Text), CDbl(txtTipoCambio.Text), strCodMonedaParEvaluacion, strCodMonedaParPorDefecto), 2)
'
'    If cboCuentas.ListIndex <> 0 Then
'        If txtTipoCambio.Visible = True Then
'            lblMovCuenta.Caption = Round(ObtenerMontoArbitraje(CDbl(txtMonto.Text), CDbl(txtTipoCambio.Text), strCodMonedaParEvaluacion, strCodMonedaParPorDefecto), 2)
'        Else
'            lblMovCuenta.Caption = CDbl(txtMonto.Text)
'        End If
'    End If
    
     '** BMM NUEVO
    
   strFecha = gstrFechaActual  'Convertyyyymmdd(dtpFechaContable.Value)
   strFechaMas1Dia = Convertyyyymmdd(DateAdd("d", 1, gdatFechaActual))  'dtpFechaContable.Value))
    
   lblSaldoCuenta.Caption = "0"
    
   lblSaldoCuenta.Caption = ObtenerSaldoFinalCuenta(strCodFondo, gstrCodAdministradora, strCodFileBanco, _
   strCodAnaliticaBanco, strFecha, strFechaMas1Dia, strCodCuenta, strCodMoneda)
    
    'Call ObtenerSaldos
    '************

End Sub

Private Sub ObtenerSaldos()

    Dim adoTemporal As ADODB.Recordset
    Dim strFecha    As String, strFechaMas1Dia  As String
    
    strFecha = gstrFechaActual  'Convertyyyymmdd(dtpFechaContable.Value)
    strFechaMas1Dia = Convertyyyymmdd(DateAdd("d", 1, gdatFechaActual))  'dtpFechaContable.Value))
    
    Set adoTemporal = New ADODB.Recordset
    With adoComm
        .CommandText = "{ call up_ACObtenerSaldoCuentaContable('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
            strCodFileBanco & "','" & strCodAnaliticaBanco & "','" & strFecha & "','" & strFechaMas1Dia & "','" & _
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

'Private Sub cboCuenta_Click()
'    strCodCuenta = Valor_Caracter
'    If cboCuenta.ListIndex < 0 Then Exit Sub
'    strCodCuenta = Trim(arrCuenta(cboCuenta.ListIndex))
'
'End Sub

Private Sub cboCuentasBancarias_Click()

    If cboCuentasBancarias.ListIndex < 0 Then Exit Sub
   
    strCodCuentaBancaria = Trim(Right(arrCuentasBancarias(cboCuentasBancarias.ListIndex), 10))

    Call Buscar


End Sub

Private Sub cboEstado_Click()

    strEstadoGasto = Valor_Caracter
    If cboEstado.ListIndex < 0 Then Exit Sub
    
    strEstadoGasto = Trim(arrEstadoGasto(cboEstado.ListIndex))

End Sub

Private Sub cboEstadoOperacionBusqueda_Click()

    strEstadoOperacionBusqueda = Valor_Caracter
    If cboEstadoOperacionBusqueda.ListIndex < 0 Then Exit Sub
    strEstadoOperacionBusqueda = Trim(arrEstadoOperacionBusqueda(cboEstadoOperacionBusqueda.ListIndex))

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

End Sub

Private Sub cboGasto_Click()

    strCodGasto = Valor_Caracter: strCodAnaliticaGasto = Valor_Caracter
    If cboGasto.ListIndex <= 0 Then Exit Sub
    
    strCodGasto = Trim(arrGasto(cboGasto.ListIndex))
    
    Dim adoRegistro     As ADODB.Recordset
    
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        .CommandText = "SELECT CodDetalleFile FROM InversionDetalleFile " & _
            "WHERE CodFile='" & strCodFileGasto & "' AND DescripDetalleFile='" & strCodGasto & "'"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            strCodAnaliticaGasto = Format(adoRegistro("CodDetalleFile"), "00000000")
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    

End Sub

Private Sub cboModalidadPago_Click()

    strCodModalidadPago = Valor_Caracter
    If cboModalidadPago.ListIndex < 0 Then Exit Sub
    
    strCodModalidadPago = Trim(arrModalidadPago(cboModalidadPago.ListIndex))

End Sub

Private Sub cboMoneda_Click()
    strCodMoneda = Valor_Caracter: strSignoMoneda = Valor_Caracter
    If cboMoneda.ListIndex < 0 Then Exit Sub
    
    strCodMoneda = Trim(arrMoneda(cboMoneda.ListIndex))
    
    strSignoMoneda = ObtenerSignoMoneda(strCodMoneda)
    strCodSignoMoneda = ObtenerCodSignoMoneda(strCodMoneda)
    
    lblSignoMoneda.Caption = strCodSignoMoneda
  
    Call CargarCuentas
'    If strCodMoneda <> Valor_Caracter And strTipMov = "01" Then
'        '*** Cuentas ***
'        strSQL = "SELECT CodCuenta CODIGO, DescripCuenta DESCRIP FROM PlanContable  " & _
'                "WHERE CodAdministradora='" & gstrCodAdministradora & "' AND CodCuenta LIKE '46911[1-2]%' AND IndMovimiento='X' AND CodMoneda ='" & strCodMoneda & "'"
'        CargarControlLista strSQL, cboCuenta, arrCuenta(), Valor_Caracter
'
'        If cboCuenta.ListCount > 0 Then cboCuenta.ListIndex = 0
'    End If
        
    
End Sub


Private Sub CargarCuentasBancarias()

    Dim strSQL As String
        
    strSQL = "SELECT (CodFile + CodAnalitica + CodBanco + CodCuentaActivo) CODIGO,(RTRIM(DescripCuenta) + SPACE(1) + NumCuenta) DESCRIP FROM BancoCuenta " & _
             "WHERE IndVigente='X' AND " & _
             "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
    
'    strSQL = "SELECT (CodFile + CodAnalitica + CodBanco + CodCuentaActivo) CODIGO,(RTRIM(DescripCuenta) + SPACE(1) + NumCuenta) DESCRIP FROM BancoCuenta " & _
'             "WHERE CodMoneda='" & strCodMoneda & "' AND IndVigente='X' AND " & _
'             "CodFondo='" & gstrCodFondoContable & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
    
    CargarControlLista strSQL, cboCuentasBancarias, arrCuentasBancarias(), Sel_Todos
    If cboCuentasBancarias.ListCount > 0 Then cboCuentasBancarias.ListIndex = 0
            
End Sub
Private Sub CargarCuentas()

    Dim strSQL As String
        
    
    strSQL = "SELECT (CodFile + CodAnalitica + CodBanco + CodCuentaActivo) CODIGO,(RTRIM(DescripCuenta) + SPACE(1) + NumCuenta) DESCRIP FROM BancoCuenta " & _
             "WHERE CodMoneda='" & strCodMoneda & "' AND IndVigente='X' AND " & _
             "CodFondo='" & gstrCodFondoContable & "' AND CodAdministradora='" & gstrCodAdministradora & "' "
    
    CargarControlLista strSQL, cboCuentas, arrCuentas(), Sel_Defecto
    If cboCuentas.ListCount > 0 Then cboCuentas.ListIndex = 0
            
End Sub



Private Sub cboTipoGasto_Click()

    strCodTipoGasto = Valor_Caracter
    If cboTipoGasto.ListIndex < 0 Then Exit Sub
    
    strCodTipoGasto = Trim(arrTipoGasto(cboTipoGasto.ListIndex))

End Sub

Private Sub cboTipoMov_Click()
     
     strTipMov = Valor_Caracter
     If cboTipoMov.ListIndex < 0 Then Exit Sub
     strTipMov = Trim(arrTipMov(cboTipoMov.ListIndex))

End Sub

Private Sub cmdImprimir_Click()
    Call Imprimir
End Sub

Private Sub cmdProcesar_Click()

    Dim intContador                 As Integer
    Dim intRegistro                 As Integer
    Dim strGastoOperacionXML    As String
    Dim objGastoOperacionXML   As DOMDocument60
    Dim strFechaGrabar              As String
    Dim strMsgError                 As String
    Dim strTipoOperacion            As String
    Dim adoRegistro                 As ADODB.Recordset
    Dim adoLiquidacion              As ADODB.Recordset
    
    Dim strSQL                              As String
    Dim strCodProceso                       As String
    Dim strNumCheque                        As String
    Dim strMovimientoFondoLiquidacionXML    As String
    Dim objMovimientoFondoLiquidacionXML    As DOMDocument60
    Dim strNumOrdenCobroPago                As String
        
    On Error GoTo ErrorHandler
        
    If TodoOkBackOffice() Then
        '*** Realizar proceso de contabilización ***
        If MsgBox("Datos correctos. ¿ Procedemos a enviar estas operaciones a Backoffice de Tesoreria?", vbQuestion + vbYesNo, "Observación") = vbNo Then Exit Sub
    
        intContador = tdgConsulta.SelBookmarks.Count - 1
               
        strFechaGrabar = Convertyyyymmdd(dtpFechaActual.Value) & Space(1) & Format(Time, "hh:mm")

               
        Set adoRegistro = New ADODB.Recordset
        
        With adoComm
        
            Set objGastoOperacionXML = Nothing
            strGastoOperacionXML = ""
                      
            For intRegistro = 0 To intContador
                
                adoConsulta.MoveFirst
                
                adoConsulta.Move CLng(tdgConsulta.SelBookmarks(intRegistro) - 1), 0
                                
                tdgConsulta.Refresh

                .CommandText = "{ call up_CNContabilizarRegistroCompra('" & _
                gstrCodFondoContable & "','" & gstrCodAdministradora & "','" & _
                strFechaGrabar & "','" & _
                adoConsulta("NumRegistro") & "') }"
                
                .Execute .CommandText
                
                Set adoLiquidacion = New ADODB.Recordset
                    
                strSQL = "SELECT MF.NumOrdenCobroPago,MF.DescripOrden,RC.CodCuentaBanco,RC.CodFileBanco,RC.CodAnaliticaBanco,BC.CodBanco,RC.CodFormaPago,MF.CodMoneda,MF.MontoOrdenCobroPago,RC.CodMonedaPago,RC.TipoCambioPago,RC.TipoProveedor,RC.CodProveedor " & _
                        "FROM RegistroCompra RC " & _
                        "JOIN MovimientoFondo MF ON (MF.CodFondo = RC.CodFondo AND MF.CodAdministradora = RC.CodAdministradora AND MF.NumOperacion = RC.NumRegistro) " & _
                        "JOIN BancoCuenta BC ON (BC.CodFondo = RC.CodFondo AND BC.CodAdministradora = RC.CodAdministradora AND BC.CodCuentaActivo=RC.CodCuentaBanco AND BC.CodFile=RC.CodFileBanco AND BC.CodAnalitica=RC.CodAnaliticaBanco) " & _
                        "JOIN InstitucionPersona IP ON(IP.CodPersona = MF.CodContraparte AND IP.TipoPersona = MF.TipoContraparte) " & _
                        "WHERE RC.CodFondo = '" & strCodFondo & "' AND RC.CodAdministradora = '" & gstrCodAdministradora & "' AND RC.NumRegistro = Convert(int,'" & adoConsulta("NumRegistro") & "') AND MF.ClaseOperacion= '" & Valor_NumRegistroCompra & "'"
                                    
                With adoLiquidacion
                    .ActiveConnection = gstrConnectConsulta
                    .CursorLocation = adUseClient
                    .CursorType = adOpenStatic
                    .LockType = adLockBatchOptimistic
                    .Open strSQL
                End With
                
                Call ConfiguraRecordsetAuxiliar
                
                adoRegistroAux.AddNew
                adoRegistroAux.Fields("CodFondo") = strCodFondo
                adoRegistroAux.Fields("CodAdministradora") = gstrCodAdministradora
                adoRegistroAux.Fields("NumOrdenCobroPago") = adoLiquidacion("NumOrdenCobroPago")
                
                Call XMLADORecordset(objMovimientoFondoLiquidacionXML, "MovimientoFondoLiquidacion", "Movimiento", adoRegistroAux, strMsgError)
                strMovimientoFondoLiquidacionXML = objMovimientoFondoLiquidacionXML.xml 'CrearXMLDetalle(objTipoCambioReemplazoXML)
                
                .CommandText = "{ call up_ACProcMovimientoFondoLiquidacion('" & _
                            strCodFondo & "','" & gstrCodAdministradora & "','" & strFechaGrabar & "','" & _
                            adoLiquidacion("DescripOrden") & "','" & adoLiquidacion("CodCuentaBanco") & "','" & adoLiquidacion("CodFileBanco") & "','" & _
                            adoLiquidacion("CodAnaliticaBanco") & "','" & adoLiquidacion("CodBanco") & "','" & adoLiquidacion("CodFormaPago") & "','','" & _
                            adoLiquidacion("CodMoneda") & "'," & adoLiquidacion("MontoOrdenCobroPago") & ",'" & adoLiquidacion("CodMonedaPago") & "'," & _
                            adoLiquidacion("MontoOrdenCobroPago") & ",''," & adoLiquidacion("MontoOrdenCobroPago") & ",'01','01'," & _
                            adoLiquidacion("TipoCambioPago") & ",'','T','08','" & strMovimientoFondoLiquidacionXML & "','" & XML_TipoCambioReemplazo & "','','" & _
                            adoLiquidacion("CodProveedor") & "','" & adoLiquidacion("TipoProveedor") & "')}"
                
                adoConn.Execute .CommandText
                
            Next
                        
        End With
        
        Me.MousePointer = vbDefault
        
        MsgBox Mensaje_Proceso_Exitoso, vbExclamation
        
        frmMainMdi.stbMdi.Panels(3).Text = "Acción"
        
       
        With tabPagos
            .TabEnabled(0) = True
            .TabEnabled(1) = False
            .Tab = 0
        End With
        
        Call Buscar
        tdgConsulta.ReBind
        Me.Refresh
    End If

    Exit Sub

ErrorHandler:
    
    If err.Number <> 0 Then
        MsgBox err.Number & " " & err.Description, vbCritical + vbOKOnly, Me.Caption
        Me.MousePointer = vbDefault
    End If


End Sub

Private Sub cmdProveedor_Click()
    Dim adoAuxiliar As ADODB.Recordset
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
        
        frmBus.Caption = " Relación de Proveedores Bancarios"
        '.sSql = "{ call up_ACSelDatos(44) }"
        
        .sSql = "SELECT CodPersona CODIGO,IP.TipoPersona, AP.DescripParametro TipoIdentidad,IP.NumIdentidad, " & _
                "IP.DescripPersona DESCRIP, IP.Direccion1 + IP.Direccion2 Direccion " & _
                "FROM InstitucionPersona IP " & _
                "JOIN AuxiliarParametro AP ON(AP.CodParametro=IP.TipoIdentidad AND AP.CodTipoParametro='TIPIDE') " & _
                "WHERE IP.TipoPersona='" & Codigo_Tipo_Persona_Emisor & "' AND IndBanco='" & Valor_Indicador & "' AND IndVigente='" & Valor_Indicador & "' " & _
                "AND EXISTS (SELECT * FROM BancoCuenta WHERE CodBanco = IP.CodPersona) ORDER BY DescripPersona"
        
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
            lblCodProveedor.Caption = .iParams(1).Valor
        End If
            
 
        
        adoComm.CommandText = "select IP2.CodPersona as CodProveedor " & _
                                "from InstitucionPersona IP1 " & _
                                "join InstitucionPersona IP2 on (IP1.NumIdentidad = IP2.NumIdentidad and IP1.TipoIdentidad = IP2.TipoIdentidad ) " & _
                                "where IP1.TipoPersona = '" & Codigo_Tipo_Persona_Emisor & "' and IP2.TipoPersona = '" & Codigo_Tipo_Persona_Proveedor & "' and IP1.CodPersona = '" & lblCodProveedor.Caption & "' "
        
        Set adoAuxiliar = adoComm.Execute
        
        If adoAuxiliar.EOF Then
            MsgBox "Debe Registrar a la entidad bancaria como Proveedor", vbCritical
            lblProveedor.Caption = Valor_Caracter
            lblTipoDocID.Caption = Valor_Caracter
            lblNumDocID.Caption = Valor_Caracter
            lblCodProveedor.Caption = Valor_Caracter
            Exit Sub
        Else
            strCodProveedor = adoAuxiliar("CodProveedor").Value
        End If
        
        Call CargarCuentas
    
    End With
    
    Set frmBus = Nothing

End Sub

Private Sub cmdReservar_Click()
    Call Reversar
End Sub

Private Sub dtpFechaMovimBCDesde_Change()
    Call Buscar
End Sub

Private Sub dtpFechaMovimBCHasta_Change()
    Call Buscar
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
        Case vModify
            Call Modificar
    End Select
    
End Sub
Public Sub Imprimir()
    Call SubImprimir2(1)
End Sub

Public Sub Eliminar()
    
    Dim adoRegistro As ADODB.Recordset
    Dim strSQL As String, strNumGasto As String
    
    Set adoRegistro = New ADODB.Recordset
    
    If strEstado = Reg_Consulta Or strEstado = Reg_Edicion Then
        strEstadoOperacion = "03"
        
        If Not TodoOkAnular() Then Exit Sub
        
        If MsgBox("Se procederá a anular el Movimiento Bancario Nro. " & tdgConsulta.Columns("NumRegistro").Value & vbNewLine & vbNewLine & "Seguro de continuar ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            
            adoComm.CommandText = "UPDATE RegistroCompra SET Estado='" & strEstadoOperacion & "' " & _
            "WHERE NumRegistro='" & tdgConsulta.Columns("NumRegistro").Value & "' AND " & _
            "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
            "Estado<>'" & strEstadoOperacion & "'"
            
            adoConn.Execute adoComm.CommandText
            
            strSQL = "SELECT NumGasto FROM RegistroCompra " & _
            "WHERE NumRegistro='" & tdgConsulta.Columns("NumRegistro") & "' AND " & _
            "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' "
            
            adoComm.CommandText = strSQL
            
            Set adoRegistro = adoComm.Execute
            
            If Not adoRegistro.EOF Then
                
                Do Until adoRegistro.EOF
                
                    strNumGasto = adoRegistro("NumGasto")
                    adoRegistro.MoveNext
                
                Loop
                
                adoRegistro.Close: Set adoRegistro = Nothing
                
            End If
            
            strSQL = "UPDATE FondoGasto SET IndVigente='' " & _
            "WHERE NumGasto='" & strNumGasto & "' AND " & _
            "CodFondo='" & strCodFondo & "' AND " & _
            "CodAdministradora='" & gstrCodAdministradora & "' AND IndVigente='" & Valor_Indicador & "'"
            
            adoComm.CommandText = strSQL
            
            Call Buscar
            
        End If
    End If
End Sub

Public Sub Reversar()
    
    Dim strFechaGrabar  As String
    Dim strNumRegistro  As String
    Dim motivo          As String
    Dim str_msg, str_pwd         As String
    Dim adoRegistro     As ADODB.Recordset
    If strEstado <> Reg_Edicion Then
        If strEstado <> Reg_Consulta Then Exit Sub
    End If
    
    If Not TodoOkReversar() Then Exit Sub
    
'    If MsgBox("Desea Anular el Movimiento Bancario Nro. " & tdgConsulta.Columns("NumRegistro").Value & " ?", vbYesNo + vbQuestion + vbDefaultButton2) <> vbYes Then
'        Exit Sub
'    End If

    If MsgBox("Desea reversar el Movimiento Bancario Nro. " & tdgConsulta.Columns("NumRegistro").Value & " ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
        
        Me.MousePointer = vbHourglass
        
        If gdatFechaActual > tdgConsulta.Columns(2).Value Then 'cambiar la condicion por la fecha
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
        'If motivo = "" Then Exit Sub

    Else

        Exit Sub

    End If


    On Error GoTo Ctrl_Error
                                        
    With adoComm
        
        .CommandType = adCmdText
        
        strFechaGrabar = Convertyyyymmdd(gdatFechaActual) & Space(1) & Format(Time, "hh:ss")
    
        strNumRegistro = tdgConsulta.Columns("NumRegistro").Value
        
        '*** Cabecera ***
        .CommandText = "{ call up_TEProcAnularOperacionTesoreria('" & _
            strCodFondo & "','" & gstrCodAdministradora & "','" & _
            Valor_NumRegistroCompra & "','" & strNumRegistro & "','" & Space(1) + motivo + Space(1) + "') }"
            
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

Public Sub Salir()

    Unload Me
    
End Sub
Public Sub Cancelar()

    cmdOpcion.Visible = True
    cmdReservar.Visible = True
    With tabPagos
        .TabEnabled(0) = True
        .TabEnabled(1) = False
        .Tab = 0
    End With
    strEstado = Reg_Consulta
    
End Sub
Public Sub Grabar()
                        
    Dim adoRegistro                 As ADODB.Recordset
    Dim strNumComprobante           As String
    Dim strTipoAuxiliar             As String
    Dim strCodAuxiliar              As String
    Dim intNumGasto                 As Long
    Dim numMontoGasto               As Double
    Dim numPorcenGasto              As Double
    Dim strCodAplicacionDevengo     As String
    Dim mensaje As String
    
    On Error GoTo ErrorHandler
    
    If strEstado = Reg_Consulta Then Exit Sub
    
    If Not TodoOk() Then Exit Sub
    
    If strEstado = Reg_Adicion Then
        mensaje = Mensaje_Adicion
    Else
        mensaje = Mensaje_Edicion
    End If
    
    If MsgBox(mensaje, vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption) <> vbYes Then Exit Sub
        
    Me.MousePointer = vbHourglass
     
    strNumComprobante = Trim(txtNroVoucher.Text)
    
    numPorcenGasto = 0
    numMontoGasto = CDbl(txtMontoPago.Text)
    
    strTipoAuxiliar = "02"
    
    strCodAplicacionDevengo = "01"
    
    strCodAuxiliar = Codigo_Tipo_Persona_Proveedor & Trim(lblCodProveedor.Caption)
    
    txtTipoCambioPago.Text = CStr(ObtenerTipoCambioMoneda(Codigo_TipoCambio_Sunat, Codigo_Valor_TipoCambioVenta, dtpFechaActual.Value, Codigo_Moneda_Local, strCodMoneda))
   
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
             
         .CommandText = "{ call up_GNManMovimientoBancario('" & strCodFondo & "','" & _
                gstrCodAdministradora & "'," & intSecuencialGasto & ",'   ','" & Convertyyyymmdd(dtpFechaActual.Value) & "','" & strCodGasto & "','" & _
                strCodFileGasto & "','" & strCodProveedor & "','" & Codigo_Tipo_Persona_Proveedor & "','" & strCodProveedor & "','" & Trim(txtDescripIngreso.Text) & "','" & _
                Convertyyyymmdd(CVDate(Valor_Fecha)) & "','" & Convertyyyymmdd(dtpFechaActual.Value) & "','" & Convertyyyymmdd(dtpFechaActual.Value) & "','" & _
                strCodTipoGasto & "','','" & strEstadoGasto & "'," & gdblTipoCambio & ",'" & strCodMoneda & "','" & strCodTipoValor & "'," & _
                numMontoGasto & "," & numPorcenGasto & ",'" & strCodTipoTasa & "','" & strCodPeriodoTasa & "',0,'" & strCodBaseCalculo & "',0,'" & _
                strCodModalidadPago & "','" & strCodTipoGasto & "',0,'','" & _
                Convertyyyymmdd(dtpFechaActual.Value) & "','00','','01','" & _
                strCodAplicacionDevengo & "','','" & strCodCreditoFiscal & "','','" & strNumComprobante & "', " & CDec(txtMontoPago.Text) & ", '" & _
                strFormPago & "','" & strTipMov & "', '" & strCodCuenta & "', '" & strCodFileBanco & "', '" & strCodAnaliticaBanco & "', " & CDec(txtTipoCambioPago.Text) & ", '" & _
                Convertyyyymmdd(Valor_Fecha) & "','" & Estado_Activo & "', '" & IIf(strEstado = Reg_Adicion, "", tdgConsulta.Columns(1)) & "','" & IIf(strEstado = Reg_Adicion, "I", "U") & "') }"
                
        adoConn.Execute .CommandText
        
        cmdProcesar.Enabled = True
        
    End With
                                
    Me.MousePointer = vbDefault
                
    MsgBox Mensaje_Adicion_Exitosa, vbExclamation
        
    frmMainMdi.stbMdi.Panels(3).Text = "Acción"
    
    cmdReservar.Visible = True
    cmdOpcion.Visible = True
    With tabPagos
        .TabEnabled(0) = True
        .TabEnabled(1) = False
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

Private Function TodoOk() As Boolean
        
    TodoOk = False
    
    
    Dim adoRegistro As ADODB.Recordset
    Dim str_msg As String, str_Tipo As String
    Dim str_pwd As String
    Dim dblTipoCambioTmp As Double
    
     If Trim(strCodGasto) = Valor_Caracter Then
        MsgBox "Debe Seleccionar el Concepto.", vbCritical
        cboGasto.SetFocus
        Exit Function
    End If
    
    If Trim(txtDescripIngreso.Text) = Valor_Caracter Then
        MsgBox "Debe Ingresar la Descripción del Movimiento Bancario.", vbCritical
        txtDescripIngreso.SetFocus
        Exit Function
    End If
        
    If Trim(strTipMov) = Valor_Caracter Then
        MsgBox "Debe Seleccionar Tipo de Movimiento.", vbCritical
        cboTipoMov.SetFocus
        Exit Function
    End If
                        
    If Trim(strFormPago) = Valor_Caracter Then
        MsgBox "Debe Seleccionar la forma de pago.", vbCritical
        cboFormaPago.SetFocus
        Exit Function
    End If
                            
    If Trim(lblProveedor.Caption) = "" Then
        MsgBox "Debe Indicar el Banco.", vbCritical
        Exit Function
    End If
             
    If Trim(strCodMoneda) = Valor_Caracter Then
        MsgBox "Debe Seleccionar la Moneda.", vbCritical
        cboMoneda.SetFocus
        Exit Function
    End If
    
    If cboCuentas.ListIndex = 0 Then
        MsgBox "Seleccione la Cuenta", vbCritical, gstrNombreEmpresa
        cboCuentas.SetFocus
        Exit Function
    End If
        
    If CCur(txtMontoPago.Text) = 0 Then
        MsgBox "El Monto de Pago no puede ser cero!.", vbCritical, gstrNombreEmpresa
        Exit Function
    End If
    
    If Trim(txtNroVoucher.Text) = "" Then
        MsgBox "Ingrese el numero de Voucher o Transaccion", vbCritical, gstrNombreEmpresa
        txtNroVoucher.SetFocus
        Exit Function
    End If
    
    If strCodMoneda <> Codigo_Moneda_Local Then
        
        dblTipoCambioTmp = ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, gdatFechaActual, Codigo_Moneda_Local, strCodMoneda)
'       dblTipoCambioOrden = ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, dtpFechaMovimBCDesde.Value, Codigo_Moneda_Local, strCodMonedaOrden)

        If dblTipoCambioTmp = 0 Then
            MsgBox "No existe tipo de cambio para procesar la operacion", vbOKOnly + vbCritical, Me.Caption
            Exit Function
        End If
        
    End If
    
    '*** Si todo paso OK ***
    TodoOk = True
  
End Function
Private Function TodoOkBackOffice() As Boolean
        
    TodoOkBackOffice = False
        
    If tdgConsulta.SelBookmarks.Count - 1 = -1 Then
        MsgBox "Debe seleccionar registros para enviar !", vbCritical, gstrNombreEmpresa
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

Private Function TodoOkAnular() As Boolean
    
    TodoOkAnular = False
        
    If tdgConsulta.SelBookmarks.Count - 1 = -1 Or tdgConsulta.SelBookmarks.Count - 1 > 0 Then
        MsgBox "Debe Seleccionar un Registro para Anular", vbCritical, gstrNombreEmpresa
        Exit Function
    End If
        
    If cboEstadoOperacionBusqueda.ListIndex > -1 Then
        If strEstadoOperacionBusqueda = "02" Then
            MsgBox "No se puede anular un registro ya procesado ", vbOKOnly + vbCritical, Me.Caption
            Exit Function
        ElseIf strEstadoOperacionBusqueda = "03" Then
            MsgBox "Este registro ya esta anulado ", vbOKOnly + vbCritical, Me.Caption
            Exit Function
        ElseIf strEstadoOperacionBusqueda = "04" Then
            MsgBox "No se puede anular un registro ya contabilizado ", vbOKOnly + vbCritical, Me.Caption
            Exit Function
        ElseIf strEstadoOperacionBusqueda = "05" Then
            MsgBox "No se puede Anular un registro que ha sido Reversado ", vbOKOnly + vbCritical, Me.Caption
            Exit Function
        End If
       
    Else
        MsgBox "Debe seleccionar algun estado de operacion", vbCritical, gstrNombreEmpresa
        If cboEstadoOperacionBusqueda.Enabled Then cboEstadoOperacionBusqueda.SetFocus
        Exit Function
    End If
        
    '*** Si todo paso OK ***
    TodoOkAnular = True
  
End Function

Private Function TodoOkReversar() As Boolean
    
    
    TodoOkReversar = False
        
    If tdgConsulta.SelBookmarks.Count - 1 = -1 Or tdgConsulta.SelBookmarks.Count - 1 > 0 Then
        MsgBox "Debe seleccionar un registro para reversar", vbCritical, gstrNombreEmpresa
        Exit Function
    End If
        
    If cboEstadoOperacionBusqueda.ListIndex >= 0 Then
        If strEstadoOperacionBusqueda <> "02" And strEstadoOperacionBusqueda <> "04" Then
            MsgBox "Solo se puede reversar operaciones ya procesadas o contabilizadas", vbOKOnly + vbCritical, Me.Caption
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

Public Sub SubImprimir(index As Integer)

    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    Dim strFechaDesde           As String, strFechaHasta        As String
    Dim strSeleccionRegistro    As String
 
    
    Select Case index
        Case 1
        
            gstrNameRepo = "MovimientoBancario"
            
            strSeleccionRegistro = "{MovimientoFondo.FechaRegistro} IN 'Fch1' TO 'Fch2'"
            gstrSelFrml = strSeleccionRegistro
            frmRangoFecha.Show vbModal
                        
            If gstrSelFrml = "0" Then
                Set frmReporte = New frmVisorReporte
    
                ReDim aReportParamS(4)
                ReDim aReportParamFn(1)
                ReDim aReportParamF(1)
                            
                aReportParamFn(0) = "Fondo"
                aReportParamFn(1) = "NombreEmpresa"
                            
                aReportParamF(0) = Trim(cboFondo.Text)
                aReportParamF(1) = gstrNombreEmpresa & Space(1)
                            
                aReportParamS(0) = strCodFondo
                aReportParamS(1) = gstrCodAdministradora
                aReportParamS(2) = Convertyyyymmdd(CVDate(gstrFchDel)) 'Convertyyyymmdd(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10))
                aReportParamS(3) = Convertyyyymmdd(CVDate(gstrFchAl)) 'Convertyyyymmdd(DateAdd("d", 1, CVDate(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10))))
                aReportParamS(4) = strEstadoOperacionBusqueda
                
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
    Dim strFechaDesde           As String, strFechaHasta        As String
    Dim strSeleccionRegistro    As String
 
  
            gstrNameRepo = "MovimientoBancarioGrilla"
                        
                Set frmReporte = New frmVisorReporte
    
                ReDim aReportParamS(6)
                ReDim aReportParamFn(5)
                ReDim aReportParamF(5)
                            
                aReportParamFn(0) = "Fondo"
                aReportParamFn(1) = "NombreEmpresa"
                aReportParamFn(2) = "Usuario"
                aReportParamFn(3) = "Hora"
                aReportParamFn(4) = "Fecha"
                aReportParamFn(5) = "Estado"
                 
                aReportParamF(0) = Trim(cboFondo.Text)
                aReportParamF(1) = gstrNombreEmpresa & Space(1)
                aReportParamF(2) = gstrLogin
                aReportParamF(3) = Format(Time(), "hh:mm:ss")
                aReportParamF(4) = gstrNombreEmpresa & Space(1)
                aReportParamF(5) = cboEstadoOperacionBusqueda.Text
                            
                aReportParamS(0) = strCodFondo
                aReportParamS(1) = gstrCodAdministradora
                aReportParamS(2) = Convertyyyymmdd(dtpFechaMovimBCDesde.Value)
                aReportParamS(3) = Convertyyyymmdd(dtpFechaMovimBCHasta.Value)
                aReportParamS(4) = strEstadoOperacionBusqueda
                aReportParamS(5) = Codigo_Comprobante_Documento_Emitido_Bancos
               If strCodCuentaBancaria <> Valor_Caracter Then
                 aReportParamS(6) = strCodCuentaBancaria
               Else
                 aReportParamS(6) = Valor_Comodin
               End If
                
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


    If Trim(strCodCuentaBancaria) <> Valor_Caracter Then

        strSQL = "SELECT NumRegistro,CodTipoComprobante,CodProveedor,DescripRegistro,RC.CodMoneda,ValorTotal, " & _
        "TCP.DescripTipoComprobantePago DescripTipoComprobante, CodSigno,FechaRegistro,DescripPersona DescripProveedor,RC.NumGasto, " & _
        "RC.CodFileGasto,RC.CodCuenta,RC.NumComprobante,RC.CodFormaPago,RC.CodCuentaBanco,RC.CodFileBanco,RC.CodAnaliticaBanco,BC.CodBanco,BC.DescripCuenta " & _
        "FROM RegistroCompra RC JOIN TipoComprobantePago TCP ON(TCP.CodTipoComprobantePago=RC.CodTipoComprobante) " & _
        "JOIN Moneda MON ON(MON.CodMoneda=RC.CodMoneda) " & _
        "JOIN InstitucionPersona IP ON(IP.CodPersona=RC.CodProveedor AND IP.TipoPersona=RC.TipoProveedor) " & _
        "JOIN BancoCuenta BC ON (BC.CodFondo=RC.CodFondo AND BC.CodAdministradora=RC.CodAdministradora " & _
        "AND BC.CodCuentaActivo=RC.CodCuentaBanco AND BC.CodFile=RC.CodFileBanco AND BC.CodAnalitica=RC.CodAnaliticaBanco) " & _
        "WHERE (FechaRegistro>='" & Convertyyyymmdd(dtpFechaMovimBCDesde.Value) & "' AND FechaRegistro<'" & Convertyyyymmdd(DateAdd("d", 1, dtpFechaMovimBCHasta.Value)) & "') AND " & _
        "RC.CodAdministradora='" & gstrCodAdministradora & "' AND RC.CodFondo='" & strCodFondo & "' AND RC.CodTipoComprobante = '" & Codigo_Comprobante_Documento_Emitido_Bancos & "' AND RC.CodCuentaBanco =  '" & strCodCuentaBancaria & "' AND RC.Estado = '" & strEstadoOperacionBusqueda & "' " & _
        "ORDER BY NumRegistro"

    Else
        
        strSQL = "SELECT NumRegistro,CodTipoComprobante,CodProveedor,DescripRegistro,RC.CodMoneda,ValorTotal, " & _
        "TCP.DescripTipoComprobantePago DescripTipoComprobante, CodSigno,FechaRegistro,DescripPersona DescripProveedor,RC.NumGasto, " & _
        "RC.CodFileGasto,RC.CodCuenta,RC.NumComprobante,RC.CodFormaPago,RC.CodCuentaBanco,RC.CodFileBanco,RC.CodAnaliticaBanco,BC.CodBanco,BC.DescripCuenta " & _
        "FROM RegistroCompra RC JOIN TipoComprobantePago TCP ON(TCP.CodTipoComprobantePago=RC.CodTipoComprobante) " & _
        "JOIN Moneda MON ON(MON.CodMoneda=RC.CodMoneda) " & _
        "JOIN InstitucionPersona IP ON(IP.CodPersona=RC.CodProveedor AND IP.TipoPersona=RC.TipoProveedor) " & _
        "JOIN BancoCuenta BC ON (BC.CodFondo=RC.CodFondo AND BC.CodAdministradora=RC.CodAdministradora " & _
        "AND BC.CodCuentaActivo=RC.CodCuentaBanco AND BC.CodFile=RC.CodFileBanco AND BC.CodAnalitica=RC.CodAnaliticaBanco) " & _
        "WHERE (FechaRegistro>='" & Convertyyyymmdd(dtpFechaMovimBCDesde.Value) & "' AND FechaRegistro<'" & Convertyyyymmdd(DateAdd("d", 1, dtpFechaMovimBCHasta.Value)) & "') AND " & _
        "RC.CodAdministradora='" & gstrCodAdministradora & "' AND RC.CodFondo='" & strCodFondo & "' AND RC.CodTipoComprobante = '" & Codigo_Comprobante_Documento_Emitido_Bancos & "' AND RC.Estado = '" & strEstadoOperacionBusqueda & "' " & _
        "ORDER BY NumRegistro"
    
    
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
    
    
    If adoConsulta.RecordCount > 0 Then
       
       Dim adoRegistroTotal As ADODB.Recordset
       
       strEstado = Reg_Consulta
       
       txtTotalSeleccionado.Text = "0"
       
       txtTotalSeleccionadoME.Text = "0"
       
       Set adoRegistroTotal = New ADODB.Recordset
       
        With adoComm
            strSQL = "SELECT COALESCE(SUM(ValorTotal),0) MontoTotal FROM RegistroCompra RC " & _
                " JOIN BancoCuenta BC ON (BC.CodFondo=RC.CodFondo AND BC.CodAdministradora=RC.CodAdministradora " & _
                "AND BC.CodCuentaActivo=RC.CodCuentaBanco AND BC.CodFile=RC.CodFileBanco AND BC.CodAnalitica=RC.CodAnaliticaBanco) " & _
                "WHERE RC.CodFondo='" & strCodFondo & "' AND RC.CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "Estado = '" & strEstadoOperacionBusqueda & "' AND (FechaRegistro>='" & Convertyyyymmdd(dtpFechaMovimBCDesde.Value) & "' AND " & _
                "FechaRegistro<'" & Convertyyyymmdd(DateAdd("d", 1, dtpFechaMovimBCHasta.Value)) & "') AND CodTipoComprobante = '" & Codigo_Comprobante_Documento_Emitido_Bancos & "' "
'
                
                .CommandText = strSQL + "AND RC.CodMoneda = '" + Codigo_Moneda_Local + "'"
                Set adoRegistroTotal = .Execute
                
                If Not adoRegistroTotal.EOF Then
                    txtTotal.Text = CStr(adoRegistroTotal("MontoTotal"))
                End If
                
                .CommandText = strSQL + "AND RC.CodMoneda = '" + Codigo_Moneda_Dolar_Americano + "'"
                Set adoRegistroTotal = .Execute
                
                If Not adoRegistroTotal.EOF Then
                    txtTotalME.Text = CStr(adoRegistroTotal("MontoTotal"))
                End If
                
                adoRegistroTotal.Close: Set adoRegistroTotal = Nothing
        
        End With
            
    Else
        txtTotal.Text = "0": txtTotalSeleccionado.Text = "0"
    End If
    
    
    Call AutoAjustarGrillas
    
    tdgConsulta.Refresh
    tdgConsulta.MoveFirst
    
    If adoConsulta.RecordCount > 0 Then strEstado = Reg_Consulta
  
  
End Sub

Public Sub Modificar()
       
    If strEstadoOperacionBusqueda = "02" Then
        MsgBox "No se puede modificar un registro ya procesado ", vbOKOnly + vbCritical, Me.Caption
        Exit Sub
    ElseIf strEstadoOperacionBusqueda = "03" Then
        MsgBox "No se puede modificar un registro anulado ", vbOKOnly + vbCritical, Me.Caption
        Exit Sub
    ElseIf strEstadoOperacionBusqueda = "04" Then
        MsgBox "No se puede modificar un registro ya contabilizado ", vbOKOnly + vbCritical, Me.Caption
        Exit Sub
    ElseIf strEstadoOperacionBusqueda = "05" Then
        MsgBox "No se puede modificar un registro Reversado ", vbOKOnly + vbCritical, Me.Caption
        Exit Sub
    End If
    
    If strEstado = Reg_Consulta Then
    
        strEstado = Reg_Edicion
        
        LlenarFormulario strEstado
        
        cmdOpcion.Visible = False
        
        With tabPagos
            .TabEnabled(0) = False
            .TabEnabled(1) = True
            .Tab = 1
        End With
        
'        If strEstadoOperacionBusqueda = "04" Then
'            strEstado = Reg_Adicion
'        End If
    End If
    
End Sub

Private Sub LlenarFormulario(strModo As String)

    Dim intNumSecuencial    As Integer, intRegistro As Integer
    Dim adoRegistro As New ADODB.Recordset, adoRegistro2 As New ADODB.Recordset, adoRegistro3 As New ADODB.Recordset
    Dim strSQL As String, strAuxCodBanco As String
    
    Select Case strModo
    
        Case Reg_Edicion
            
            lblDescripFondo.Caption = Trim(cboFondo.Text)
            
            cboGasto.ListIndex = -1
            If cboGasto.ListCount > 0 Then cboGasto.ListIndex = 0
            
            Call CargarGastos
            
            intRegistro = ObtenerItemLista(arrGasto(), tdgConsulta.Columns(10))
            If intRegistro >= 0 Then cboGasto.ListIndex = intRegistro
            
            txtDescripIngreso.Text = Trim(tdgConsulta.Columns(4))
            
            intRegistro = ObtenerItemLista(arrTipMov(), tdgConsulta.Columns(11))
            If intRegistro >= 0 Then cboTipoMov.ListIndex = intRegistro
            
            intRegistro = ObtenerItemLista(arrFormaPago(), tdgConsulta.Columns(12))
            If intRegistro >= 0 Then cboFormaPago.ListIndex = intRegistro
        
            With adoComm
            
                .CommandText = "SELECT IP.CodPersona CODIGO, IP.TipoPersona, AP.DescripParametro TipoIdentidad, " & _
                "IP.NumIdentidad, IP.DescripPersona DESCRIP,IP.Direccion1 + IP.Direccion2 Direccion FROM InstitucionPersona IP " & _
                "JOIN AuxiliarParametro AP ON(AP.CodParametro=IP.TipoIdentidad AND AP.CodTipoParametro='TIPIDE') " & _
                "WHERE IP.TipoPersona='04' AND IP.IndVigente='X' AND IP.IndBanco = 'X' AND CodPersona='" & tdgConsulta.Columns(13) & "'"
                
                Set adoRegistro2 = .Execute
                
            End With
            
            If Not adoRegistro2.EOF Then
            
                Do Until adoRegistro2.EOF
            
                    lblProveedor.Caption = adoRegistro2.Fields("DESCRIP")
                    lblCodProveedor.Caption = adoRegistro2.Fields("CODIGO")
                    lblTipoDocID.Caption = adoRegistro2.Fields("TipoIdentidad")
                    lblNumDocID.Caption = adoRegistro2.Fields("NumIdentidad")
                    
                    adoRegistro2.MoveNext
            
                Loop
                
                adoRegistro2.Close: Set adoRegistro2 = Nothing
                
            End If
            
            intRegistro = ObtenerItemLista(arrMoneda(), tdgConsulta.Columns(7))
            If intRegistro >= 0 Then cboMoneda.ListIndex = intRegistro
            
            dtpFechaActual.Value = tdgConsulta.Columns(2)
            
            txtMontoPago.Text = tdgConsulta.Columns(5)
            
            txtNroVoucher.Text = Trim(tdgConsulta.Columns(14))
            
            intRegistro = ObtenerItemLista(arrEstadoGasto(), Valor_Indicador)
            If intRegistro >= 0 Then cboEstado.ListIndex = intRegistro
            
            
            With adoComm
            
                .CommandText = "SELECT CodBanco  FROM BancoCuenta WHERE CodMoneda='" & strCodMoneda & "' " & _
                "AND IndVigente='" & Valor_Indicador & "' AND CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' " & _
                "AND CodFile='" & tdgConsulta.Columns(16) & "' AND CodAnalitica='" & tdgConsulta.Columns(17) & "' AND CodCuentaActivo='" & tdgConsulta.Columns(15) & "'"
                
                Set adoRegistro3 = .Execute
                
            End With
            
            If Not adoRegistro3.EOF Then
            
                Do Until adoRegistro3.EOF
            
                    strAuxCodBanco = adoRegistro3.Fields("CodBanco")
                    
                    adoRegistro3.MoveNext
            
                Loop
                
                adoRegistro3.Close: Set adoRegistro3 = Nothing
                
            End If
            
            intRegistro = ObtenerItemLista(arrCuentas(), Trim(tdgConsulta.Columns(16)) & Trim(tdgConsulta.Columns(17)) & strAuxCodBanco & Trim(tdgConsulta.Columns(15)))
            If intRegistro >= 0 Then cboCuentas.ListIndex = intRegistro
            
                 
        Case Reg_Adicion
            ''Llenar los combos del formulario
                        
            lblDescripFondo.Caption = Trim(cboFondo.Text)
            lblProveedor.Caption = "": lblTipoDocID.Caption = "": lblCodProveedor.Caption = 0
                        
            Call CargarGastos
                        
            cboGasto.ListIndex = -1
            If cboGasto.ListCount > 0 Then cboGasto.ListIndex = 0
                                                
            cboFormaPago.ListIndex = -1
            If cboFormaPago.ListCount >= 0 Then cboFormaPago.ListIndex = 2
                        
                        
            cboTipoGasto.ListIndex = -1
            cboModalidadPago.ListIndex = -1
            cboCreditoFiscal.ListIndex = -1
            
            '*** POR DEFAULT CARGO ***
            intRegistro = ObtenerItemLista(arrTipMov(), Tipo_Retiro)
            cboTipoMov.ListIndex = intRegistro
            '*************************
             
            cboMoneda.ListIndex = 0
            
            txtMontoPago.Text = "0"
            
            txtNroVoucher.Text = Valor_Caracter
            
            lblNumDocID.Caption = Valor_Caracter
            
            lblCodProveedor.Caption = Valor_Caracter
            
            txtDescripIngreso = Valor_Caracter
            
            intRegistro = ObtenerItemLista(arrTipoGasto(), Codigo_Tipo_Gasto_Unico)
            If intRegistro >= 0 Then cboTipoGasto.ListIndex = intRegistro
                        
            intRegistro = ObtenerItemLista(arrEstadoGasto(), Valor_Indicador)
            If intRegistro >= 0 Then cboEstado.ListIndex = intRegistro
                        
            intRegistro = ObtenerItemLista(arrMoneda(), strCodMoneda)
            If intRegistro >= 0 Then cboMoneda.ListIndex = intRegistro
            
            intRegistro = ObtenerItemLista(arrModalidadPago(), Codigo_Modalidad_Pago_Vencimiento)
            If intRegistro >= 0 Then cboModalidadPago.ListIndex = intRegistro
                        
                        
            intRegistro = ObtenerItemLista(arrCreditoFiscal(), Codigo_Tipo_Credito_RentaGravada)
            If intRegistro >= 0 Then cboCreditoFiscal.ListIndex = 1 ' intRegistro
                        
    End Select
    
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
        "WHERE CodFondo='" & strCodFondo & "' AND FCG.CodAdministradora='" & gstrCodAdministradora & "' AND " & _
        "(PCG.CodCuenta LIKE '639111%' or PCG.CodCuenta LIKE '632213%' or PCG.CodCuenta LIKE '644311%'  or PCG.CodCuenta LIKE '641211%' or PCG.CodCuenta LIKE '679311')  " & _
        "ORDER BY DescripCuenta"
        CargarControlLista strSQL, cboGasto, arrGasto(), Sel_Defecto
    
End Sub

'Public Sub SubImprimir(index As Integer)
'
'    Dim frmReporte              As frmVisorReporte
'    Dim aReportParamS(), aReportParamF(), aReportParamFn()
'    Dim strFechaDesde           As String, strFechaHasta        As String
'    Dim intAccion               As Integer
'    Dim lngNumError             As Long
'
'
'
'    Select Case index
'        Case 1
'            gstrNameRepo = "PagoCuotaSuscripcion"
'
'            Set frmReporte = New frmVisorReporte
'
'            ReDim aReportParamS(1)
'            ReDim aReportParamFn(2)
'            ReDim aReportParamF(2)
'
'            aReportParamFn(0) = "Usuario"
'            aReportParamFn(1) = "Hora"
'            aReportParamFn(2) = "NombreEmpresa"
'
'            aReportParamF(0) = gstrLogin
'            aReportParamF(1) = Format(Time(), "hh:mm:ss")
'            aReportParamF(2) = gstrNombreEmpresa & Space(1)
'
'            aReportParamS(0) = "001"
'            aReportParamS(1) = gstrCodAdministradora
'
'    End Select
'
'    gstrSelFrml = Valor_Caracter
'    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"
'
'    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())
'
'    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
'    frmReporte.Show vbModal
'
'    Set frmReporte = Nothing
'
'    Screen.MousePointer = vbNormal
'
'End Sub
Public Sub Adicionar()
    
    strEstado = Reg_Adicion
    LlenarFormulario strEstado
    cmdOpcion.Visible = False
    cmdReservar.Visible = False
    With tabPagos
        .TabEnabled(0) = False
        .TabEnabled(1) = True
        .Tab = 1
    End With
                
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
    
    '*** Tipo de Movimiento ***
    strSQL = "select CodParametro CODIGO,DescripParametro DESCRIP from AuxiliarParametro Where CodTipoParametro = 'TIPMVB' Order By DescripParametro"
    CargarControlLista strSQL, cboTipoMov, arrTipMov(), Valor_Caracter
    If cboTipoMov.ListCount > 0 Then cboTipoMov.ListIndex = 0
    
    
    '*** Cuentas Bancarias  ***
'     strSQL = "SELECT (CodFile + CodAnalitica + CodBanco + CodCuentaActivo) CODIGO,(RTRIM(DescripCuenta) + SPACE(1) + NumCuenta) DESCRIP FROM BancoCuenta " & _
'              "WHERE IndVigente='X' AND " & _
'              "CodFondo='" & gstrCodFondoContable & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
'    CargarControlLista strSQL, cboCuentasBancarias, arrCuentas(), Sel_Defecto
'    If cboCuentasBancarias.ListCount > 0 Then cboCuentasBancarias.ListIndex = 0
    Call CargarCuentasBancarias

    '*** Moneda ***
    strSQL = "{ call up_ACSelDatos(2) }"
    CargarControlLista strSQL, cboMoneda, arrMoneda(), Sel_Defecto
    If cboMoneda.ListCount > 0 Then cboMoneda.ListIndex = 0
             
    '*** Forma de Pago ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='MEDPAG' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboFormaPago, arrFormaPago(), Sel_Defecto
                    
    '*** Tipo de Gasto ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPPAG' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboTipoGasto, arrTipoGasto(), Valor_Caracter
    
    
    '*** Estados ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP From AuxiliarParametro WHERE CodTipoParametro='INDREG' AND CodParametro<>'03' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboEstado, arrEstadoGasto(), Valor_Caracter
    
    
    '*** Modalidad de Pago ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='MODPAG' ORDER BY CodParametro"
    CargarControlLista strSQL, cboModalidadPago, arrModalidadPago(), Valor_Caracter
    
    
    '*** Tipo Crédito Fiscal ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP From AuxiliarParametro WHERE CodTipoParametro='CREFIS' ORDER BY CodParametro"
    CargarControlLista strSQL, cboCreditoFiscal, arrCreditoFiscal(), Valor_Caracter
    
    
    '**** Estado del Movimiento Bancario ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP from AuxiliarParametro Where CodTipoParametro = 'ESACU2' Order By CodParametro"
    CargarControlLista strSQL, cboEstadoOperacionBusqueda, arrEstadoOperacionBusqueda(), Valor_Caracter
    If cboEstadoOperacionBusqueda.ListCount > 0 Then cboEstadoOperacionBusqueda.ListIndex = 0
    
        
End Sub
Private Sub InicializarValores()

    '*** Valores Iniciales ***
    strEstado = Reg_Defecto
    tabPagos.Tab = 0
    tabPagos.TabEnabled(1) = False
    
    strCodFileGasto = "099"
    strCodTipoValor = Codigo_Tipo_Costo_Monto
    
    dtpFechaMovimBCDesde.Value = gdatFechaActual
    dtpFechaMovimBCHasta.Value = dtpFechaMovimBCDesde.Value
    
    '*** Ancho por defecto de las columnas de la grilla ***
'    tdgConsulta.Columns(0).Width = tdgConsulta.Width * 0.01 * 16
'    tdgConsulta.Columns(1).Width = tdgConsulta.Width * 0.01 * 9
'    tdgConsulta.Columns(2).Width = tdgConsulta.Width * 0.01 * 16
'    tdgConsulta.Columns(3).Width = tdgConsulta.Width * 0.01 * 34
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    
End Sub


Private Sub Form_Resize()
    Call AutoAjustarGrillas
End Sub

Private Sub lblNumSecuencial_Click()

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

Private Sub tdgConsulta_SelChange(Cancel As Integer)
    

    Dim dblMonto                    As Double, dblMontoAcumulado             As Double
    Dim dblMontoContable            As Double, dblMontoContableAcumulado    As Double
    Dim intRegistro                 As Integer, intContador                 As Integer
    Dim intNumGastoSel              As Long, dblMontoLiq                    As Double
    Dim dblMontoAcumuladoLiq        As Double
    
    If tdgConsulta.SelBookmarks.Count < 1 Then Exit Sub
    
    adoConsulta.MoveFirst
    adoConsulta.Move CLng(tdgConsulta.SelBookmarks.Count - 1), 0
    
    txtTotalSeleccionado.Text = "0"
    txtTotalSeleccionadoME.Text = "0"
    

    intContador = tdgConsulta.SelBookmarks.Count - 1
    
    adoConsulta.MoveFirst
    
    dblMontoAcumuladoLiq = 0
    
    'Si la seleccion es multiple, se puede ingresar el tipo y numero de documento
    If intContador > 0 Then
        Call HabilitarSeleccionMultiple(True)
    Else
        Call HabilitarSeleccionMultiple(False)
    End If
    
    For intRegistro = 0 To intContador
        adoConsulta.MoveFirst
        
        'tdgConsulta.Row = tdgConsulta.SelBookmarks(intRegistro) - 1
        adoConsulta.Move CLng(tdgConsulta.SelBookmarks(intRegistro) - 1), 0
        tdgConsulta.Refresh
        
        If strEstadoOperacionBusqueda = "01" Then

          If intRegistro = 0 Then

              strCodMonedaOrden = tdgConsulta.Columns("CodMoneda")


'              If strCodMonedaOrden <> Codigo_Moneda_Local Then
'
'                dblTipoCambioOrden = ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, gdatFechaActual, Codigo_Moneda_Local, strCodMonedaOrden)
''                  dblTipoCambioOrden = ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, dtpFechaMovimBCDesde.Value, Codigo_Moneda_Local, strCodMonedaOrden)
'
'                If dblTipoCambioOrden = 0 Then
'                    MsgBox "No existe tipo de cambio para procesar la(s) orden(es) seleccionada(s)", vbCritical, Me.Caption
'                    tdgConsulta.ReBind
'                    tdgConsulta.Row = 0
'                    Exit Sub
'                End If
'
'              Else
'                dblTipoCambioOrden = 1
'              End If

          End If
          
          If tdgConsulta.Columns("CodMoneda") <> strCodMonedaOrden Then
              MsgBox "No se pueden confirmar masivamente ordenes con distinta moneda", vbCritical, Me.Caption
              tdgConsulta.ReBind
              tdgConsulta.Row = 0
              Exit Sub
          End If
        
        End If

       
        dblMonto = CDbl(tdgConsulta.Columns("Importe"))
        
        dblMontoContable = Round(dblMonto * dblTipoCambioOrden, 2)
        
        dblMontoAcumulado = dblMontoAcumulado + dblMonto
        dblMontoContableAcumulado = dblMontoContableAcumulado + dblMontoContable
              
  
        dblMontoLiq = dblMonto
        dblMontoAcumuladoLiq = dblMontoAcumuladoLiq + dblMontoLiq
        
    Next
    
    If strCodMonedaOrden = Codigo_Moneda_Local Then
        txtTotalSeleccionado.Text = CStr(dblMontoAcumuladoLiq)
    Else
        txtTotalSeleccionadoME.Text = CStr(dblMontoAcumuladoLiq)
    End If

End Sub




'Private Sub txtCodParticipe_LostFocus()
'
'    Dim rst As Boolean
'    Dim adoRegistro As ADODB.Recordset
'
'    If Trim(txtCodParticipe.Text) <> Valor_Caracter Then
'        rst = False
'
'        txtCodParticipe.Text = Format(txtCodParticipe.Text, "00000000000000000000")
'
'        With adoComm
'            Set adoRegistro = New ADODB.Recordset
'
'            .CommandText = ""
'            strSQL = "SELECT CodParticipe,AP1.DescripParametro TipoIdentidad,NumIdentidad,DescripParticipe,FechaIngreso,TipoIdentidad CodTipoIdentidad,AP2.DescripParametro TipoMancomuno "
'            strSQL = strSQL & "FROM ParticipeContrato JOIN AuxiliarParametro AP1 ON(AP1.CodParametro=ParticipeContrato.TipoIdentidad AND AP1.CodTipoParametro='TIPIDE') "
'            strSQL = strSQL & "JOIN AuxiliarParametro AP2 ON(AP2.CodParametro=ParticipeContrato.TipoMancomuno AND AP2.CodTipoParametro='TIPMAN') "
'            strSQL = strSQL & "WHERE CodParticipe='" & Trim(txtCodParticipe.Text) & "'"
'            .CommandText = strSQL
'            Set adoRegistro = .Execute
'
'            Do Until adoRegistro.EOF
'                rst = True
'                lblDescripParticipe.Caption = Trim(adoRegistro("DescripParticipe"))
'                strCodParticipe = txtCodParticipe.Text
'                adoRegistro.MoveNext
'            Loop
'            adoRegistro.Close: Set adoRegistro = Nothing
'
'            If Not rst Then
'                strCodParticipe = ""
'                txtCodParticipe.Text = ""
'                lblDescripParticipe.Caption = ""
'                MsgBox "Codigo de Cliente Incorrecto", vbCritical, Me.Caption
'            End If
'
'        End With
'    Else
'        strCodParticipe = ""
'        txtCodParticipe.Text = ""
'        lblDescripParticipe.Caption = ""
'    End If
'
'
'End Sub

'Private Sub txtCodParticipeBusqueda_LostFocus()
'
'    Dim rst As Boolean
'    Dim adoRegistro As ADODB.Recordset
'
'    If Trim(txtCodParticipeBusqueda.Text) <> Valor_Caracter Then
'        rst = False
'
'        txtCodParticipeBusqueda.Text = Format(txtCodParticipeBusqueda.Text, "00000000000000000000")
'
'        With adoComm
'            Set adoRegistro = New ADODB.Recordset
'
'            .CommandText = ""
'            strSQL = "SELECT CodParticipe,AP1.DescripParametro TipoIdentidad,NumIdentidad,DescripParticipe,FechaIngreso,TipoIdentidad CodTipoIdentidad,AP2.DescripParametro TipoMancomuno "
'            strSQL = strSQL & "FROM ParticipeContrato JOIN AuxiliarParametro AP1 ON(AP1.CodParametro=ParticipeContrato.TipoIdentidad AND AP1.CodTipoParametro='TIPIDE') "
'            strSQL = strSQL & "JOIN AuxiliarParametro AP2 ON(AP2.CodParametro=ParticipeContrato.TipoMancomuno AND AP2.CodTipoParametro='TIPMAN') "
'            strSQL = strSQL & "WHERE CodParticipe='" & Trim(txtCodParticipeBusqueda.Text) & "'"
'            .CommandText = strSQL
'            Set adoRegistro = .Execute
'
'            Do Until adoRegistro.EOF
'                rst = True
'                lblDescripParticipeBusqueda.Caption = Trim(adoRegistro("DescripParticipe"))
'                strCodParticipeBusqueda = txtCodParticipeBusqueda.Text
'                adoRegistro.MoveNext
'            Loop
'            adoRegistro.Close: Set adoRegistro = Nothing
'
'            If Not rst Then
'                strCodParticipeBusqueda = ""
'                txtCodParticipeBusqueda.Text = ""
'                lblDescripParticipeBusqueda.Caption = ""
'                MsgBox "Codigo de Cliente Incorrecto", vbCritical, Me.Caption
'            End If
'
'        End With
'    Else
'        strCodParticipeBusqueda = ""
'        txtCodParticipeBusqueda.Text = ""
'        lblDescripParticipeBusqueda.Caption = ""
'    End If
'
'    'Call Buscar
'
'End Sub

'Private Sub txtMontoPago_Change()
'
'    Call FormatoCajaTexto(txtMontoPago, Decimales_Monto)
'
'End Sub
'
'Private Sub txtMontoPago_KeyPress(KeyAscii As Integer)
'
'    Call ValidaCajaTexto(KeyAscii, "M", txtMontoPago, Decimales_Monto)
'
'End Sub

Private Sub txtMontoPago_GotFocus()
    'Call FormatoGotFocus(txtMontoPago)

End Sub

Private Sub txtMontoPago_KeyPress(KeyAscii As Integer)
    
'    Dim numeros As String
'    Dim tecla
'    Dim posi As Integer, cuenta As Integer, posi2 As Integer
'    Dim numerosDecimales As Integer
'
'    numerosDecimales = Decimales_Monto
'
'    numeros = "0123456789."
'
'    tecla = Chr(KeyAscii)
'
'    If tecla = vbTab Or tecla = vbBack Then Exit Sub
'
'    posi = InStr(1, numeros, tecla)
'
'    If InStr(1, numeros, tecla) <> 0 Then
'        If tecla = "." Then
'            For cuenta = 1 To Len(txtMontoPago)
'                If Mid(txtMontoPago, cuenta, 1) = "." Then
'                KeyAscii = 0
'                Exit For
'                End If
'            Next cuenta
'        Else
'            posi2 = InStr(1, txtMontoPago, ".")
'            If posi2 > 0 Then
'                If posi2 <= Len(txtMontoPago) - numerosDecimales Then
'                    KeyAscii = 0
'                End If
'            End If
'        End If
'    Else
'        KeyAscii = 0
'    End If
'
End Sub

Private Sub txtMontoPago_LostFocus()
    Call FormatoCajaTexto(txtMontoPago, Decimales_Monto)
    With txtMontoPago
        .SelStart = 0
        .SelLength = Len(.Text)
         Call FormatoCajaTexto(txtMontoPago, Decimales_Monto)
    End With
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
       .Fields.Append "NumOrdenCobroPago", adVarChar, 10
       .LockType = adLockBatchOptimistic
    End With

    adoRegistroAux.Open

End Sub


Private Sub txtTotal_Change()

Call FormatoCajaTexto(txtTotal, Decimales_Monto)

End Sub

Private Sub txtTotalME_Change()

Call FormatoCajaTexto(txtTotalME, Decimales_Monto)

End Sub

Private Sub txtTotalSeleccionado_Change()

Call FormatoCajaTexto(txtTotalSeleccionado, Decimales_Monto)

End Sub

Private Sub txtTotalSeleccionadoME_Change()

Call FormatoCajaTexto(txtTotalSeleccionadoME, Decimales_Monto)

End Sub



Public Sub HabilitarSeleccionMultiple(blnHabilita As Boolean)
    
    If blnHabilita Then
        strIndSeleccionMultiple = Valor_Indicador
    Else
        strIndSeleccionMultiple = Valor_Caracter
    End If
        
End Sub

Private Sub AutoAjustarGrillas()
    
    Dim i As Integer
    
    If Not adoConsulta.EOF Then
        If adoConsulta.RecordCount > 0 Then
            For i = 1 To tdgConsulta.Columns.Count - 1
                tdgConsulta.Columns(i).AutoSize
            Next
            
            tdgConsulta.Columns(15).AutoSize
        End If
    End If
    
    tdgConsulta.Refresh

End Sub
