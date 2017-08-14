VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{830D5A06-9B70-4F8C-98B6-7A19C4E7760B}#1.0#0"; "TAMControls.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmLiquidacionOrdenCaja 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ordenes de Cobro y Pagos"
   ClientHeight    =   9075
   ClientLeft      =   1260
   ClientTop       =   1425
   ClientWidth     =   12945
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
   ScaleHeight     =   9075
   ScaleWidth      =   12945
   Begin VB.CommandButton cmdLiquidar 
      Caption         =   "&Liquidar"
      Height          =   735
      Left            =   720
      Picture         =   "frmLiquidacionOrdenCaja.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   8340
      Width           =   1200
   End
   Begin TAMControls2.ucBotonEdicion2 cmdOpcion 
      Height          =   735
      Left            =   2070
      TabIndex        =   58
      Top             =   8340
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   1296
      Buttons         =   2
      Caption0        =   "&Eliminar"
      Tag0            =   "4"
      ToolTipText0    =   "Eliminar"
      Caption1        =   "&Buscar"
      Tag1            =   "5"
      ToolTipText1    =   "Buscar"
      UserControlWidth=   2700
   End
   Begin TabDlg.SSTab tabOrdenCobroPago 
      Height          =   8145
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   12885
      _ExtentX        =   22728
      _ExtentY        =   14367
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      ForeColor       =   8388608
      TabCaption(0)   =   "Lista"
      TabPicture(0)   =   "frmLiquidacionOrdenCaja.frx":054C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtTotalSeleccionadoME"
      Tab(0).Control(1)=   "txtTotalME"
      Tab(0).Control(2)=   "txtTotalSeleccionado"
      Tab(0).Control(3)=   "txtTotal"
      Tab(0).Control(4)=   "fraCriterio"
      Tab(0).Control(5)=   "tdgConsulta"
      Tab(0).Control(6)=   "lblDescrip(20)"
      Tab(0).Control(7)=   "lblDescrip(19)"
      Tab(0).Control(8)=   "lblDescrip(17)"
      Tab(0).Control(9)=   "lblDescrip(16)"
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Liquidación"
      TabPicture(1)   =   "frmLiquidacionOrdenCaja.frx":0568
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "cmdAccion"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraConfirmacion"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame fraConfirmacion 
         Caption         =   "Orden de..."
         Height          =   6795
         Left            =   300
         TabIndex        =   8
         Top             =   480
         Width           =   12255
         Begin VB.TextBox txtNumReferencia 
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
            Left            =   2520
            TabIndex        =   48
            Text            =   " "
            Top             =   5850
            Width           =   7365
         End
         Begin VB.TextBox txtNumDocumento 
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
            Left            =   7500
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   41
            Text            =   " "
            Top             =   840
            Width           =   2775
         End
         Begin VB.ComboBox cboTipoDocumento 
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
            Left            =   1920
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   840
            Width           =   3165
         End
         Begin VB.TextBox txtNroDocumento 
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
            Left            =   7500
            TabIndex        =   29
            Top             =   6330
            Visible         =   0   'False
            Width           =   2385
         End
         Begin VB.ComboBox cboFormaPago 
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
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   6330
            Width           =   2295
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
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   7500
            TabIndex        =   24
            Text            =   " 1"
            Top             =   3000
            Width           =   1845
         End
         Begin VB.ComboBox cboMonedaPago 
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
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   3000
            Width           =   2565
         End
         Begin VB.ComboBox cboCuentas 
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
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   3810
            Width           =   7920
         End
         Begin VB.CheckBox ChkMonDiferente 
            Caption         =   "Transacción en distinta moneda"
            Enabled         =   0   'False
            Height          =   255
            Left            =   360
            TabIndex        =   25
            Top             =   6810
            Visible         =   0   'False
            Width           =   3255
         End
         Begin VB.TextBox txtDescripMotivo 
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
            Left            =   2520
            TabIndex        =   9
            Text            =   " "
            Top             =   5445
            Width           =   7365
         End
         Begin MSComCtl2.DTPicker dtpFechaContable 
            Height          =   315
            Left            =   2160
            TabIndex        =   10
            Top             =   375
            Width           =   1635
            _ExtentX        =   2884
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
            Format          =   79233025
            CurrentDate     =   38068
         End
         Begin TAMControls.TAMTextBox txtMonto 
            Height          =   315
            Left            =   7500
            TabIndex        =   53
            Top             =   1920
            Width           =   2415
            _ExtentX        =   4260
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
            Container       =   "frmLiquidacionOrdenCaja.frx":0584
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
         Begin TAMControls.TAMTextBox txtMontoOriginal 
            Height          =   315
            Left            =   7500
            TabIndex        =   64
            Top             =   1560
            Width           =   2415
            _ExtentX        =   4260
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
            Container       =   "frmLiquidacionOrdenCaja.frx":05A0
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
         Begin VB.Label lblMonedaOrdenOriginal 
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
            TabIndex        =   65
            Top             =   1560
            Width           =   465
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto Orden"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   23
            Left            =   5280
            TabIndex        =   63
            Top             =   1560
            Width           =   1110
         End
         Begin VB.Line Line4 
            X1              =   2160
            X2              =   3480
            Y1              =   2160
            Y2              =   2160
         End
         Begin VB.Line Line3 
            X1              =   3480
            X2              =   3480
            Y1              =   1560
            Y2              =   2160
         End
         Begin VB.Line Line2 
            X1              =   2160
            X2              =   3480
            Y1              =   1560
            Y2              =   1560
         End
         Begin VB.Line Line1 
            X1              =   2160
            X2              =   2160
            Y1              =   1560
            Y2              =   2160
         End
         Begin VB.Label lblDescripOrden 
            AutoSize        =   -1  'True
            Caption         =   "COBRO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2280
            TabIndex        =   62
            Top             =   1680
            Width           =   1110
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Nro. de Orden"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   21
            Left            =   5310
            TabIndex        =   55
            Top             =   420
            Width           =   1215
         End
         Begin VB.Label lblNroOrden 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   7500
            TabIndex        =   54
            Top             =   390
            Width           =   2775
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Nro. Referencia Banco"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   18
            Left            =   360
            TabIndex        =   47
            Top             =   5880
            Width           =   2175
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   3
            X1              =   360
            X2              =   11880
            Y1              =   1320
            Y2              =   1320
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Nro. de Documento"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   15
            Left            =   5310
            TabIndex        =   42
            Top             =   900
            Width           =   1665
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Documento"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   14
            Left            =   360
            TabIndex        =   40
            Top             =   900
            Width           =   1410
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   2
            X1              =   360
            X2              =   11880
            Y1              =   2800
            Y2              =   2800
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto Contable"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   13
            Left            =   5280
            TabIndex        =   38
            Top             =   2340
            Width           =   1350
         End
         Begin VB.Label lblMontoContable 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7500
            TabIndex        =   37
            Top             =   2310
            Width           =   2415
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
            TabIndex        =   36
            Top             =   2370
            Width           =   465
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "XXXXX"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   6
            Left            =   3660
            TabIndex        =   35
            Top             =   6840
            Visible         =   0   'False
            Width           =   615
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
            TabIndex        =   34
            Top             =   1920
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
            TabIndex        =   33
            Top             =   4800
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
            Height          =   315
            Index           =   0
            Left            =   10080
            TabIndex        =   32
            Top             =   4380
            Width           =   465
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
            TabIndex        =   31
            Top             =   3030
            Width           =   975
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   0
            X1              =   360
            X2              =   11880
            Y1              =   3550
            Y2              =   3550
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Nro Cheque"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   12
            Left            =   5400
            TabIndex        =   30
            Top             =   6330
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Forma de Pago"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   11
            Left            =   360
            TabIndex        =   28
            Top             =   6330
            Width           =   1455
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   1
            X1              =   360
            X2              =   11880
            Y1              =   5220
            Y2              =   5220
         End
         Begin VB.Label lblMovCuenta 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7530
            TabIndex        =   23
            Top             =   4350
            Width           =   2475
         End
         Begin VB.Label lblSaldoCuenta 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7530
            TabIndex        =   22
            Top             =   4770
            Width           =   2475
         End
         Begin VB.Label lblMonto 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   390
            TabIndex        =   20
            Top             =   4470
            Visible         =   0   'False
            Width           =   1545
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto a Liquidar"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   5280
            TabIndex        =   18
            Top             =   1950
            Width           =   1440
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Liquidación"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   360
            TabIndex        =   17
            Top             =   420
            Width           =   1575
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Depositar a la Cuenta"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   4
            Left            =   390
            TabIndex        =   16
            Top             =   3870
            Width           =   1860
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Motivo"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   5
            Left            =   360
            TabIndex        =   15
            Top             =   5475
            Width           =   1395
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Disponible"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   7
            Left            =   5280
            TabIndex        =   14
            Top             =   4770
            Width           =   1440
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Moneda de Pago"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   8
            Left            =   360
            TabIndex        =   13
            Top             =   3030
            Width           =   1455
         End
         Begin VB.Label lblTipoCambio 
            Caption         =   "Tipo de Cambio Arbitraje"
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   5250
            TabIndex        =   12
            Top             =   3030
            Width           =   2115
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Movimiento Cuenta"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   10
            Left            =   5280
            TabIndex        =   11
            Top             =   4350
            Width           =   1875
         End
      End
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   8040
         TabIndex        =   57
         Top             =   7320
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   1296
         Buttons         =   2
         Caption0        =   "&Procesar"
         Tag0            =   "2"
         ToolTipText0    =   "Procesar"
         Caption1        =   "&Cancelar"
         Tag1            =   "8"
         ToolTipText1    =   "Cancelar"
         UserControlWidth=   2700
      End
      Begin VB.TextBox txtTotalSeleccionadoME 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   -71970
         TabIndex        =   51
         Top             =   6930
         Width           =   2000
      End
      Begin VB.TextBox txtTotalME 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   -67320
         TabIndex        =   49
         Top             =   6960
         Width           =   2000
      End
      Begin VB.TextBox txtTotalSeleccionado 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   -71970
         TabIndex        =   44
         Top             =   6540
         Width           =   2000
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   -67320
         TabIndex        =   43
         Top             =   6480
         Width           =   2000
      End
      Begin VB.Frame fraCriterio 
         Caption         =   "Criterios de búsqueda"
         Height          =   1760
         Left            =   -74700
         TabIndex        =   1
         Top             =   480
         Width           =   12255
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
            Left            =   6720
            Style           =   2  'Dropdown List
            TabIndex        =   60
            Top             =   772
            Width           =   2895
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
            Left            =   1995
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   772
            Width           =   3255
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
            Left            =   1995
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   360
            Width           =   7605
         End
         Begin MSComCtl2.DTPicker dtpFechaConsulta 
            Height          =   315
            Left            =   1995
            TabIndex        =   4
            Top             =   1185
            Width           =   1335
            _ExtentX        =   2355
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
            Format          =   79233025
            CurrentDate     =   38068
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Moneda"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   22
            Left            =   5640
            TabIndex        =   61
            Top             =   772
            Width           =   1065
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Fondo"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   360
            TabIndex        =   6
            Top             =   380
            Width           =   1185
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Obligaciones Al"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   1
            Left            =   360
            TabIndex        =   5
            Top             =   1205
            Width           =   1425
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Estado"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   9
            Left            =   360
            TabIndex        =   2
            Top             =   792
            Width           =   1185
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Height          =   3765
         Left            =   -74700
         OleObjectBlob   =   "frmLiquidacionOrdenCaja.frx":05BC
         TabIndex        =   26
         Top             =   2400
         Width           =   12255
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Monto Seleccionado ME"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   20
         Left            =   -74280
         TabIndex        =   52
         Top             =   6930
         Width           =   2295
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Monto Total ME"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   19
         Left            =   -68880
         TabIndex        =   50
         Top             =   6930
         Width           =   1455
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Monto Seleccionado MN"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   17
         Left            =   -74280
         TabIndex        =   46
         Top             =   6480
         Width           =   2295
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Monto Total MN"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   16
         Left            =   -68880
         TabIndex        =   45
         Top             =   6480
         Width           =   1455
      End
   End
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   9600
      TabIndex        =   56
      Top             =   8310
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
      Caption0        =   "&Salir"
      Tag0            =   "9"
      ToolTipText0    =   "Salir"
      UserControlWidth=   1200
   End
End
Attribute VB_Name = "frmLiquidacionOrdenCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrFondo()          As String, arrTipoCuenta()          As String
Dim arrMonedaPago()     As String, arrEstado()              As String
Dim arrCtaCte()         As String, arrCtaAhorro()           As String
Dim arrCuentas()        As String, arrFormaPago()           As String
Dim arrTipoDocumento()  As String, arrMoneda()              As String

Dim strCodFondo         As String, strCodTipoCuenta         As String
Dim strCodMonedaPago    As String, strCodEstado             As String
Dim strCodCtaCte        As String, strCodCtaAhorro          As String
Dim strEstado           As String, strCodMoneda             As String
Dim strCodMonedaOrden   As String
Dim strCodFile          As String, strCodAnalitica          As String
Dim strCodCuenta        As String, strCodBanco              As String
Dim strCodFormaPago     As String, strSQL                   As String
'*** Para Registro de Ventas ***
Dim curRVMtoSto         As Currency, curRVMtoTot       As Currency
Dim curRVMtoIgv         As Currency

'*** Para Actualizar Tabla de Cupones ***
Dim strVarCodFile       As String, strVarCodAnal     As String
Dim strVarFech          As String, strCodMonedaBusqueda     As String

Dim strMoneda           As String, strDescripOrden      As String
Dim intRow              As Integer, intOrdenG         As Integer
Dim intNumIni           As Integer, intNumFin         As Integer
Dim strModalidadCambio  As String, strCodTipoCambio   As String
Dim dblTipoCambioOrden  As Double, strTipoDocumento   As String
Dim strNumDocumento     As String, strIndSeleccionMultiple  As String
Dim adoRegistro As ADODB.Recordset
Dim adoRegistroAux      As ADODB.Recordset
Dim strCodMonedaParEvaluacion As String
Dim strCodMonedaParPorDefecto As String
Dim intDiasDesplazamiento As Long, strTipoContraparte As String
Dim datFechaConsulta As Date
Dim SelStartTmp As Long
Dim indSortAsc  As Boolean
Dim indSortDesc As Boolean

Public Sub Abrir()

End Sub

Public Sub Adicionar()

End Sub



Public Sub Ayuda()

End Sub

Public Sub Buscar()

    Dim strSQL As String
    
    Set adoRegistro = New ADODB.Recordset
        
    strSQL = "SELECT NumOrdenCobroPago,NumOperacion,FechaRegistro,FechaObligacion,DescripOrden," & _
        "MontoOrdenCobroSaldo,MontoOrdenLiquidacion,TipoOperacion,AP.CodSigno DescripMoneda,MF.CodMoneda,ValorTipoCambio," & _
        "MF.CodContraparte,MF.TipoContraparte,IP.DescripPersona,MF.CodCuenta, " & _
        "MF.CodFile,MF.CodAnalitica,MF.NumGasto,MF.TipoDocumento, MF.NumDocumento " & _
        " FROM MovimientoFondo MF " & _
        " JOIN Moneda AP ON(AP.CodMoneda=MF.CodMoneda) " & _
        " JOIN InstitucionPersona IP ON(IP.CodPersona = MF.CodContraparte AND IP.TipoPersona = MF.TipoContraparte) " & _
        "WHERE CodFondo = '" & strCodFondo & "' AND CodAdministradora = '" & gstrCodAdministradora & "' AND " & _
        "EstadoOrden = '" & strCodEstado & "' AND FechaObligacion < '" & Convertyyyymmdd(DateAdd("d", 1, dtpFechaConsulta.Value)) & "'"
        
    If strCodMonedaBusqueda <> Valor_Caracter Then
        strSQL = strSQL & " AND MF.CodMoneda='" & strCodMonedaBusqueda & "' "
    End If
       
    strSQL = strSQL & "ORDER BY FechaObligacion,NumOrdenCobroPago"
                        
    strEstado = Reg_Defecto
    With adoRegistro
        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open strSQL
'        .ConnectionString = gstrConnectConsulta
'        .RecordSource = strSQL
'        .Refresh
    End With
    
    'tdgConsulta.Refresh
    tdgConsulta.DataSource = adoRegistro
    
    
    If adoRegistro.RecordCount > 0 Then
       
       Dim adoRegistroTotal As ADODB.Recordset
       
       strEstado = Reg_Consulta
       
       txtTotalSeleccionado.Text = "0"
        txtTotalSeleccionadoME.Text = "0"
       
       Set adoRegistroTotal = New ADODB.Recordset
       
        With adoComm
            strSQL = "SELECT COALESCE(SUM(MF.MontoOrdenCobroSaldo),0) MontoTotal FROM MovimientoFondo MF " & _
                "JOIN Moneda AP ON(AP.CodMoneda=MF.CodMoneda) JOIN InstitucionPersona IP ON(IP.CodPersona = MF.CodContraparte AND IP.TipoPersona = MF.TipoContraparte) " & _
                "WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "EstadoOrden = '" & strCodEstado & "' AND FechaObligacion < '" & Convertyyyymmdd(DateAdd("d", 1, dtpFechaConsulta.Value)) & "' AND " & _
                "TipoContraparte = '" & Me.Tag & "' "
                    
              If strCodMonedaBusqueda = Valor_Caracter Then
                
                
                .CommandText = strSQL + "AND MF.CodMoneda = '" + Codigo_Moneda_Local + "'"
                Set adoRegistroTotal = .Execute
                
                If Not adoRegistroTotal.EOF Then
                    txtTotal.Text = CStr(adoRegistroTotal("MontoTotal"))
                End If
                
                .CommandText = strSQL + "AND MF.CodMoneda = '" + Codigo_Moneda_Dolar_Americano + "'"
                Set adoRegistroTotal = .Execute
                
                If Not adoRegistroTotal.EOF Then
                    txtTotalME.Text = CStr(adoRegistroTotal("MontoTotal"))
                End If
                 
                 
               ElseIf strCodMonedaBusqueda = Codigo_Moneda_Local Then
                
                txtTotalME.Text = 0
                
                 .CommandText = strSQL + "AND MF.CodMoneda = '" + Codigo_Moneda_Local + "'"
                Set adoRegistroTotal = .Execute
                
                If Not adoRegistroTotal.EOF Then
                    txtTotal.Text = CStr(adoRegistroTotal("MontoTotal"))
                End If
                
               
               ElseIf strCodMonedaBusqueda = Codigo_Moneda_Dolar_Americano Then
               
                txtTotal.Text = 0
                
               .CommandText = strSQL + "AND MF.CodMoneda = '" + Codigo_Moneda_Dolar_Americano + "'"
                Set adoRegistroTotal = .Execute
                
                If Not adoRegistroTotal.EOF Then
                    txtTotalME.Text = CStr(adoRegistroTotal("MontoTotal"))
                End If
                
               
               End If
               
                 adoRegistroTotal.Close: Set adoRegistroTotal = Nothing
               
        End With
        
        
        
    Else
        txtTotal.Text = "0": txtTotalSeleccionado.Text = "0"
        txtTotalME.Text = "0": txtTotalSeleccionadoME.Text = "0"
    End If
    
        
End Sub

Public Sub Cancelar()

    cmdOpcion.Visible = True
    cmdLiquidar.Visible = True
    With tabOrdenCobroPago
        .TabEnabled(0) = True
        .TabEnabled(1) = False
        .Tab = 0
    End With
    Call Buscar
    txtDescripMotivo.Text = Valor_Caracter
    txtNumReferencia.Text = Valor_Caracter
    tdgConsulta.ReBind
    
End Sub

Private Sub CargarCuentasBancarias()

    Dim strSQL As String
        
    strSQL = "SELECT (CodFile + CodAnalitica + CodBanco + CodCuentaActivo) CODIGO,(RTRIM(DescripCuenta) + SPACE(1) + NumCuenta) DESCRIP FROM BancoCuenta " & _
        "WHERE CodMoneda='" & strCodMonedaPago & "' AND IndVigente='X' AND " & _
        "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
            
    fraConfirmacion.Caption = "Orden de Cobro"
    lblDescripOrden.Caption = "COBRO"
'    If strCodTipoCuenta = Codigo_Tipo_Cuenta_Corriente Then
'        CargarControlLista strSQL, cboCtaCte, arrCtaCte(), Sel_Defecto
'        cboCtaCte.Visible = True: cboCtaAhorro.Visible = False
'        If cboCtaCte.ListCount > 0 Then cboCtaCte.ListIndex = 0
'        lblDescrip(4).Caption = "Depositar a la Cuenta Corriente"
'        If CCur(lblMonto.Caption) < 0 Then
'            lblDescrip(4).Caption = "Retirar de la Cuenta Corriente"
'            fraConfirmacion.Caption = "Orden de Pago"
'        End If
'    ElseIf strCodTipoCuenta = Codigo_Tipo_Cuenta_Ahorro Then
'        CargarControlLista strSQL, cboCtaAhorro, arrCtaAhorro(), Sel_Defecto
'        cboCtaAhorro.Visible = True: cboCtaCte.Visible = False
'        If cboCtaAhorro.ListCount > 0 Then cboCtaAhorro.ListIndex = 0
'        lblDescrip(4).Caption = "Depositar a la Cuenta de Ahorro"
'        If CCur(lblMonto.Caption) < 0 Then
'            lblDescrip(4).Caption = "Retirar de la Cuenta de Ahorro"
'            fraConfirmacion.Caption = "Orden de Pago"
'        End If
'    End If

    CargarControlLista strSQL, cboCuentas, arrCuentas(), Sel_Defecto
    'cboCtaAhorro.Visible = True: cboCtaCte.Visible = False
    If cboCuentas.ListCount > 0 Then cboCuentas.ListIndex = 0
    lblDescrip(4).Caption = "Depositar a la Cuenta"
    'If CCur(lblMonto.Caption) < 0 Then
    If CCur(txtMonto.Text) < 0 Then
        lblDescrip(4).Caption = "Retirar de la Cuenta"
        fraConfirmacion.Caption = "Orden de Pago"
        lblDescripOrden.Caption = "PAGO"
    End If
        
End Sub


Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Ordenes de Cobro/Pago"
    
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

Public Sub Eliminar()

    If strEstado = Reg_Consulta Or strEstado = Reg_Edicion Then
        
        
        
        If strCodEstado = Estado_Caja_Confirmado Then
            MsgBox "No se puede anular un registro ya confirmado ", vbOKOnly + vbCritical, Me.Caption
            Exit Sub
        ElseIf strCodEstado = Estado_Caja_Anulado Then
            MsgBox "Este registro ya esta anulado ", vbOKOnly + vbCritical, Me.Caption
            Exit Sub
        End If
        
        If MsgBox("Se procederá a eliminar la orden de cobro/pago." & vbNewLine & vbNewLine & "Seguro de continuar ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
    
            adoComm.CommandText = "UPDATE MovimientoFondo SET EstadoOrden='" & Estado_Caja_Anulado & "' " & _
                "WHERE NumOrdenCobroPago=" & tdgConsulta.Columns(0) & " AND " & _
                "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
            adoConn.Execute adoComm.CommandText
            
            tabOrdenCobroPago.TabEnabled(0) = True
            tabOrdenCobroPago.TabEnabled(1) = False
            tabOrdenCobroPago.Tab = 0
            Call Buscar
            
            Exit Sub
        End If
        
    End If
    
End Sub

Public Sub Grabar()

    Dim intContador                         As Integer
    Dim intRegistro                         As Integer
    Dim strFechaGrabar                      As String
    Dim strCodProceso                       As String
    Dim strNumCheque                        As String
    Dim strMovimientoFondoLiquidacionXML    As String
    Dim objMovimientoFondoLiquidacionXML    As DOMDocument60
    Dim strTipoCambioReemplazoXML           As String
    Dim objTipoCambioReemplazoXML           As DOMDocument60
    Dim strMsgError                         As String
    Dim strCodTipoCambio                    As String
    
    Dim adoRegistroAuxTC As ADODB.Recordset
    
    Set adoRegistroAuxTC = New ADODB.Recordset
    
    With adoRegistroAuxTC.Fields
        .Append "CodMonedaOrigen", adVarChar, 2
        .Append "CodMonedaCambio", adVarChar, 2
        .Append "ValorTipoCambio", adDecimal, 19
     
        .Item("ValorTipoCambio").Precision = 19
        .Item("ValorTipoCambio").NumericScale = 6
    
    End With
    
    On Error GoTo ErrorHandler
    
    If strEstado = Reg_Defecto Then Exit Sub
    
    If TodoOk() Then
        
        '*** Realizar proceso de contabilización ***
        If MsgBox("Datos correctos. ¿ Procedemos a la Contabilización ?", vbQuestion + vbYesNo, "Observación") = vbNo Then Exit Sub
    
        intContador = tdgConsulta.SelBookmarks.Count - 1
               
        strFechaGrabar = Convertyyyymmdd(dtpFechaContable.Value) & Space(1) & Format(Time, "hh:mm")
        
        If tdgConsulta.Columns("TipoOperacion") = "46" Then
            strCodProceso = "10" 'liquidacion de clientes
        Else
            strCodProceso = "08" 'liquidacion de proveedores
        End If
        
        
        If gstrValorTipoCambioOperacion = "COMPRA" Then
            strCodTipoCambio = "01"
        Else
            strCodTipoCambio = "02"
        End If
        
        
        strNumCheque = Trim(txtNroDocumento.Text)
                                             
        Call ConfiguraRecordsetAuxiliar
        
        With adoComm
                      
            For intRegistro = 0 To intContador
                adoRegistro.MoveFirst
                
                adoRegistro.Move CLng(tdgConsulta.SelBookmarks(intRegistro) - 1), 0
                                
                adoRegistroAux.AddNew
                adoRegistroAux.Fields("CodFondo") = strCodFondo
                adoRegistroAux.Fields("CodAdministradora") = gstrCodAdministradora
                adoRegistroAux.Fields("NumOrdenCobroPago") = tdgConsulta.Columns("NumOrdenCobroPago")
            Next
            
            Call XMLADORecordset(objMovimientoFondoLiquidacionXML, "MovimientoFondoLiquidacion", "Movimiento", adoRegistroAux, strMsgError)
                strMovimientoFondoLiquidacionXML = objMovimientoFondoLiquidacionXML.xml '
                
            ''CrearXMLDetalle (objTipoCambioReemplazoXML)
            adoRegistroAuxTC.Open
            adoRegistroAuxTC.AddNew
            adoRegistroAuxTC("CodMonedaOrigen") = strCodMonedaOrden
            adoRegistroAuxTC("CodMonedaCambio") = strCodMonedaPago
            adoRegistroAuxTC("ValorTipoCambio") = txtTipoCambio.Text
            
            Call XMLADORecordset(objTipoCambioReemplazoXML, "TipoCambioReemplazo", "MonedaTipoCambio", adoRegistroAuxTC, strMsgError)
            strTipoCambioReemplazoXML = objTipoCambioReemplazoXML.xml 'CrearXMLDetalle(objTipoCambioReemplazoXML)

           
            .CommandText = "{ call up_ACProcMovimientoFondoLiquidacion('" & _
                            strCodFondo & "','" & gstrCodAdministradora & "','" & strFechaGrabar & "','" & _
                            Trim(txtDescripMotivo.Text) & "','" & strCodCuenta & "','" & strCodFile & "','" & _
                            strCodAnalitica & "','" & strCodBanco & "','" & strCodFormaPago & "', '" & strNumCheque & "', '" & _
                            strCodMonedaOrden & "'," & CDbl(lblMovCuenta.Caption) & ",'" & strCodMonedaPago & "'," & _
                            CDbl(lblMovCuenta.Caption) & ",''," & CDbl(lblMontoContable.Caption) & ",'" & _
                            strCodTipoCambio & "','" & strCodTipoCambio & "'," & dblTipoCambioOrden & ",'" & _
                            Trim(txtNumReferencia.Text) & "','" & _
                            Trim(frmMainMdi.Tag) & "','" & strCodProceso & "','" & _
                            strMovimientoFondoLiquidacionXML & "','" & _
                            strTipoCambioReemplazoXML & "','" & _
                            "','" & tdgConsulta.Columns("CodContraparte") & "','" & tdgConsulta.Columns("TipoContraparte") & "')}"

            adoConn.Execute .CommandText
                                              
        End With
        
        Me.MousePointer = vbDefault
        
        MsgBox Mensaje_Proceso_Exitoso, vbExclamation
        
        frmMainMdi.stbMdi.Panels(3).Text = "Acción"
        
        cmdOpcion.Visible = True
        cmdLiquidar.Visible = True
        With tabOrdenCobroPago
            .TabEnabled(0) = True
            .TabEnabled(1) = False
            .Tab = 0
        End With
        txtDescripMotivo.Text = Valor_Caracter
        txtNumReferencia.Text = Valor_Caracter
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
       .Fields.Append "NumOrdenCobroPago", adVarChar, 10
       .LockType = adLockBatchOptimistic
    End With

    adoRegistroAux.Open
            
End Sub

Private Function TodoOk() As Boolean
        
    TodoOk = False
    
    If cboCuentas.ListIndex = 0 Then
        MsgBox "Seleccione la Cuenta", vbCritical, gstrNombreEmpresa
        cboCuentas.SetFocus
        Exit Function
    End If
                        
    If Trim(txtDescripMotivo.Text) = Valor_Caracter Then
        MsgBox "Descripción de Motivo de Liquidación No Válido", vbCritical, gstrNombreEmpresa
        txtDescripMotivo.SetFocus
        Exit Function
    End If
    
    If cboFormaPago.ListIndex = 0 Then
        MsgBox "Seleccione la Forma de Pago", vbCritical, gstrNombreEmpresa
        cboFormaPago.SetFocus
        Exit Function
    End If
    
'    If strIndSeleccionMultiple = Valor_Indicador Then
'        If cboTipoDocumento.ListIndex = -1 Then
'            MsgBox "Cuando realiza una selección múltiple debe ingresar el tipo de documento de la liquidación!", vbCritical, gstrNombreEmpresa
'            cboTipoDocumento.SetFocus
'            Exit Function
'        End If
'
'        If Trim(txtNumDocumento.Text) = "" Then
'            MsgBox "Cuando realiza una selección múltiple debe ingresar el número de documento de la liquidación!", vbCritical, gstrNombreEmpresa
'            txtNumDocumento.SetFocus
'            Exit Function
'        End If
'
'    End If
    
    If CCur(lblMovCuenta.Caption) < 0 Then
        If Abs(CCur(lblMovCuenta.Caption)) > CCur(lblSaldoCuenta.Caption) Then
            If MsgBox("El monto a pagar va a sobregirar la cuenta." & vbNewLine & vbNewLine & _
                "Seguro de Continuar ?", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then Exit Function
        End If
    End If
    
    'gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion
    
    If strCodMonedaOrden <> Codigo_Moneda_Local Then
        If ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, dtpFechaContable.Value, strCodMonedaOrden, Codigo_Moneda_Local) = 0 Then
            MsgBox "No existe tipo de cambio para procesar la(s) orden(es) seleccionada(s)", vbCritical, Me.Caption
        End If
    End If
    
    'Verifica liquidacion parcial
    If CDbl(lblMonto.Caption) <> CDbl(txtMonto.Text) And Abs(CDbl(lblMonto.Caption)) < Abs(CDbl(txtMonto.Text)) Then
        If MsgBox("Ud. esta intentado liquidar un monto diferente al monto inicial de " & lblMonedaOrden.Caption & " " & lblMonto.Caption & "." & vbNewLine & vbNewLine & _
                "Se liquidará totalmente la orden." & vbNewLine & vbNewLine & _
                "Seguro de Continuar?", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then Exit Function
    End If
    
    'Verifica liquidacion parcial
    If CDbl(lblMonto.Caption) <> CDbl(txtMonto.Text) And Abs(CDbl(lblMonto.Caption)) > Abs(CDbl(txtMonto.Text)) Then
        If MsgBox("Ud. esta intentado liquidar un monto diferente al monto inicial de " & lblMonedaOrden.Caption & " " & lblMonto.Caption & "." & vbNewLine & vbNewLine & _
                  "Esta quedando pendiente de liquidar el monto de " & lblMonedaOrden.Caption & " " & Format(CStr(CDbl(lblMonto.Caption) - CDbl(txtMonto.Text)), "###,###,###,###,###,##0.00") & "." & vbNewLine & vbNewLine & _
                "Seguro de Continuar?", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then Exit Function
    End If
    
    
    '*** Si todo paso OK ***
    TodoOk = True
  
End Function
Public Sub Imprimir()

    Call SubImprimir(1)
    
End Sub

Public Sub Modificar()

    If tdgConsulta.SelBookmarks.Count < 1 Then Exit Sub
    
    If strCodEstado <> Estado_Caja_NoConfirmado Then
        
        MsgBox "Solo se pueden liquidar registros con estado no confirmado", vbOKOnly + vbCritical, Me.Caption
        Exit Sub
    
    End If
    
    If strEstado = Reg_Consulta Then
        strEstado = Reg_Edicion
        'Call Deshabilita
        LlenarFormulario strEstado
        'cmdOpcion.Visible = False
        With tabOrdenCobroPago
            .TabEnabled(0) = False
            .TabEnabled(1) = True
            .Tab = 1
        End With
    End If
    
End Sub

Private Sub LlenarFormulario(strModo As String)
    
    Select Case strModo
        Case Reg_Edicion
            Dim intRegistro             As Integer
            Dim strCodTipoOperacion     As String
          
            lblNroOrden.Caption = tdgConsulta.Columns("NumOrdenCobroPago").Value
            
            cboMonedaPago.ListIndex = -1
            intRegistro = ObtenerItemLista(arrMonedaPago(), strCodMonedaOrden)
            If intRegistro > 0 Then cboMonedaPago.ListIndex = intRegistro
                                            
            txtDescripMotivo.Text = strDescripOrden
            
            cboFormaPago.ListIndex = -1
            
            '*** POR DEFECTO FORMA DE PAGO CUENTA ***
            intRegistro = ObtenerItemLista(arrFormaPago(), Codigo_FormaPago_Cuenta)
            If intRegistro >= 0 Then cboFormaPago.ListIndex = intRegistro
            '****************************************
            
'            If IsNumeric(tdgConsulta.Columns(8).Value) Then
'                If CDbl(tdgConsulta.Columns(8).Value) > 0 Then
'                    txtTipoCambio.Text = CStr(tdgConsulta.Columns(8).Value)
'                Else
'                    txtTipoCambio.Text = CStr(gdblTipoCambio)
'                End If
'            End If
            
'            txtTipoCambio.Text = CStr(ObtenerTipoCambio(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, dtpFechaContable.Value, strCodMonedaOrden))
'            If CDbl(txtTipoCambio.Text) = 0 Then txtTipoCambio.Text = CStr(ObtenerTipoCambio(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, DateAdd("d", -1, dtpFechaContable.Value), strCodMonedaOrden))
            
            cboCuentas.Enabled = True
            strCodTipoOperacion = Trim(tdgConsulta.Columns(9).Value)
            
            'Si la seleccion NO es multiple, jala el tipo y numero de documento
            If strIndSeleccionMultiple = Valor_Caracter Then
                intRegistro = ObtenerItemLista(arrTipoDocumento(), tdgConsulta.Columns("TipoDocumento").Value)
                If intRegistro >= 0 Then cboTipoDocumento.ListIndex = intRegistro
                txtNumDocumento.Text = Trim(tdgConsulta.Columns("NumDocumento").Value)
                lblNroOrden.Caption = tdgConsulta.Columns("NumOrdenCobroPago").Value
            Else
            'Si la seleccion ES multiple, estos controles se setean para que el usuario ingrese los valores
                cboTipoDocumento.ListIndex = -1
                txtNumDocumento.Text = ""
                lblNroOrden.Caption = "MULTIPLES ORDENES"
            End If
            
            cmdOpcion.Visible = False
            cmdLiquidar.Visible = False
'            If strCodTipoOperacion = Codigo_Caja_Gasto Then
'                If strCodMonedaOrden <> Codigo_Moneda_Local Then
'                    txtTipoCambio.Text = CStr(ObtenerTipoCambio(Codigo_TipoCambio_Conasev, Codigo_Valor_TipoCambioVenta, DateAdd("d", -1, dtpFechaContable.Value), strCodMonedaOrden))
'                Else
'                    txtTipoCambio.Text = CStr(ObtenerTipoCambio(Codigo_TipoCambio_Conasev, Codigo_Valor_TipoCambioVenta, DateAdd("d", -1, dtpFechaContable.Value), "02"))
'                End If
'
'                If CDbl(txtTipoCambio.Text) = 0 Then
'                    MsgBox "Tipo de Cambio para Liquidación de Gastos No Registrada", vbCritical, Me.Caption
'                    cboCuentas.Enabled = False
'                End If
'            End If
                    
    End Select
    
End Sub
Private Sub ObtenerSaldos()

    Dim adoTemporal As ADODB.Recordset
    Dim strFecha    As String, strFechaMas1Dia  As String
    
    strFecha = Convertyyyymmdd(dtpFechaContable.Value)
    strFechaMas1Dia = Convertyyyymmdd(DateAdd("d", 1, dtpFechaContable.Value))
    
    Set adoTemporal = New ADODB.Recordset
    With adoComm
'        .CommandText = "{ call up_ACObtenerSaldoCuentaContable('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
'            strCodFile & "','" & strCodAnalitica & "','" & strFecha & "','" & strFechaMas1Dia & "','" & _
'            strCodCuenta & "','" & strCodMonedaPago & "') }"
            
        .CommandText = "{ call up_ACObtenerSaldoFinalCuenta('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
            strCodFile & "','" & strCodAnalitica & "','" & strFecha & "','" & strFechaMas1Dia & "','" & _
            strCodCuenta & "','" & strCodMonedaPago & "') }"
            
        Set adoTemporal = .Execute
        
        If Not adoTemporal.EOF Then
            lblSaldoCuenta.Caption = CStr(adoTemporal("SaldoCuenta"))
        Else
            lblSaldoCuenta.Caption = "0"
        End If
        adoTemporal.Close: Set adoTemporal = Nothing
    End With
    
End Sub

Public Sub Salir()

    Unload Me
    
End Sub

Public Sub SubImprimir(index As Integer)

    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    Dim strFechaHasta           As String
    Dim strFechaDesde           As String
    Dim strSeleccionRegistro    As String
        
   
    Select Case index
        Case 1
        
            gstrNameRepo = "MovimientoFondo"
            
            strTipoContraparte = Me.Tag
            
            strSeleccionRegistro = "{MovimientoFondo.FechaOrden} IN 'Fch1' TO 'Fch2'"
            gstrSelFrml = strSeleccionRegistro
            frmRangoFecha.Show vbModal
            
        If gstrSelFrml <> "0" Then
        
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
                                    
'            datFechaSiguiente = DateAdd("d", 0, dtpFechaConsulta.Value)
'                                Convertyyyymmdd(datFechaiguiente)
            
            strFechaHasta = Convertyyyymmdd(gstrFchAl)
            
            aReportParamS(0) = strCodFondo
            aReportParamS(1) = gstrCodAdministradora
            aReportParamS(2) = strFechaHasta
            aReportParamS(3) = strCodEstado
            aReportParamS(4) = strCodMoneda
'            aReportParamS(5) = strTipoContraparte
'            aReportParamS(6) = Convertyyyymmdd(gstrFchDel)
            
            
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



Private Sub cboCuentas_Click()

    strCodFile = Valor_Caracter: strCodAnalitica = Valor_Caracter
    strCodBanco = Valor_Caracter: strCodCuenta = Valor_Caracter
    If cboCuentas.ListIndex < 0 Then Exit Sub
   
    strCodFile = Left(Trim(arrCuentas(cboCuentas.ListIndex)), 3)
    strCodAnalitica = Mid(Trim(arrCuentas(cboCuentas.ListIndex)), 4, 8)
    strCodBanco = Mid(Trim(arrCuentas(cboCuentas.ListIndex)), 12, 8)
    strCodCuenta = Trim(Right(arrCuentas(cboCuentas.ListIndex), 10))
   
    lblMonedaCuentaLiquidacion(0).Caption = ObtenerCodSignoMoneda(strCodMonedaPago)
    lblMonedaCuentaLiquidacion(1).Caption = ObtenerCodSignoMoneda(strCodMonedaPago)
   
    If Trim(txtTipoCambio.Text) = "" Then txtTipoCambio.Text = 0#
   
    lblMovCuenta.Caption = Round(ObtenerMontoArbitraje(CDbl(txtMonto.Text), CDbl(txtTipoCambio.Text), strCodMonedaParEvaluacion, strCodMonedaParPorDefecto), 2)
  
    If cboCuentas.ListIndex <> 0 Then
        If txtTipoCambio.Visible = True Then
            lblMovCuenta.Caption = Round(ObtenerMontoArbitraje(CDbl(txtMonto.Text), CDbl(txtTipoCambio.Text), strCodMonedaParEvaluacion, strCodMonedaParPorDefecto), 2)
        Else
            lblMovCuenta.Caption = CDbl(txtMonto.Text)
        End If
    End If
    
    Call ObtenerSaldos

    
End Sub


Private Sub cboEstado_Click()

    strCodEstado = Valor_Caracter
    If cboEstado.ListIndex < 0 Then Exit Sub
    
    strCodEstado = Trim(arrEstado(cboEstado.ListIndex))
    
    Call Buscar
    
End Sub


Private Sub cboFondo_Click()

    Dim adoRegistro As ADODB.Recordset
    
    strCodFondo = ""
    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(arrFondo(cboFondo.ListIndex))
    
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Fecha Vigente, Valor Cuota, Valor Cuota T-1, Moneda y Cantidad Inicial de Cuotas del Fondo ***
        .CommandText = "{ call up_ACSelDatosParametro(23,'" & strCodFondo & "','" & gstrCodAdministradora & "') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            dtpFechaConsulta.Value = adoRegistro("FechaCuota")
            dtpFechaContable.Value = adoRegistro("FechaCuota")

            intDiasDesplazamiento = ObtenerParametroDesplazamientoFechaTipoCambio()
            datFechaConsulta = DateAdd("d", intDiasDesplazamiento, CVDate(dtpFechaContable.Value))
            
            strCodMoneda = adoRegistro("CodMoneda")
            txtTipoCambio.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, dtpFechaContable.Value, strCodMoneda, Codigo_Moneda_Local))
            If CDbl(txtTipoCambio.Text) = 0 Then txtTipoCambio.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, DateAdd("d", -1, dtpFechaContable.Value), strCodMoneda, Codigo_Moneda_Local))
            gdblTipoCambio = CDbl(txtTipoCambio.Text)
'            gdblTipoCambio = CDbl(adoRegistro("ValorTipoCambio"))
'            txtTipoCambio.Text = gdblTipoCambio
            
            gdatFechaActual = adoRegistro("FechaCuota")
            frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
                        
        Call Buscar
        
        'lblMonto.Caption = "0"
        txtMonto.Text = "0.00"
        txtMontoOriginal.Text = "0.00"
        Call CargarCuentasBancarias
    End With
    
End Sub
Private Sub cboFormaPago_Click()

    strCodFormaPago = Valor_Caracter
    If cboFormaPago.ListIndex < 0 Then Exit Sub
    
    strCodFormaPago = Trim(arrFormaPago(cboFormaPago.ListIndex))
    
    Select Case strCodFormaPago
        Case Codigo_FormaPago_Cuenta
            txtNroDocumento.Visible = False
            lblDescrip(12).Visible = False

        Case Codigo_FormaPago_Cheque
            txtNroDocumento.Visible = True
            lblDescrip(12).Visible = True
            lblDescrip(12).Caption = "Num.Cheque"

        Case Else
            txtNroDocumento.Visible = False
            lblDescrip(12).Visible = False

    End Select
    
'    If cboFormaPago.ListIndex = 0 Then
'        txtNroDocumento.Visible = True
'        lblDescrip(12).Visible = True
'    Else
'        txtNroDocumento.Visible = False
'        lblDescrip(12).Visible = False
'    End If
    
End Sub

Private Sub cboMoneda_Click()

    strCodMonedaBusqueda = Valor_Caracter
    If cboMoneda.ListIndex < 0 Then Exit Sub
    strCodMonedaBusqueda = Trim(arrMoneda(cboMoneda.ListIndex))
    
    Call Buscar
    
End Sub

Private Sub cboMonedaPago_Click()
  
''''''''''''''''''
    Dim strCodSignoMonedaOrden  As String
    Dim strCodSignoMonedaPago   As String
  
    strCodMonedaPago = Valor_Caracter
    If cboMonedaPago.ListIndex < 0 Then Exit Sub
       
    cboCuentas.Clear
   
    strCodMonedaPago = Trim(arrMonedaPago(cboMonedaPago.ListIndex))
    strCodMonedaParEvaluacion = strCodMonedaOrden & strCodMonedaPago
   
    If strCodMonedaOrden <> strCodMonedaPago Then
        strCodMonedaParPorDefecto = ObtenerMonedaParPorDefecto(gstrCodClaseTipoCambioOperacionFondo, strCodMonedaParEvaluacion) 'SBS
    Else
        strCodMonedaParPorDefecto = strCodMonedaParEvaluacion
    End If
       
    If strCodMonedaParPorDefecto = "0000" Then strCodMonedaParPorDefecto = strCodMonedaParEvaluacion
   
    If strCodMonedaOrden <> strCodMonedaPago Then
       
        ChkMonDiferente.Value = vbChecked
        lblTipoCambio.Visible = True
        txtTipoCambio.Visible = True
        lblMonedaTC.Visible = True
   
        'lblMonedaTC.Caption = "(" & strCodSignoMonedaPago & "/" & strCodSignoMonedaOrden & ")"
        lblMonedaTC.Caption = "(" + Trim(ObtenerCodSignoMoneda(Mid(strCodMonedaParPorDefecto, 1, 2))) + "/" + Trim(ObtenerCodSignoMoneda(Mid(strCodMonedaParPorDefecto, 3, 2))) + ")"
   
        dblTipoCambioOrden = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, datFechaConsulta, Mid(strCodMonedaParPorDefecto, 1, 2), Mid(strCodMonedaParPorDefecto, 3, 2)))
        txtTipoCambio.Text = 0#
        txtTipoCambio.Text = dblTipoCambioOrden 'CStr(ObtenerTipoCambioMoneda(strCodTipoCambio, strValorCambio, datFechaConsulta, Mid(strCodMonedaParPorDefecto, 1, 2), Mid(strCodMonedaParPorDefecto, 3, 2)))
   
    Else
   
        dblTipoCambioOrden = 1
        ChkMonDiferente.Value = vbUnchecked
        lblTipoCambio.Visible = False
        txtTipoCambio.Visible = False
        txtTipoCambio.Text = "1"
        lblMonedaTC.Visible = False
    End If
 
    Call CargarCuentasBancarias
    
    
End Sub

Private Sub cboTipoDocumento_Click()

    strTipoDocumento = Valor_Caracter
    If cboTipoDocumento.ListIndex < 0 Then Exit Sub
    
    strTipoDocumento = arrTipoDocumento(cboTipoDocumento.ListIndex)

End Sub

Private Sub cmdLiquidar_Click()
    Call Modificar
End Sub





Private Sub Form_Activate()

    Call CargarReportes
    
End Sub

Private Sub Form_Deactivate()

    Call OcultarReportes
    
End Sub


Private Sub Form_Load()
    
    Dim intRegistro As Integer
    
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

    Dim intRegistro  As Integer
    
    '*** Fondos ***
    strSQL = "{ call up_ACSelDatosParametro(74,'" & gstrCodAdministradora & "','" & gstrCodFondoContable & "') }"
    CargarControlLista strSQL, cboFondo, arrFondo(), Valor_Caracter
    
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
    
    '*** Estado Orden Cobro/Pago ***
    strSQL = "SELECT CodParametro CODIGO, DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='ESTCAJ' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboEstado, arrEstado(), Valor_Caracter
    
    intRegistro = ObtenerItemLista(arrEstado(), Estado_Caja_NoConfirmado)
    If intRegistro >= 0 Then cboEstado.ListIndex = intRegistro
    
    '*** Moneda Pago***
    strSQL = "{ call up_ACSelDatos(2) }"
    CargarControlLista strSQL, cboMonedaPago, arrMonedaPago(), Sel_Defecto
    '*** Moneda ***
    CargarControlLista strSQL, cboMoneda, arrMoneda(), Sel_Todos
    If cboMoneda.ListCount > 0 Then cboMoneda.ListIndex = 0
    
    '*** Forma de Pago ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='MEDPAG' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboFormaPago, arrFormaPago(), Sel_Defecto
    
    '*** Tipo de Comprobante Sunat ***
    strSQL = "SELECT CodTipoComprobantePago CODIGO,DescripTipoComprobantePago DESCRIP From TipoComprobantePago ORDER BY DescripTipoComprobantePago"
    CargarControlLista strSQL, cboTipoDocumento, arrTipoDocumento(), Sel_Defecto
    
    '*** Tipo Cuenta ***
'    strSQL = "SELECT CodParametro CODIGO, DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='CTAFON' ORDER BY DescripParametro"
'    CargarControlLista strSQL, cboTipoCuenta, arrTipoCuenta(), Sel_Defecto
    
'    strSQL = "SELECT CodParametro CODIGO, DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='FORPAG' ORDER BY DescripParametro"
'    CargarControlLista strSQL, cboFormaPago, arrFormaPago(), ""
    
            
End Sub
Private Sub InicializarValores()
    
    '*** Valores Iniciales ***
    strEstado = Reg_Defecto
    tabOrdenCobroPago.Tab = 0
    txtNroDocumento.Text = Reg_Defecto
    
    indSortDesc = False
    
    If Me.Tag = "06" Then 'clientes
        Me.Caption = "Ordenes de Cobro y Pagos (COMITENTES)"
    End If
    
    If Me.Tag = "03" Then 'brokers
        Me.Caption = "Ordenes de Cobro y Pagos (BROKERS)"
    End If
   
    If Me.Tag = "04" Then 'clientes
        Me.Caption = "Ordenes de Cobro y Pagos (PROVEEDORES)"
    End If
   
    If Me.Tag = "05" Then 'clientes
        Me.Caption = "Ordenes de Cobro y Pagos (CLIENTES)"
    End If
   
    tabOrdenCobroPago.TabEnabled(1) = False
    
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgConsulta.Columns(0).Width = tdgConsulta.Width * 0.01 * 9
    tdgConsulta.Columns(1).Width = tdgConsulta.Width * 0.01 * 10
    tdgConsulta.Columns(2).Width = tdgConsulta.Width * 0.01 * 26
    tdgConsulta.Columns(4).Width = tdgConsulta.Width * 0.01 * 6
    tdgConsulta.Columns(5).Width = tdgConsulta.Width * 0.01 * 18
    tdgConsulta.Columns(6).Width = tdgConsulta.Width * 0.01 * 16
    tdgConsulta.Columns(3).Width = tdgConsulta.Width * 0.01 * 26
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    
End Sub
Private Sub Form_Unload(Cancel As Integer)

    Call OcultarReportes
    Set frmLiquidacionOrdenCaja = Nothing
    
End Sub









Private Sub lblMonto_Change()

    Call FormatoMillarEtiqueta(lblMonto, Decimales_Monto)
    
End Sub

Private Sub lblMontoContable_Change()

    Call FormatoMillarEtiqueta(lblMontoContable, Decimales_Monto)

End Sub

Private Sub lblMovCuenta_Change()

    Call FormatoMillarEtiqueta(lblMovCuenta, Decimales_Monto)
    
End Sub


Private Sub lblSaldoCuenta_Change()

    Call FormatoMillarEtiqueta(lblSaldoCuenta, Decimales_Monto)
    
End Sub

Private Sub tabOrdenCobroPago_Click(PreviousTab As Integer)

    Select Case tabOrdenCobroPago.Tab
        Case 1
            If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vQuery)
            If strEstado = Reg_Defecto Then tabOrdenCobroPago.Tab = 0
                                
    End Select
    
End Sub

Private Sub tdgConsulta_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)

    If tdgConsulta.Columns(ColIndex).DataField = "MontoOrdenCobroSaldo" Or tdgConsulta.Columns(ColIndex).DataField = "MontoOrdenLiquidacion" Then

       Call DarFormatoValor(Value, Decimales_Monto)

    End If
    
'    If ColIndex = 1 Then
'
'        For i = 1 To adoConsulta.Recordset.RecordCount
'        tdgConsulta.Columns(i).Order
'        Next
'
'    End If
    
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

    Call OrdenarDBGrid(ColIndex, adoRegistro, tdgConsulta)
    
    numColindex = ColIndex
    
    '****
    strPrevColumTDB = strColNameTDB
    '***
    
End Sub

Private Sub tdgConsulta_SelChange(Cancel As Integer)
    

    Dim dblMonto                    As Double, dblMontoAcumulado             As Double
    Dim dblMontoContable            As Double, dblMontoContableAcumulado    As Double
    Dim intRegistro                 As Integer, intContador                 As Integer
    Dim intNumGastoSel              As Long, dblMontoLiq                    As Double
    Dim dblMontoAcumuladoLiq        As Double
    Dim intNum                      As Integer
    
    If tdgConsulta.SelBookmarks.Count < 1 Then Exit Sub
  
    adoRegistro.MoveFirst
    adoRegistro.Move CLng(tdgConsulta.SelBookmarks.Count - 1), 0
    
    txtTotalSeleccionado.Text = "0"
   
    
    intContador = tdgConsulta.SelBookmarks.Count - 1

    adoRegistro.MoveFirst
    
    dblMontoAcumuladoLiq = 0
    
    'Si la seleccion es multiple, se puede ingresar el tipo y numero de documento
    If intContador > 0 Then
        Call HabilitarSeleccionMultiple(True)
    Else
        Call HabilitarSeleccionMultiple(False)
    End If
    
    For intRegistro = 0 To intContador
        adoRegistro.MoveFirst
        
        'tdgConsulta.Row = tdgConsulta.SelBookmarks(intRegistro) - 1
        
        If indSortDesc Then
            intNum = ObtenerBookMarkInverso(tdgConsulta.SelBookmarks(intRegistro), adoRegistro.RecordCount)
        Else
            intNum = tdgConsulta.SelBookmarks(intRegistro)
        End If
        
        adoRegistro.Move CLng(intNum - 1), 0
        tdgConsulta.Refresh
        
        If intRegistro = 0 Then
            strCodMonedaOrden = tdgConsulta.Columns("CodMoneda")
            lblMonedaOrden.Caption = ObtenerCodSignoMoneda(strCodMonedaOrden)
            lblMonedaOrdenOriginal.Caption = ObtenerCodSignoMoneda(strCodMonedaOrden)
            'gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion

            If strCodMonedaOrden <> Codigo_Moneda_Local Then
        
                dblTipoCambioOrden = ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, dtpFechaContable.Value, strCodMonedaOrden, Codigo_Moneda_Local)
                    
                If dblTipoCambioOrden = 0 Then
                    MsgBox "No existe tipo de cambio para procesar la(s) orden(es) seleccionada(s)", vbCritical, Me.Caption
                    tdgConsulta.ReBind
                    tdgConsulta.Row = 0
                    Exit Sub
                End If
                        
            Else
                dblTipoCambioOrden = 1
            End If
            
        End If
        
        If tdgConsulta.Columns("CodMoneda") <> strCodMonedaOrden Then
            MsgBox "No se pueden confirmar masivamente ordenes con distinta moneda", vbCritical, Me.Caption
            tdgConsulta.ReBind
            tdgConsulta.Row = 0
            Exit Sub
        End If

'        If tdgConsulta.Columns("NumGasto") <> intNumGastoSel Then
'            MsgBox "No se pueden confirmar masivamente ordenes de gastos y de distinto origen a la vez", vbCritical, Me.Caption
'            tdgConsulta.ReBind
'            tdgConsulta.Row = 0
'            Exit Sub
'        End If
        strDescripOrden = Trim(CStr(tdgConsulta.Columns("DescripOrden")))
        
        dblMonto = CDbl(tdgConsulta.Columns("MontoOrdenCobroSaldo"))
        
        dblMontoContable = Round(dblMonto * dblTipoCambioOrden, 2)
        
        dblMontoAcumulado = dblMontoAcumulado + dblMonto
        dblMontoContableAcumulado = dblMontoContableAcumulado + dblMontoContable
              
        lblMonto.Caption = CStr(dblMontoAcumulado)
        txtMonto.Text = CStr(dblMontoAcumulado)
        txtMontoOriginal.Text = CStr(dblMontoAcumulado)
        lblMontoContable.Caption = CStr(dblMontoContableAcumulado)
        
        dblMontoLiq = dblMonto
        
        dblMontoAcumuladoLiq = dblMontoAcumuladoLiq + dblMontoLiq
        
    Next
    
    If strCodMonedaOrden = Codigo_Moneda_Local Then
        txtTotalSeleccionado.Text = CStr(dblMontoAcumuladoLiq)
    Else
        txtTotalSeleccionadoME.Text = CStr(dblMontoAcumuladoLiq)
    End If

End Sub

Private Sub txtMonto_Change()

'    Dim dblMonto                    As Double, dblMontoAcumulado             As Double
'    Dim dblMontoContable            As Double, dblMontoContableAcumulado    As Double
'    Dim intRegistro                 As Integer, intContador                 As Integer
'    Dim intNumGastoSel              As Long, dblMontoLiq                    As Double
'    Dim dblMontoAcumuladoLiq        As Double
'
'
'    If Trim(txtMonto.Text) = Valor_Caracter Then
'        dblMonto = 0
'    Else
'        dblMonto = CDbl(txtMonto.Text)  'CDbl(tdgConsulta.Columns("MontoOrdenCobroPago"))
'    End If
'    dblMontoContable = Round(dblMonto * dblTipoCambioOrden, 2)
'
'    dblMontoAcumulado = dblMonto
'    dblMontoContableAcumulado = dblMontoContable
'
'    txtMonto.Text = CStr(dblMontoAcumulado)
'
'    'If Len(txtMonto.Text) > 0 Then txtMonto.SelStart = Len(txtMonto.Text)
'
'    lblMontoContable.Caption = CStr(dblMontoContableAcumulado)
'
'    dblMontoAcumuladoLiq = dblMontoAcumuladoLiq + dblMontoLiq
'
'    Call cboCuentas_Click


End Sub

Private Sub txtMonto_LostFocus()

    Dim dblMonto                    As Double, dblMontoAcumulado             As Double
    Dim dblMontoContable            As Double, dblMontoContableAcumulado    As Double
    Dim intRegistro                 As Integer, intContador                 As Integer
    Dim intNumGastoSel              As Long, dblMontoLiq                    As Double
    Dim dblMontoAcumuladoLiq        As Double
    
    
    If Trim(txtMonto.Text) = Valor_Caracter Then
        dblMonto = 0
    Else
        dblMonto = CDbl(txtMonto.Text)  'CDbl(tdgConsulta.Columns("MontoOrdenCobroPago"))
    End If
    dblMontoContable = Round(dblMonto * dblTipoCambioOrden, 2)

    dblMontoAcumulado = dblMonto
    dblMontoContableAcumulado = dblMontoContable

    txtMonto.Text = CStr(dblMontoAcumulado)
        
    lblMontoContable.Caption = CStr(dblMontoContableAcumulado)

    dblMontoAcumuladoLiq = dblMontoAcumuladoLiq + dblMontoLiq

    Call cboCuentas_Click
End Sub

Private Sub txtTipoCambio_Change()

    Call FormatoCajaTexto(txtTipoCambio, Decimales_TipoCambio)
    
    lblMovCuenta.Caption = Round(ObtenerMontoArbitraje(CDbl(txtMonto.Text), CDbl(txtTipoCambio.Text), strCodMonedaParEvaluacion, strCodMonedaParPorDefecto), 2)
  
    If cboCuentas.ListIndex <> 0 Then
        If txtTipoCambio.Visible = True Then
            lblMovCuenta.Caption = Round(ObtenerMontoArbitraje(CDbl(txtMonto.Text), CDbl(txtTipoCambio.Text), strCodMonedaParEvaluacion, strCodMonedaParPorDefecto), 2)
        Else
            lblMovCuenta.Caption = CDbl(txtMonto.Text)
        End If
    End If
    
End Sub

Private Sub txtTipoCambio_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtTipoCambio, Decimales_TipoCambio)
    
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

Public Sub HabilitarSeleccionMultiple(blnHabilita As Boolean)
    
    If blnHabilita Then
        strIndSeleccionMultiple = Valor_Indicador
    Else
        strIndSeleccionMultiple = Valor_Caracter
    End If
        
    txtMonto.Locked = blnHabilita
    cboTipoDocumento.Locked = Not blnHabilita
    txtNumDocumento.Locked = Not blnHabilita
 
End Sub

Private Function ObtenerBookMarkInverso(ByVal intNum As Integer, intRecordCount) As Integer
    
    Dim intReturn As Integer
    
    intReturn = (intRecordCount - intNum) + 1
    
    ObtenerBookMarkInverso = intReturn
    
End Function

Private Sub txtTotalSeleccionadoME_Change()

    Call FormatoCajaTexto(txtTotalSeleccionadoME, Decimales_Monto)

End Sub
