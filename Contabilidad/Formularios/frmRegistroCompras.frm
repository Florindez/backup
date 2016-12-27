VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{830D5A06-9B70-4F8C-98B6-7A19C4E7760B}#1.0#0"; "TAMControls.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmRegistroCompras 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comprobantes de Pago"
   ClientHeight    =   9030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9030
   ScaleWidth      =   10755
   Begin MSAdodcLib.Adodc adoConsulta 
      Height          =   330
      Left            =   6960
      Top             =   8130
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin TabDlg.SSTab tabRegistroCompras 
      Height          =   7800
      Left            =   0
      TabIndex        =   25
      Top             =   120
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   13758
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Compras"
      TabPicture(0)   =   "frmRegistroCompras.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "tdgPendientes"
      Tab(0).Control(1)=   "tdgConsulta"
      Tab(0).Control(2)=   "fraCompras(0)"
      Tab(0).Control(3)=   "ucBotonNavegacion1"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Definición del Registro"
      TabPicture(1)   =   "frmRegistroCompras.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraCompras(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Definición de Obligación"
      TabPicture(2)   =   "frmRegistroCompras.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblMensaje"
      Tab(2).Control(1)=   "cmdAccion"
      Tab(2).Control(2)=   "fraCompras(2)"
      Tab(2).Control(3)=   "cmdContabilizar"
      Tab(2).ControlCount=   4
      Begin VB.CommandButton cmdContabilizar 
         Caption         =   "Contabilizar"
         Enabled         =   0   'False
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
         Left            =   -69360
         Picture         =   "frmRegistroCompras.frx":0054
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   6270
         Width           =   1215
      End
      Begin TAMControls.ucBotonNavegacion ucBotonNavegacion1 
         Height          =   30
         Left            =   -69450
         TabIndex        =   70
         Top             =   5220
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   53
      End
      Begin VB.Frame fraCompras 
         Caption         =   "Definición de Obligación"
         Height          =   5655
         Index           =   2
         Left            =   -74760
         TabIndex        =   51
         Top             =   480
         Width           =   9975
         Begin VB.ComboBox cboTipoValorCambio 
            Height          =   315
            Left            =   2400
            Style           =   2  'Dropdown List
            TabIndex        =   68
            Top             =   4080
            Width           =   2295
         End
         Begin VB.TextBox txtMontoDetraccion 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   7200
            TabIndex        =   22
            Top             =   2520
            Width           =   2295
         End
         Begin VB.TextBox txtTipoCambioPago 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   2400
            TabIndex        =   23
            Top             =   5040
            Width           =   2295
         End
         Begin VB.TextBox txtMontoUnico 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   7200
            TabIndex        =   20
            Top             =   1080
            Width           =   2295
         End
         Begin VB.ComboBox cboMonedaDetraccion 
            Height          =   315
            Left            =   2400
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   2520
            Width           =   2295
         End
         Begin VB.ComboBox cboMonedaUnico 
            Height          =   315
            Left            =   2400
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   1080
            Width           =   2295
         End
         Begin MSComCtl2.DTPicker dtpFechaPago 
            Height          =   345
            Left            =   2400
            TabIndex        =   18
            Top             =   480
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   609
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
            Format          =   50397185
            CurrentDate     =   39042
         End
         Begin MSComCtl2.DTPicker dtpFechaTipoCambioPago 
            Height          =   315
            Left            =   7200
            TabIndex        =   71
            Top             =   4080
            Width           =   2295
            _ExtentX        =   4048
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
            Format          =   50397185
            CurrentDate     =   39042
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Tipo Cambio"
            Height          =   195
            Index           =   32
            Left            =   5460
            TabIndex        =   72
            Top             =   4110
            Width           =   1380
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Valor Cambio Sunat"
            Height          =   195
            Index           =   23
            Left            =   480
            TabIndex        =   67
            Top             =   4080
            Width           =   1395
         End
         Begin VB.Label lblMontoTotal 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7200
            TabIndex        =   24
            Top             =   5040
            Width           =   2295
         End
         Begin VB.Line Line2 
            BorderColor     =   &H8000000C&
            BorderStyle     =   6  'Inside Solid
            X1              =   480
            X2              =   9480
            Y1              =   3480
            Y2              =   3480
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Detracción"
            Height          =   195
            Index           =   29
            Left            =   480
            TabIndex        =   59
            Top             =   1920
            Width           =   780
         End
         Begin VB.Line Line1 
            BorderColor     =   &H8000000C&
            BorderStyle     =   6  'Inside Solid
            X1              =   1680
            X2              =   9480
            Y1              =   2040
            Y2              =   2040
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto Total"
            Height          =   195
            Index           =   28
            Left            =   5490
            TabIndex        =   58
            Top             =   5040
            Width           =   855
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto"
            Height          =   195
            Index           =   30
            Left            =   5400
            TabIndex        =   57
            Top             =   2520
            Width           =   450
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Index           =   31
            Left            =   480
            TabIndex        =   56
            Top             =   2640
            Width           =   585
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Cambio"
            Height          =   195
            Index           =   27
            Left            =   480
            TabIndex        =   55
            Top             =   5040
            Width           =   1110
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto"
            Height          =   195
            Index           =   26
            Left            =   5400
            TabIndex        =   54
            Top             =   1080
            Width           =   450
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Index           =   25
            Left            =   480
            TabIndex        =   53
            Top             =   1080
            Width           =   585
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            Height          =   195
            Index           =   22
            Left            =   480
            TabIndex        =   52
            Top             =   495
            Width           =   450
         End
      End
      Begin VB.Frame fraCompras 
         Caption         =   "Definición del Registro"
         Height          =   6765
         Index           =   1
         Left            =   360
         TabIndex        =   31
         Top             =   600
         Width           =   9975
         Begin VB.ComboBox cboGasto 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   78
            Top             =   1540
            Width           =   6975
         End
         Begin VB.TextBox txtSerieComprobante 
            Height          =   315
            Left            =   7200
            MaxLength       =   3
            TabIndex        =   69
            Top             =   1144
            Width           =   615
         End
         Begin VB.ComboBox cboDetraccion 
            Height          =   315
            Left            =   7200
            Style           =   2  'Dropdown List
            TabIndex        =   66
            Tag             =   "0"
            Top             =   5550
            Width           =   2295
         End
         Begin VB.TextBox txtPeriodoFiscal 
            Height          =   315
            Left            =   7200
            TabIndex        =   14
            Top             =   4065
            Width           =   2295
         End
         Begin VB.ComboBox cboCreditoFiscal 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmRegistroCompras.frx":0640
            Left            =   2520
            List            =   "frmRegistroCompras.frx":0642
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   4065
            Width           =   2295
         End
         Begin VB.ComboBox cboAfectacion 
            Height          =   315
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   3660
            Width           =   2295
         End
         Begin VB.CommandButton cmdProveedor 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   9150
            TabIndex        =   9
            ToolTipText     =   "Buscar Proveedor"
            Top             =   1928
            Width           =   345
         End
         Begin VB.ComboBox cboMoneda 
            Enabled         =   0   'False
            Height          =   315
            Left            =   7200
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   5175
            Width           =   2295
         End
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   2520
            TabIndex        =   17
            Top             =   6180
            Width           =   2295
         End
         Begin VB.TextBox txtIgv 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   2520
            TabIndex        =   16
            Top             =   5460
            Width           =   2295
         End
         Begin VB.TextBox txtDescripcion 
            Height          =   315
            Left            =   2520
            MaxLength       =   30
            TabIndex        =   11
            Top             =   3045
            Width           =   6975
         End
         Begin VB.ComboBox cboTipoComprobante 
            Height          =   315
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   737
            Width           =   6975
         End
         Begin VB.TextBox txtNumComprobante 
            Height          =   315
            Left            =   7920
            MaxLength       =   10
            TabIndex        =   7
            Top             =   1144
            Width           =   1575
         End
         Begin MSComCtl2.DTPicker dtpFechaRegistro 
            Height          =   315
            Left            =   7200
            TabIndex        =   4
            Top             =   360
            Width           =   2295
            _ExtentX        =   4048
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
            Format          =   50397185
            CurrentDate     =   39042
         End
         Begin MSComCtl2.DTPicker dtpFechaComprobante 
            Height          =   315
            Left            =   2520
            TabIndex        =   6
            Top             =   1144
            Width           =   2295
            _ExtentX        =   4048
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
            Format          =   50397185
            CurrentDate     =   39042
         End
         Begin TAMControls.TAMTextBox txtMontoNoGravado 
            Height          =   315
            Left            =   2520
            TabIndex        =   75
            Top             =   5820
            Width           =   2295
            _ExtentX        =   4048
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
            Container       =   "frmRegistroCompras.frx":0644
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
         Begin TAMControls.TAMTextBox txtSubTotal 
            Height          =   315
            Left            =   2520
            TabIndex        =   76
            Top             =   5130
            Width           =   2295
            _ExtentX        =   4048
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
            Container       =   "frmRegistroCompras.frx":0660
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
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto Gasto"
            Height          =   195
            Index           =   34
            Left            =   480
            TabIndex        =   77
            Top             =   4800
            Width           =   915
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto No Gravado"
            Height          =   195
            Index           =   33
            Left            =   480
            TabIndex        =   74
            Top             =   5865
            Width           =   1365
         End
         Begin VB.Label lblTipoDocID 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2520
            TabIndex        =   73
            Top             =   2310
            Width           =   4215
         End
         Begin VB.Label lblNumDocID 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   6840
            TabIndex        =   65
            Top             =   2310
            Width           =   2655
         End
         Begin VB.Line Line4 
            BorderColor     =   &H8000000C&
            BorderStyle     =   6  'Inside Solid
            X1              =   480
            X2              =   9480
            Y1              =   4480
            Y2              =   4480
         End
         Begin VB.Line Line3 
            BorderColor     =   &H8000000C&
            BorderStyle     =   6  'Inside Solid
            X1              =   480
            X2              =   9480
            Y1              =   3480
            Y2              =   3480
         End
         Begin VB.Label lblAnalitica 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   7200
            TabIndex        =   63
            Top             =   3660
            Width           =   2295
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Analítica"
            Height          =   195
            Index           =   24
            Left            =   5400
            TabIndex        =   62
            Top             =   3675
            Width           =   630
         End
         Begin VB.Label lblMontoGasto 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2520
            TabIndex        =   61
            Top             =   4770
            Width           =   2295
         End
         Begin VB.Label lblCodProveedor 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   5370
            TabIndex        =   60
            Top             =   6180
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Periodo Registro Crédito Fiscal"
            Height          =   435
            Index           =   21
            Left            =   5400
            TabIndex        =   50
            Top             =   4080
            Width           =   1455
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Crédito Fiscal"
            Height          =   195
            Index           =   20
            Left            =   480
            TabIndex        =   49
            Top             =   4080
            Width           =   945
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Impuesto"
            Height          =   195
            Index           =   19
            Left            =   480
            TabIndex        =   48
            Top             =   3675
            Width           =   645
         End
         Begin VB.Label lblDireccion 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2520
            TabIndex        =   10
            Top             =   2670
            Width           =   6960
         End
         Begin VB.Label lblProveedor 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2520
            TabIndex        =   8
            Top             =   1928
            Width           =   6600
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Comprobante"
            Height          =   195
            Index           =   18
            Left            =   480
            TabIndex        =   47
            Top             =   1164
            Width           =   1440
         End
         Begin VB.Label lblNumSecuencial 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2520
            TabIndex        =   3
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            Height          =   195
            Index           =   17
            Left            =   5400
            TabIndex        =   46
            Top             =   375
            Width           =   450
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Retención y/o Detracción"
            Height          =   435
            Index           =   16
            Left            =   5400
            TabIndex        =   45
            Top             =   5580
            Width           =   1725
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Gasto Relacionado"
            Height          =   195
            Index           =   14
            Left            =   480
            TabIndex        =   43
            Top             =   1541
            Width           =   1365
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Documento ID"
            Height          =   195
            Index           =   13
            Left            =   480
            TabIndex        =   42
            Top             =   2310
            Width           =   1035
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Index           =   12
            Left            =   5400
            TabIndex        =   41
            Top             =   5205
            Width           =   585
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Comprobante"
            Height          =   195
            Index           =   11
            Left            =   480
            TabIndex        =   40
            Top             =   757
            Width           =   1305
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Precio de Venta"
            Height          =   195
            Index           =   10
            Left            =   480
            TabIndex        =   39
            Top             =   6195
            Width           =   1140
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "IGV"
            Height          =   195
            Index           =   9
            Left            =   480
            TabIndex        =   38
            Top             =   5505
            Width           =   270
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Dirección"
            Height          =   195
            Index           =   8
            Left            =   480
            TabIndex        =   37
            Top             =   2685
            Width           =   675
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Valor de Venta"
            Height          =   195
            Index           =   7
            Left            =   480
            TabIndex        =   36
            Top             =   5145
            Width           =   1050
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Descripción"
            Height          =   195
            Index           =   6
            Left            =   480
            TabIndex        =   35
            Top             =   3060
            Width           =   840
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor"
            Height          =   195
            Index           =   5
            Left            =   480
            TabIndex        =   34
            Top             =   1948
            Width           =   735
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Num. Registro"
            Height          =   195
            Index           =   4
            Left            =   480
            TabIndex        =   33
            Top             =   380
            Width           =   1005
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Num. Comprobante"
            Height          =   195
            Index           =   3
            Left            =   5400
            TabIndex        =   32
            Top             =   1170
            Width           =   1365
         End
      End
      Begin VB.Frame fraCompras 
         Caption         =   "Criterios de búsqueda"
         Height          =   1335
         Index           =   0
         Left            =   -74640
         TabIndex        =   27
         Top             =   570
         Width           =   9975
         Begin MSComCtl2.DTPicker dtpFechaDesde 
            Height          =   285
            Left            =   3600
            TabIndex        =   1
            Top             =   840
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   503
            _Version        =   393216
            Format          =   50397185
            CurrentDate     =   39042
         End
         Begin VB.ComboBox cboFondo 
            Height          =   315
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   360
            Width           =   6255
         End
         Begin MSComCtl2.DTPicker dtpFechaHasta 
            Height          =   285
            Left            =   7200
            TabIndex        =   2
            Top             =   840
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   503
            _Version        =   393216
            Format          =   50397185
            CurrentDate     =   39042
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Registro :"
            Height          =   195
            Index           =   15
            Left            =   840
            TabIndex        =   44
            Top             =   840
            Width           =   1395
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            Height          =   195
            Index           =   2
            Left            =   6000
            TabIndex        =   30
            Top             =   840
            Width           =   420
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            Height          =   195
            Index           =   1
            Left            =   2520
            TabIndex        =   29
            Top             =   840
            Width           =   465
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fondo"
            Height          =   195
            Index           =   0
            Left            =   840
            TabIndex        =   28
            Top             =   360
            Width           =   450
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmRegistroCompras.frx":067C
         Height          =   2535
         Left            =   -74640
         OleObjectBlob   =   "frmRegistroCompras.frx":0696
         TabIndex        =   26
         Top             =   2100
         Width           =   9975
      End
      Begin TrueOleDBGrid60.TDBGrid tdgPendientes 
         Bindings        =   "frmRegistroCompras.frx":5912
         Height          =   2655
         Left            =   -74640
         OleObjectBlob   =   "frmRegistroCompras.frx":592E
         TabIndex        =   64
         Top             =   4830
         Width           =   9975
      End
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   -67830
         TabIndex        =   82
         Top             =   6270
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
      Begin VB.Label lblMensaje 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   -74400
         TabIndex        =   80
         Top             =   6480
         Width           =   4335
      End
   End
   Begin MSAdodcLib.Adodc adoPendientes 
      Height          =   330
      Left            =   240
      Top             =   8130
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin TAMControls2.ucBotonEdicion2 cmdOpcion 
      Height          =   735
      Left            =   360
      TabIndex        =   81
      Top             =   8040
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   1296
      Buttons         =   5
      Caption0        =   "&Nuevo"
      Tag0            =   "0"
      ToolTipText0    =   "Nuevo"
      Caption1        =   "&Modificar"
      Tag1            =   "1"
      ToolTipText1    =   "Modificar"
      Caption2        =   "&Anular"
      Tag2            =   "4"
      ToolTipText2    =   "Anular"
      Caption3        =   "&Buscar"
      Tag3            =   "5"
      ToolTipText3    =   "Buscar"
      Caption4        =   "&Imprimir"
      Tag4            =   "6"
      ToolTipText4    =   "Imprimir"
      UserControlWidth=   7200
   End
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   9000
      TabIndex        =   83
      Top             =   8040
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
      Caption0        =   "&Salir"
      Tag0            =   "9"
      ToolTipText0    =   "Salir"
      UserControlWidth=   1200
   End
End
Attribute VB_Name = "frmRegistroCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' /* HMC*/
Dim X As New XArrayDBObject.XArrayDB ' XArrayDB
' /*    */

Dim arrFondo()                  As String, arrMoneda()                  As String
Dim arrTipoComprobante()        As String, arrMonedaUnico()             As String
Dim arrMonedaDetraccion()       As String, arrCuentaFondoUnico()        As String
Dim arrCuentaFondoDetraccion()  As String, arrAfectacion()              As String
Dim arrCreditoFiscal()          As String, arrFormaPagoUnico()          As String
Dim arrFormaPagoDetraccion()    As String, arrGasto()                   As String
Dim arrDetraccion()             As String, arrTipoValorCambio()         As String

Dim strCodFondo                 As String, strCodMoneda                 As String
Dim strCodTipoComprobante       As String, strCodMonedaUnico            As String
Dim strCodMonedaDetraccion      As String, strCodCuentaFondoUnico       As String
Dim strCodCuentaFondoDetraccion As String, strCodAfectacion             As String
Dim strCodCreditoFiscal         As String, strCodFormaPagoUnico         As String
Dim strCodFormaPagoDetraccion   As String, strCodFileUnico              As String
Dim strCodAnaliticaUnico        As String, strCodBancoUnico             As String
Dim strCodCuentaUnico           As String, strCodFileDetraccion         As String
Dim strCodAnaliticaDetraccion   As String, strCodBancoDetraccion        As String
Dim strCodCuentaDetraccion      As String, strCodGasto                  As String
Dim strIndDetraccion            As String, strCodAnalitica              As String
Dim strCodDetalleGasto          As String, strDetraccionSiNo            As String
Dim strIndImpuesto              As String, strIndRetencion              As String
Dim strCodValorTipoCambio       As String, strCodTipoGasto              As String
Dim strCodFile                  As String, strCodAplicacionDevengo      As String
Dim strEstado                   As String, strSQL                       As String
Dim adoConsulta1 As ADODB.Recordset
Dim adoPendientes1 As ADODB.Recordset
Dim indSortAsc                  As Boolean, indSortDesc                 As Boolean

Private Sub Calculos()

    Dim intRegistro As Integer
    
    If Trim(txtSubTotal.Text) = Valor_Caracter Or Trim(txtIgv.Text) = Valor_Caracter Or Trim(txtTotal.Text) = Valor_Caracter Then Exit Sub
    
    Call cboTipoValorCambio_Click
    
    If strCodAfectacion = Codigo_Afecto Then
        If strIndImpuesto = Valor_Indicador Then
            'txtSubTotal.Text = lblMontoGasto.Caption
            txtIgv.Text = CStr(Round(CDbl(txtSubTotal.Text) * gdblTasaIgv, 2))
            txtTotal.Text = CCur(txtSubTotal.Text) + CCur(txtIgv.Text) + CCur(txtMontoNoGravado.Value)
        ElseIf strIndRetencion = Valor_Indicador Then
            'txtSubTotal.Text = lblMontoGasto.Caption
            txtIgv.Text = CStr(CCur(txtSubTotal.Text) * gdblTasaRetencion)
            txtTotal.Text = CStr(CCur(txtSubTotal.Text) - CCur(txtIgv.Text))
        Else
            'txtSubTotal.Text = lblMontoGasto.Caption
            txtIgv.Text = "0"
            txtTotal.Text = txtSubTotal.Text
        End If
    Else
        If strIndImpuesto = Valor_Indicador Then
            'txtSubTotal.Text = lblMontoGasto.Caption 'txtTotal.Text
            txtTotal.Text = CCur(txtSubTotal.Text) + CCur(txtMontoNoGravado.Value)
        ElseIf strIndRetencion = Valor_Indicador Then
            'txtSubTotal.Text = lblMontoGasto.Caption
            txtTotal.Text = CCur(txtSubTotal.Text) + CCur(txtMontoNoGravado.Value)
        Else
            'txtSubTotal.Text = lblMontoGasto.Caption
            txtTotal.Text = CCur(txtSubTotal.Text) + CCur(txtMontoNoGravado.Value)
        End If
        txtIgv.Text = "0"
    End If
    
    
    If strDetraccionSiNo = Codigo_Respuesta_Si Then
        If strIndImpuesto = Valor_Indicador Then
            If strCodMoneda <> Codigo_Moneda_Local Then
                txtMontoDetraccion.Text = CStr(Round(CCur(txtTotal.Text) * gdblTasaDetraccion * CDbl(txtTipoCambioPago.Text), 2))
            Else
                txtMontoDetraccion.Text = CStr(Round(CCur(txtTotal.Text) * gdblTasaDetraccion, 2))
            End If
            txtMontoUnico.Text = CStr(CCur(txtTotal.Text) - (CCur(txtTotal.Text) * gdblTasaDetraccion))
            
            intRegistro = ObtenerItemLista(arrMonedaDetraccion(), Codigo_Moneda_Local)
            If intRegistro >= 0 Then cboMonedaDetraccion.ListIndex = intRegistro
            
        ElseIf strIndRetencion = Valor_Indicador Then
            txtMontoDetraccion.Text = CStr(Round(CCur(txtSubTotal.Text) * gdblTasaRetencion, 0))
            txtMontoUnico.Text = CStr(CCur(txtSubTotal.Text) - CCur(txtSubTotal.Text) * gdblTasaRetencion)
        
            intRegistro = ObtenerItemLista(arrMonedaDetraccion(), strCodMoneda)
            If intRegistro >= 0 Then cboMonedaDetraccion.ListIndex = intRegistro
        
        End If
        lblMontoTotal.Caption = CStr(CCur(txtTotal.Text)) 'CStr(CCur(txtMontoUnico.Text) + CCur(txtMontoDetraccion.Text))
    Else
        txtMontoDetraccion.Text = "0"
        If strIndImpuesto = Valor_Indicador Then
            txtMontoUnico.Text = CStr(CCur(txtTotal.Text))
        ElseIf strIndRetencion = Valor_Indicador Then
            txtMontoUnico.Text = CStr(CCur(txtSubTotal.Text)) + CCur(txtMontoNoGravado.Value)
        Else
            txtMontoUnico.Text = CStr(CCur(txtSubTotal.Text)) + CCur(txtMontoNoGravado.Value)
        End If
        lblMontoTotal.Caption = CStr(CCur(txtMontoUnico.Text))
    End If
    

    
    
End Sub
Private Sub CalculosPago()

    If Trim(txtMontoUnico.Text) = Valor_Caracter Or Trim(txtMontoDetraccion.Text) = Valor_Caracter Then Exit Sub
                
    lblMontoTotal.Caption = CStr(CCur(txtMontoUnico.Text) + CCur(txtMontoDetraccion.Text))
    
End Sub
Private Sub CargarPendientes()

 Set adoPendientes1 = New ADODB.Recordset


    strSQL = "SELECT FG.CodCuenta,NumGasto,CodFile,CodAnalitica,DescripCuenta,DescripGasto,CodTipoGasto,MontoGasto,MontoDevengo,FechaDefinicion " & _
        "FROM FondoGasto FG JOIN FondoConceptoGasto FCG ON(FCG.CodCuenta=FG.CodCuenta AND FCG.CodAdministradora=FG.CodAdministradora AND FCG.CodFondo=FG.CodFondo) " & _
        "JOIN PlanContable PCG ON(PCG.CodCuenta=FCG.CodCuenta) " & _
        "WHERE FG.CodFondo='" & strCodFondo & "' AND FG.CodAdministradora='" & gstrCodAdministradora & "' AND IndConfirma='' AND IndVigente = 'X' /*AND FG.IndFinMes='X'*/" & _
    "UNION " & _
             "SELECT FG.CodCuenta,NumGasto,CodFile,CodAnalitica,DescripCuenta,DescripGasto,CodTipoGasto,MontoGasto,MontoDevengo,FechaDefinicion " & _
        "FROM FondoGasto FG JOIN FondoConceptoActivoFijo FCG ON(FCG.CodCuenta=FG.CodCuenta AND FCG.CodAdministradora=FG.CodAdministradora AND FCG.CodFondo=FG.CodFondo) " & _
        "JOIN PlanContable PCG ON (PCG.CodCuenta=FCG.CodCuenta) " & _
        "WHERE FG.CodFondo='" & strCodFondo & "' AND FG.CodAdministradora='" & gstrCodAdministradora & "' AND IndConfirma='' AND IndVigente = 'X' /* AND  FG.IndFinMes='X'*/ ORDER BY FechaDefinicion"
    
    

    strEstado = Reg_Defecto
    With adoPendientes1
        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open strSQL
    End With
    
    tdgPendientes.DataSource = adoPendientes1
    tdgPendientes.Refresh
End Sub

Private Sub Deshabilita()

    strIndDetraccion = Valor_Caracter
    
    txtMontoDetraccion.Text = "0"
    Call Calculos
    
End Sub

Private Sub Habilita()

    strIndDetraccion = Valor_Indicador
'    cboFormaPagoDetraccion.Enabled = True
'    cboCuentaFondoDetraccion.Enabled = True
    
'    Call ColorControlHabilitado(cboFormaPagoDetraccion)
'    Call ColorControlHabilitado(cboCuentaFondoDetraccion)
    
    'Call cboFormaPagoDetraccion_Click
    If strCodMoneda <> Codigo_Moneda_Local Then
        txtMontoDetraccion.Text = CStr(CCur(txtTotal.Text) * gdblTasaDetraccion * CDbl(txtTipoCambioPago.Text))
    Else
        txtMontoDetraccion.Text = CStr(CCur(txtTotal.Text) * gdblTasaDetraccion)
    End If
    Call Calculos
    
End Sub


Private Sub cboAfectacion_Click()

    Dim intRegistro As Integer

    strCodAfectacion = Valor_Caracter
    If cboAfectacion.ListIndex < 0 Then Exit Sub
    
    strCodAfectacion = arrAfectacion(cboAfectacion.ListIndex)
    
    If strCodAfectacion = Codigo_Inafecto Then
        txtMontoNoGravado.Text = "0"
        intRegistro = ObtenerItemLista(arrDetraccion(), Codigo_Respuesta_No)
        If intRegistro >= 0 Then cboDetraccion.ListIndex = intRegistro
        cboDetraccion.Enabled = False
    Else
        cboDetraccion.Enabled = True
    End If
    
    Call Calculos
    
End Sub


Private Sub cboCreditoFiscal_Click()

    strCodCreditoFiscal = Valor_Caracter
    If cboCreditoFiscal.ListIndex < 0 Then Exit Sub
    
    strCodCreditoFiscal = arrCreditoFiscal(cboCreditoFiscal.ListIndex)
    
    'Call Calculos
    
End Sub





Private Sub cboDetraccion_Click()

    strDetraccionSiNo = Valor_Caracter
    If cboDetraccion.ListIndex < 0 Then Exit Sub
         
    strDetraccionSiNo = Trim(arrDetraccion(cboDetraccion.ListIndex))
    
    If cboDetraccion.Tag = "0" And strEstado = Reg_Edicion Then
        cboDetraccion.Tag = "1"
        Exit Sub
    End If
    
    Call Calculos
    
    
End Sub


Private Sub cboFondo_Click()

    Dim adoRegistro As ADODB.Recordset
    
    strCodFondo = Valor_Caracter
    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(arrFondo(cboFondo.ListIndex))
    
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Fecha Vigente, Valor Cuota, Valor Cuota T-1, Moneda y Cantidad Inicial de Cuotas del Fondo ***
        .CommandText = "{ call up_ACSelDatosParametro(23,'" & strCodFondo & "','" & gstrCodAdministradora & "') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            gdatFechaActual = adoRegistro("FechaCuota")
            gdblTipoCambio = adoRegistro("ValorTipoCambio")
            gstrCodMoneda = adoRegistro("CodMoneda")
            dtpFechaDesde.Value = gdatFechaActual
            dtpFechaHasta.Value = dtpFechaDesde.Value
            
            gdblTipoCambio = ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, gdatFechaActual, Codigo_Moneda_Local, gstrCodMoneda)
            If gdblTipoCambio = 0 Then gdblTipoCambio = ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, DateAdd("d", -1, gdatFechaActual), Codigo_Moneda_Local, gstrCodMoneda)
                        
            frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
            
            '*** Gastos del Fondo ***
'            strSQL = "SELECT (FCG.CodGasto + FCG.CodDetalleGasto + DCG.CodAnalitica) CODIGO,RTRIM(CG.DescripConcepto) + '-' + RTRIM(DCG.DescripGasto) DESCRIP " & _
'                "FROM FondoConceptoGasto FCG JOIN DetalleConceptoGasto DCG ON(DCG.CodDetalleGasto=FCG.CodDetalleGasto AND DCG.CodGasto=FCG.CodGasto) " & _
'                "JOIN ConceptoGasto CG ON(CG.CodGasto=FCG.CodGasto) " & _
'                "WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' " & _
'                "ORDER BY DescripGasto"
            strSQL = "SELECT (FCG.CodCuenta + CodAnalitica) CODIGO,(RTRIM(DescripCuenta)) DESCRIP " & _
                "FROM FondoConceptoGasto FCG JOIN PlanContable PCG ON(PCG.CodCuenta=FCG.CodCuenta AND PCG.CodAdministradora=FCG.CodAdministradora) " & _
                "JOIN FondoGasto FG ON(FG.CodCuenta=FCG.CodCuenta AND FG.CodAdministradora=FCG.CodAdministradora AND FG.CodFondo=FCG.CodFondo) " & _
                "WHERE (FG.CodFile='099' OR FG.CodFile<>'098') AND FCG.CodFondo='" & strCodFondo & "' AND FCG.CodAdministradora='" & gstrCodAdministradora & "' " & _
                 " UNION " & _
                    "SELECT (FCG.CodCuenta + CodAnalitica) CODIGO,(RTRIM(DescripCuenta)) DESCRIP " & _
                "FROM FondoConceptoActivoFijo FCG JOIN PlanContable PCG ON(PCG.CodCuenta=FCG.CodCuenta AND PCG.CodAdministradora=FCG.CodAdministradora) " & _
                "JOIN FondoGasto FG ON(FG.CodCuenta=FCG.CodCuenta AND FG.CodAdministradora=FCG.CodAdministradora AND FG.CodFondo=FCG.CodFondo) " & _
                "WHERE FG.CodFile='030' AND FCG.CodFondo='" & strCodFondo & "' AND FCG.CodAdministradora='" & gstrCodAdministradora & "' "
            
            CargarControlLista strSQL, cboGasto, arrGasto(), Sel_Defecto
            
            '*** Cuentas Bancarias ***
'            strSQL = "SELECT (CodFile + CodAnalitica) CODIGO,(RTRIM(DescripCuenta) + SPACE(1) + NumCuenta) DESCRIP FROM BancoCuenta " & _
'                "WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
'            CargarControlLista strSQL, cboCuentaFondoUnico, arrCuentaFondoUnico(), Sel_Defecto
'            CargarControlLista strSQL, cboCuentaFondoDetraccion, arrCuentaFondoDetraccion(), Sel_Defecto
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    Call CargarPendientes
    
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

Public Sub Salir()

    Unload Me
    
End Sub
Public Sub Cancelar()

    cmdOpcion.Visible = True
    With tabRegistroCompras
        .TabEnabled(0) = True
        .Tab = 0
    End With
    
    lblMensaje.Caption = Valor_Caracter
    cmdAccion.Button(0).Enabled = True
    Call Buscar
    
End Sub
Public Sub Grabar()
    
    
    Dim adoRegistro         As ADODB.Recordset
    Dim adoAuxiliar         As ADODB.Recordset
    Dim strNumCaja          As String
    Dim strCodDetalleFile   As String, strCodMonedaGasto                  As String
    Dim strDescripGasto     As String, strSQLOrdenCajaDetalleI       As String
    Dim strSQLOrdenCaja     As String, strSQLOrdenCajaDetalle   As String
    Dim strSQLOrdenCajaMN   As String, strSQLOrdenCajaDetalleMN As String
    Dim strFechaAnterior    As String, strFechaSiguiente        As String
    Dim curSaldoProvision   As Currency, intCantRegistros       As Integer
    Dim dblTipCambio        As Double, dblTipoCambioGasto   As Double
    Dim datFechaFinPeriodo  As Date, intNumGasto              As Integer
    Dim strNumComprobante   As String
    Dim numRegistro         As Long
    Dim strTipoMovimientoBanco  As String, strCodCuentaBanco    As String
    Dim strCodFileBanco         As String, strCodAnaliticaBanco      As String

    Dim strCodAuxiliar      As String
    Dim mensaje As String

    If strEstado = Reg_Consulta Then Exit Sub

    
    If strEstado = Reg_Adicion Then
        mensaje = Mensaje_Adicion
    Else
        mensaje = Mensaje_Edicion
    End If
    
    If Not TodoOK() Then Exit Sub
        
    If MsgBox(mensaje, vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption) <> vbYes Then Exit Sub
        
    Me.MousePointer = vbHourglass
    

    If strEstado = Reg_Adicion Then
    intNumGasto = CInt(tdgPendientes.Columns(2).Value)
    strNumComprobante = Trim(txtSerieComprobante.Text) & "-" & Trim(txtNumComprobante.Text)
    
         If strCodMoneda <> Codigo_Moneda_Local Then
            dblTipoCambioGasto = ObtenerTipoCambioMoneda2(Codigo_TipoCambio_SBS, Codigo_Valor_TipoCambioVenta, dtpFechaComprobante.Value, strCodMoneda, Codigo_Moneda_Local)
         Else
           
            dblTipoCambioGasto = 1
        End If
    Else
         intNumGasto = CInt(tdgConsulta.Columns(7).Value)
         strNumComprobante = Trim(txtSerieComprobante.Text) & "-" & Trim(txtNumComprobante.Text)
    End If
     
'    strCodAuxiliar = Codigo_Tipo_Persona_Proveedor & Trim(lblCodProveedor.Caption)
    
    '*** Guardar ***
    With adoComm
        
          .CommandText = "{ call up_CNManRegistroCompra('" & _
                strCodFondo & "','" & gstrCodAdministradora & "','" & Convertyyyymmdd(dtpFechaRegistro.Value) & "','" & _
                lblNumSecuencial.Caption & "' ,'" & strCodTipoComprobante & "','" & Convertyyyymmdd(dtpFechaComprobante.Value) & "','" & _
                strNumComprobante & "','" & strCodGasto & "','" & Codigo_Tipo_Persona_Proveedor & "','" & Trim(lblCodProveedor.Caption) & "','" & _
                Trim(txtDescripcion.Text) & "','" & strCodAfectacion & "','" & strCodCreditoFiscal & "','" & Trim(txtPeriodoFiscal.Text) & "','" & _
                strCodMoneda & "'," & CDec(txtSubTotal.Text) & "," & CDec(txtIgv.Text) & "," & CDec(txtMontoNoGravado.Value) & "," & CDec(txtTotal.Text) & ",'" & strDetraccionSiNo & "','" & _
                strCodFile & "','" & strCodAnalitica & "','" & Convertyyyymmdd(dtpFechaPago.Value) & "','" & strCodFormaPagoUnico & "','" & _
                strCodMonedaUnico & "','" & strTipoMovimientoBanco & "', '" & strCodCuentaBanco & "', '" & strCodFileBanco & "', '" & strCodAnaliticaBanco & "', '" & _
                strCodFileUnico & "','" & strCodAnaliticaUnico & "'," & _
                CDec(txtMontoUnico.Text) & ",'" & strCodFormaPagoDetraccion & "','" & strCodMonedaDetraccion & "','" & _
                strCodFileDetraccion & "','" & strCodAnaliticaDetraccion & "'," & CDec(txtMontoDetraccion.Text) & ",'" & _
                strCodValorTipoCambio & "'," & CDec(txtTipoCambioPago.Text) & "," & CDec(lblMontoTotal.Caption) & ",'" & Convertyyyymmdd(Valor_Fecha) & "','" & _
                intNumGasto & "','" & Estado_Activo & "','" & IIf(strEstado = Reg_Adicion, "I", "U") & "') }"
        adoComm.Execute .CommandText
        
        cmdContabilizar.Enabled = True
        
        
        Set adoRegistro = New ADODB.Recordset
            
            
           With adoComm
                    .CommandText = "SELECT * FROM FondoGasto " & _
                        "WHERE NumGasto= '" & Trim(tdgPendientes.Columns(2).Value) & "' AND CodFondo='" & _
                        strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
                    Set adoRegistro = .Execute
              
            
                
            '*** Generar Movimiento Contable de Impuesto ***
                
                If strEstado = Reg_Adicion And CStr(adoRegistro("CodFile")) = "099" Then
                
                    If strCodAplicacionDevengo = Codigo_Aplica_Devengo_Inmediata Then
                        Call GenerarAsientoGasto(strCodFile, strCodAnalitica, strCodFondo, gstrCodAdministradora, strCodDetalleGasto, Codigo_Dinamica_Gasto, CCur(txtSubTotal.Text), CCur(txtMontoNoGravado.Value), dblTipoCambioGasto, strCodMoneda, Trim(txtDescripcion.Text), frmMainMdi.Tag, strCodTipoComprobante, strNumComprobante, strCodAfectacion, Codigo_Tipo_Persona_Proveedor, Trim(lblCodProveedor.Caption))
                        Call GenerarAsientoGasto(strCodFile, strCodAnalitica, strCodFondo, gstrCodAdministradora, strCodDetalleGasto, Codigo_Dinamica_Gasto_Emitida, CCur(txtSubTotal.Text), CCur(txtMontoNoGravado.Value), dblTipoCambioGasto, strCodMoneda, Trim(txtDescripcion.Text), frmMainMdi.Tag, strCodTipoComprobante, strNumComprobante, strCodAfectacion, Codigo_Tipo_Persona_Proveedor, Trim(lblCodProveedor.Caption))
                    End If
                    
                    If strCodAplicacionDevengo = Codigo_Aplica_Devengo_Periodica Then
                        Call GenerarAsientoGasto(strCodFile, strCodAnalitica, strCodFondo, gstrCodAdministradora, strCodDetalleGasto, Codigo_Dinamica_Gasto_Emitida, CCur(txtSubTotal.Text), CCur(txtMontoNoGravado.Value), dblTipoCambioGasto, strCodMoneda, Trim(txtDescripcion.Text), frmMainMdi.Tag, strCodTipoComprobante, strNumComprobante, strCodAfectacion, Codigo_Tipo_Persona_Proveedor, Trim(lblCodProveedor.Caption))
                    End If
                    
                    '*** Generar Orden si no está generada o actualizar ***
                    Call GenerarOrdenGastosFondo(strCodDetalleGasto, strCodGasto, strCodFondo, CInt(tdgPendientes.Columns("NumGasto").Value), Trim(lblCodProveedor.Caption), numRegistro)
                
                End If
            End With
            adoRegistro.Close: Set adoRegistro = Nothing
'--------------------------------------
        
    End With


        Me.MousePointer = vbDefault
                    
        MsgBox Mensaje_Edicion_Exitosa, vbExclamation
        
        frmMainMdi.stbMdi.Panels(3).Text = "Acción"
        
        cmdOpcion.Visible = True
        With tabRegistroCompras
            .TabEnabled(0) = True
            .Tab = 0
        End With
        
        txtSerieComprobante.Text = Valor_Caracter
        txtNumComprobante.Text = Valor_Caracter
        
        Call CargarPendientes
        Call Buscar
  
    
    
End Sub

Private Function TodoOK() As Boolean
        
    TodoOK = False
            
    If cboTipoComprobante.ListIndex <= 0 Then
        MsgBox "Seleccione el tipo de comprobante", vbCritical, Me.Caption
        If cboTipoComprobante.Enabled Then cboTipoComprobante.SetFocus
        Exit Function
    End If
        
    If Trim(txtSerieComprobante.Text) = Valor_Caracter Then
        MsgBox "Ingrese el número de serie, si no lo tiene ingrese cero", vbCritical, Me.Caption
        If txtSerieComprobante.Enabled Then txtSerieComprobante.SetFocus
        Exit Function
    End If
    
    If Trim(txtNumComprobante.Text) = Valor_Caracter Then
        MsgBox "Ingrese el número de comprobante", vbCritical, Me.Caption
        If txtNumComprobante.Enabled Then txtNumComprobante.SetFocus
        Exit Function
    End If
    
    If cboGasto.ListIndex <= 0 Then
        MsgBox "Seleccione el gasto relacionado", vbCritical, Me.Caption
        If cboGasto.Enabled Then cboGasto.SetFocus
        Exit Function
    End If
    
    If Trim(lblProveedor.Caption) = Valor_Caracter Then
        MsgBox "Seleccione el Proveedor", vbCritical, Me.Caption
        If cmdProveedor.Enabled Then cmdProveedor.SetFocus
        Exit Function
    End If
    
    If strDetraccionSiNo = Codigo_Respuesta_Si And strCodMoneda <> Codigo_Moneda_Local Then
        If CDbl(txtTipoCambioPago.Text) = 0 Then
            MsgBox "Tipo de Cambio SUNAT NO REGISTRADO...", vbCritical, Me.Caption
            If cboTipoValorCambio.Enabled Then cboTipoValorCambio.SetFocus
            Exit Function
        End If
    End If
    
    If Codigo_Moneda_Local <> strCodMonedaUnico Then
        'If strCodTipoComprobante = Codigo_Comprobante_Factura Then 'Factura
            If ObtenerTipoCambioMoneda(Codigo_TipoCambio_SBS, Codigo_Valor_TipoCambioVenta, dtpFechaComprobante.Value, Codigo_Moneda_Local, strCodMonedaUnico) = 0 Then
                MsgBox "Tipo de Cambio SUNAT NO REGISTRADO...", vbCritical, Me.Caption
                If cboTipoValorCambio.Enabled Then cboTipoValorCambio.SetFocus
                Exit Function
            End If
        'Else
        '    If ObtenerTipoCambio(Codigo_TipoCambio_Sunat, Codigo_Valor_TipoCambioVenta, gdatFechaActual, strCodMonedaUnico) = 0 Then
        '        If ObtenerTipoCambio(Codigo_TipoCambio_Sunat, Codigo_Valor_TipoCambioVenta, gdatFechaActual, strCodMonedaUnico) = 0 Then
        '            MsgBox "Tipo de Cambio SUNAT NO REGISTRADO...", vbCritical, Me.Caption
        '            If cboTipoValorCambio.Enabled Then cboTipoValorCambio.SetFocus
        '            Exit Function
        '        End If
        '    End If
        'End If
    End If
          
    '*** Si todo paso OK ***
    TodoOK = True
  
End Function
Public Sub Imprimir()
    
    Call SubImprimir(1)
    
End Sub

Public Sub SubImprimir(Index As Integer)

    Dim strSeleccionRegistro    As String
    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()

    Select Case Index
        
        Case 1
        
            gstrNameRepo = "ComprobantePagoRegistroGrilla"
            Set frmReporte = New frmVisorReporte
            
            ReDim aReportParamS(3)
            ReDim aReportParamFn(5)
            ReDim aReportParamF(5)
            
            aReportParamFn(0) = "Usuario"
            aReportParamFn(1) = "Hora"
            aReportParamFn(2) = "NombreEmpresa"
            aReportParamFn(3) = "Fondo"
            aReportParamFn(4) = "FechaDel"
            aReportParamFn(5) = "FechaAl"
                
            aReportParamF(0) = gstrLogin
            aReportParamF(1) = Format(Time(), "hh:mm:ss")
            aReportParamF(2) = gstrNombreEmpresa & Space(1)
            aReportParamF(3) = Trim(cboFondo.Text)
            aReportParamF(4) = CStr(dtpFechaDesde.Value)
            aReportParamF(5) = CStr(dtpFechaHasta.Value)
                            
            aReportParamS(0) = strCodFondo
            aReportParamS(1) = gstrCodAdministradora
            aReportParamS(2) = Convertyyyymmdd(dtpFechaDesde.Value)
            aReportParamS(3) = Convertyyyymmdd(DateAdd("d", 1, dtpFechaHasta.Value))
        
        Case 2, 3, 4
'            If Index = 1 Then gstrNameRepo = "RegistroCompraDAOT"
            'If index = 2 Then gstrNameRepo = "RegistroComprasParte2"
            If Index = 2 Then gstrNameRepo = "RegistroVenta"
            If Index = 3 Then gstrNameRepo = "RegistroCompra"
            If Index = 4 Then gstrNameRepo = "LibroRetenciones"
            Set frmReporte = New frmVisorReporte
            
            ReDim aReportParamS(6)
            ReDim aReportParamFn(7)
            ReDim aReportParamF(7)
            
            aReportParamFn(0) = "Usuario"
            aReportParamFn(1) = "Hora"
            aReportParamFn(2) = "NombreEmpresa"
            aReportParamFn(3) = "Fondo"
            aReportParamFn(4) = "FechaDesde"
            aReportParamFn(5) = "FechaHasta"
            aReportParamFn(6) = "TipoCambio"
            aReportParamFn(7) = "Moneda"
                
            aReportParamF(0) = gstrLogin
            aReportParamF(1) = Format(Time(), "hh:mm:ss")
            aReportParamF(2) = gstrNombreEmpresa & Space(1)
            aReportParamF(3) = Trim(cboFondo.Text)
            aReportParamF(4) = CStr(dtpFechaDesde.Value)
            aReportParamF(5) = CStr(dtpFechaHasta.Value)
            aReportParamF(6) = gdblTipoCambio
            aReportParamF(7) = Valor_Caracter
                            
            aReportParamS(0) = strCodFondo
            aReportParamS(1) = gstrCodAdministradora
            aReportParamS(2) = Convertyyyymmdd(dtpFechaDesde.Value)
            aReportParamS(3) = Convertyyyymmdd(DateAdd("d", 1, dtpFechaHasta.Value))
        
            'MsgBox "El reporte muestra el registro de compras en soles y dolares", vbInformation, Clave_Registro_NombreSistema
            gstrCodMoneda = "0"
            aReportParamS(4) = gstrCodMoneda
            aReportParamS(5) = "04"
            aReportParamS(6) = "COMPRA"
    End Select
        
    gstrSelFrml = Valor_Caracter
    
    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"

    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())

    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
    frmReporte.Show vbModal

    Set frmReporte = Nothing

    Screen.MousePointer = vbNormal
    
End Sub
Public Sub Eliminar()
    
  Dim adoContabilizado As ADODB.Recordset
    
    
     If strEstado = Reg_Consulta And adoConsulta1.RecordCount > 0 Then
        
        With adoComm
        
            .CommandText = " SELECT IndConfirma FROM FondoGasto WHERE NumGasto= " & tdgConsulta.Columns("NumGasto").Value & "  AND CodFondo=" & strCodFondo & " AND CodAdministradora=" & gstrCodAdministradora & " "
            
            Set adoContabilizado = .Execute
                
                If Trim(adoContabilizado("IndConfirma")) = Valor_Caracter Then
                    
                    
                    With adoComm
        
                        If MsgBox("Desea Anular el registro de compras Nro. " & tdgConsulta.Columns("NumRegistro").Value & " asociado al Número de Gastos " & "'" & tdgConsulta.Columns("NumGasto").Value & "'" & " ?", vbYesNo + vbQuestion + vbDefaultButton2) <> vbYes Then Exit Sub
                        
                        .CommandText = "UPDATE RegistroCompra SET Estado = '02'" & _
                                        " WHERE " & _
                                        "NumRegistro=" & tdgConsulta.Columns("NumRegistro").Value & " AND CodFondo='" & _
                                        strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
                        .Execute
                    
                    
                        .CommandText = "UPDATE FondoGasto SET IndFinMes = 'X' " & _
                                       " WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND NumGasto = " & tdgConsulta.Columns("NumGasto").Value & " AND IndVigente = 'X' "
                    
                        .Execute
                    
                        Call Buscar
                    
                        MsgBox "Operación Realizada Exitosamente", vbInformation + vbOKOnly, Me.Caption
           
             
                    End With
                    
                Else
                    
                    MsgBox "NO puede Eliminar un Registro Ya Contabilizado!", vbCritical + vbOKOnly, Me.Caption
                    
                End If
         End With
    Else
    
        MsgBox "NO existen Registros por Eliminar!", vbInformation + vbOKOnly, Me.Caption
    
    End If
    
End Sub
Public Sub Modificar()

    If strEstado = Reg_Consulta Then
        strEstado = Reg_Edicion
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabRegistroCompras
            .TabEnabled(0) = False
            .Tab = 1
        End With
        
    End If
    
End Sub

Private Sub LlenarFormulario(strModo As String)

    Dim adoRegistro     As ADODB.Recordset, intRegistro       As Integer
    Dim adoAuxiliar     As ADODB.Recordset
    Dim intPosiCaracter As Integer, intCantCaracteres As Integer
    Dim strFechaInicio      As String, strFechaFin      As String
            
            
            
    Select Case strModo
    
        Case Reg_Adicion
            
            Set adoRegistro = New ADODB.Recordset
            
            
                With adoComm
                    .CommandText = "SELECT * FROM FondoGasto " & _
                        "WHERE NumGasto= '" & Trim(tdgPendientes.Columns(2).Value) & "' AND CodFondo='" & _
                        strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
                    Set adoRegistro = .Execute
                    
                    If Not adoRegistro.EOF Then
                    
                    
                        If CStr(adoRegistro("CodFile")) = "030" Then
                             cmdContabilizar.Visible = True
                             cmdContabilizar.Enabled = True
                        Else
                            cmdContabilizar.Visible = False
                            cmdContabilizar.Enabled = False
                        End If
                    End If
                 End With
            adoRegistro.Close: Set adoRegistro = Nothing
                
            fraCompras(1).Caption = "Definición del Registro - Fondo : " & Trim(cboFondo.Text)
            fraCompras(2).Caption = "Definición de Obligación - Fondo : " & Trim(cboFondo.Text)
            
            Set adoRegistro = New ADODB.Recordset
            
        With adoComm
            
                
                lblNumSecuencial.Caption = "GENERADO"
                
                dtpFechaRegistro.Value = gdatFechaActual
                If cboTipoComprobante.ListCount > 0 Then cboTipoComprobante.ListIndex = 0
                dtpFechaComprobante.Value = gdatFechaActual
                txtNumComprobante.Text = Valor_Caracter
                txtSerieComprobante.Text = Valor_Caracter
                If cboGasto.ListCount > 0 Then cboGasto.ListIndex = 0
                lblProveedor.Caption = Valor_Caracter
                lblDireccion.Caption = Valor_Caracter
                txtDescripcion.Text = Valor_Caracter
                lblAnalitica.Caption = Valor_Caracter
                
                cmdContabilizar.Enabled = False
                
                If cboAfectacion.ListCount > 0 Then cboAfectacion.ListIndex = 0
                
                intRegistro = ObtenerItemLista(arrAfectacion(), Codigo_Afecto)
                If intRegistro >= 0 Then cboAfectacion.ListIndex = intRegistro
                
                If cboCreditoFiscal.ListCount > 0 Then cboCreditoFiscal.ListIndex = 0
                                        
                txtPeriodoFiscal.Text = Valor_Caracter
                txtSubTotal.Text = "0": txtIgv.Text = "0"
                txtTotal.Text = "0"
                
                'dtpFechaPago.MinDate = gdatFechaActual 'acr
                dtpFechaPago.Value = gdatFechaActual
                dtpFechaTipoCambioPago.Value = gdatFechaActual
                                
'                If cboFormaPagoUnico.ListCount > 0 Then cboFormaPagoUnico.ListIndex = 0
'                intRegistro = ObtenerItemLista(arrFormaPagoUnico(), Codigo_FormaPago_Efectivo)
'                If intRegistro >= 0 Then cboFormaPagoUnico.ListIndex = intRegistro
                
'                If cboCuentaFondoUnico.ListCount > 0 Then cboCuentaFondoUnico.ListIndex = 0
                
                If cboMonedaUnico.ListCount > 0 Then cboMonedaUnico.ListIndex = 0
                intRegistro = ObtenerItemLista(arrMonedaUnico(), gstrCodMoneda)
                If intRegistro >= 0 Then cboMonedaUnico.ListIndex = intRegistro
                
                txtMontoUnico.Text = "0"
                
'                If cboFormaPagoDetraccion.ListCount > 0 Then cboFormaPagoDetraccion.ListIndex = 0
'                intRegistro = ObtenerItemLista(arrFormaPagoDetraccion(), Codigo_FormaPago_Efectivo)
'                If intRegistro >= 0 Then cboFormaPagoDetraccion.ListIndex = intRegistro
                
'                If cboCuentaFondoDetraccion.ListCount > 0 Then cboCuentaFondoDetraccion.ListIndex = 0
                
                If cboMonedaDetraccion.ListCount > 0 Then cboMonedaDetraccion.ListIndex = 0
                intRegistro = ObtenerItemLista(arrMonedaDetraccion(), Codigo_Moneda_Local)
                If intRegistro >= 0 Then cboMonedaDetraccion.ListIndex = intRegistro
                cboMonedaDetraccion.Enabled = False
                
                If cboDetraccion.ListCount > 0 Then cboDetraccion.ListIndex = 0
                intRegistro = ObtenerItemLista(arrDetraccion(), Codigo_Respuesta_No)
                If intRegistro >= 0 Then cboDetraccion.ListIndex = intRegistro
                
                txtMontoDetraccion.Text = "0"
                txtTipoCambioPago.Text = gdblTipoCambio
                
                lblMontoTotal.Caption = "0"
                strCodTipoGasto = tdgPendientes.Columns(6).Value
                strCodFile = Trim(tdgPendientes.Columns(9).Value)
                
                Me.Refresh
                               
                .CommandText = "SELECT * FROM FondoGasto " & _
                    "WHERE NumGasto=" & CInt(tdgPendientes.Columns(2)) & " AND CodCuenta='" & Trim(tdgPendientes.Columns(1)) & "' AND " & _
                    "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
                Set adoRegistro = adoComm.Execute
                
                If Not adoRegistro.EOF Then
                    lblAnalitica.Caption = Trim(adoRegistro("CodFile")) & " - " & Trim(adoRegistro("CodAnalitica"))
                
                    txtDescripcion.Text = Trim(adoRegistro("DescripGasto"))
                    lblMontoGasto.Caption = CStr(adoRegistro("MontoGasto"))
                    txtSubTotal.Text = CStr(adoRegistro("MontoGasto"))
                    
                    cboGasto.Enabled = True
                    intRegistro = ObtenerItemLista(arrGasto(), adoRegistro("CodCuenta") + adoRegistro("CodAnalitica"))
                    If intRegistro >= 0 Then cboGasto.ListIndex = intRegistro
                    cboGasto.Enabled = False
                                                            
                                                            
                    intRegistro = ObtenerItemLista(arrMoneda(), adoRegistro("CodMoneda"))
                    If intRegistro >= 0 Then cboMoneda.ListIndex = intRegistro
                    
                    
                    intRegistro = ObtenerItemLista(arrMonedaUnico(), adoRegistro("CodMoneda"))
                    If intRegistro >= 0 Then cboMonedaUnico.ListIndex = intRegistro
                    
                    intRegistro = ObtenerItemLista(arrCreditoFiscal(), adoRegistro("CodCreditoFiscal"))
                    If intRegistro >= 0 Then cboCreditoFiscal.ListIndex = intRegistro

                    strCodAplicacionDevengo = adoRegistro("CodAplicacionDevengo")
                    
                    dtpFechaPago.Value = gdatFechaActual
                    
'                    If strCodAplicacionDevengo = Codigo_Aplica_Devengo_Inmediata Then
'                        If CDate(adoRegistro("FechaFinal")) >= dtpFechaPago.Value Then
'                            dtpFechaPago.Value = adoRegistro("FechaFinal")
'                        End If
'                    Else
'                        If DateAdd("d", 1, adoRegistro("FechaFinal")) >= dtpFechaPago.Value Then
'                            dtpFechaPago.Value = DateAdd("d", 1, adoRegistro("FechaFinal"))
'                        End If
'                    End If

                    Set adoAuxiliar = New ADODB.Recordset
                    
                    If Trim(tdgPendientes.Columns(9).Value) = "099" Or Trim(tdgPendientes.Columns(9).Value) <> "098" Then
                        .CommandText = "SELECT CodDetalleFile FROM InversionDetalleFile " & _
                            "WHERE CodFile='" & Trim(tdgPendientes.Columns(9).Value) & "' AND DescripDetalleFile='" & strCodGasto & "'"
                    Else
                        .CommandText = "SELECT CodDetalleFile FROM DinamicaContable " & _
                            "WHERE CodFile='" & Trim(tdgPendientes.Columns(9).Value) & "' AND CodCuenta='" & strCodGasto & "'"
                    End If
                    Set adoAuxiliar = .Execute
                    
                    If Not adoAuxiliar.EOF Then
                        strCodDetalleGasto = adoAuxiliar("CodDetalleFile")
                    End If
                    adoAuxiliar.Close: Set adoAuxiliar = Nothing

                    Set adoAuxiliar = New ADODB.Recordset
                    
                    .CommandText = "SELECT IP.CodPersona, IP.NumIdentidad, IP.DescripPersona, IP.Direccion1 + IP.Direccion2 Direccion, AP.DescripParametro TipoIdentidad " & _
                        "FROM InstitucionPersona IP " & _
                        "JOIN AuxiliarParametro AP ON (AP.CodParametro = IP.TipoIdentidad AND AP.CodTipoParametro = 'TIPIDE')" & _
                        "WHERE CodPersona='" & adoRegistro("CodProveedor") & "' AND TipoPersona='" & Codigo_Tipo_Persona_Proveedor & "'"
                    Set adoAuxiliar = .Execute
                    
                    If Not adoAuxiliar.EOF Then
                        lblTipoDocID.Caption = Trim(adoAuxiliar("TipoIdentidad"))
                        lblProveedor.Caption = Trim(adoAuxiliar("DescripPersona"))
                        lblNumDocID.Caption = Trim(adoAuxiliar("NumIdentidad"))
                        lblDireccion.Caption = Trim(adoAuxiliar("Direccion"))
                        lblCodProveedor.Caption = Trim(adoAuxiliar("CodPersona"))
                    End If
                    adoAuxiliar.Close: Set adoAuxiliar = Nothing
                    
                End If
                adoRegistro.Close: Set adoRegistro = Nothing
                
                If cboTipoValorCambio.ListCount > 0 Then cboTipoValorCambio.ListIndex = 0
                intRegistro = ObtenerItemLista(arrTipoValorCambio(), Codigo_Valor_TipoCambioVenta)
                If intRegistro >= 0 Then cboTipoValorCambio.ListIndex = intRegistro
            End With
                        
            cboTipoComprobante.SetFocus
        
'-----------------------------------------------------------
        
        Case Reg_Edicion
            
            
        strFechaInicio = Convertyyyymmdd(tdgConsulta.Columns(6).Value)
        strFechaFin = Convertyyyymmdd(DateAdd("d", 1, tdgConsulta.Columns(6).Value))
            
            Set adoRegistro = New ADODB.Recordset
            
            
            With adoComm
                .CommandText = "SELECT * FROM RegistroCompra " & _
                    "WHERE (FechaRegistro>='" & strFechaInicio & "' AND FechaRegistro<'" & strFechaFin & "') AND " & _
                    "NumRegistro= '" & Trim(tdgConsulta.Columns(1).Value) & "' AND CodFondo='" & _
                    strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
                Set adoRegistro = .Execute
                
                If Not adoRegistro.EOF Then
                
                    If CStr(adoRegistro("CodFileGasto")) = "030" Then
                         cmdContabilizar.Visible = True
                         cmdContabilizar.Enabled = True
                    Else
                        cmdContabilizar.Visible = False
                        cmdContabilizar.Enabled = False
                    End If
                    
                    fraCompras(1).Caption = "Definición del Registro - Fondo : " & Trim(cboFondo.Text)
                    fraCompras(2).Caption = "Definición de Obligación - Fondo : " & Trim(cboFondo.Text)
                    
                    lblNumSecuencial.Caption = CStr(adoRegistro("NumRegistro"))
                    dtpFechaRegistro.Value = adoRegistro("FechaRegistro")
                    
                    intRegistro = ObtenerItemLista(arrTipoComprobante(), adoRegistro("CodTipoComprobante"))
                    If intRegistro >= 0 Then cboTipoComprobante.ListIndex = intRegistro
                    
                    dtpFechaComprobante.Value = adoRegistro("FechaComprobante")
                    
                    intPosiCaracter = InStr(Trim(adoRegistro("NumComprobante")), "-")
                    txtSerieComprobante.Text = Left(Trim(adoRegistro("NumComprobante")), intPosiCaracter - 1)
                    intCantCaracteres = Len(Trim(adoRegistro("NumComprobante")))
                    txtNumComprobante.Text = Mid(Trim(adoRegistro("NumComprobante")), intPosiCaracter + 1, intCantCaracteres - intPosiCaracter)
                    
                    
                    strIndDetraccion = Valor_Caracter
                    If CCur(adoRegistro("MontoDetraccion")) > 0 Then strIndDetraccion = Valor_Indicador
                    
                    intRegistro = ObtenerItemLista(arrGasto(), adoRegistro("CodCuenta") + adoRegistro("CodAnaliticaGasto"))
                    If intRegistro >= 0 Then cboGasto.ListIndex = intRegistro
                                        
                    Set adoAuxiliar = New ADODB.Recordset
                    
                    If Trim(tdgConsulta.Columns(8).Value) = "099" Or Trim(tdgConsulta.Columns(8).Value) <> "098" Then
                        .CommandText = "SELECT CodDetalleFile FROM InversionDetalleFile " & _
                            "WHERE CodFile='" & Trim(tdgConsulta.Columns(8).Value) & "' AND DescripDetalleFile='" & strCodGasto & "'"
                    Else
                        .CommandText = "SELECT CodDetalleFile FROM DinamicaContable " & _
                            "WHERE CodFile='" & Trim(tdgConsulta.Columns(8).Value) & "' AND CodCuenta='" & strCodGasto & "'"
                    End If
                    Set adoAuxiliar = .Execute
                    
                    If Not adoAuxiliar.EOF Then
                        strCodDetalleGasto = adoAuxiliar("CodDetalleFile")
                    End If
                    adoAuxiliar.Close: Set adoAuxiliar = Nothing
                                        
                                        
                    Set adoAuxiliar = New ADODB.Recordset
                                                                                    'CodCuenta='" & Trim(tdgPendientes.Columns(1)) & "' AND
                    .CommandText = "SELECT * FROM FondoGasto " & _
                        "WHERE NumGasto=" & CInt(tdgConsulta.Columns(7)) & " AND " & _
                        "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
                    Set adoAuxiliar = adoComm.Execute
                                        
                    If Not adoAuxiliar.EOF Then
                        strCodAplicacionDevengo = adoAuxiliar("CodAplicacionDevengo")
                    End If
                    adoAuxiliar.Close: Set adoAuxiliar = Nothing
                                        
                    lblProveedor.Caption = Valor_Caracter
                    lblDireccion.Caption = Valor_Caracter
                    
                    Set adoAuxiliar = New ADODB.Recordset
                    
                    .CommandText = "SELECT IP.NumIdentidad, IP.DescripPersona, IP.Direccion1 + IP.Direccion2 Direccion, AP.DescripParametro TipoIdentidad " & _
                        "FROM InstitucionPersona IP " & _
                        "JOIN AuxiliarParametro AP ON (AP.CodParametro = IP.TipoIdentidad AND AP.CodTipoParametro = 'TIPIDE')" & _
                        "WHERE CodPersona='" & adoRegistro("CodProveedor") & "' AND TipoPersona='" & Codigo_Tipo_Persona_Proveedor & "'"
                    Set adoAuxiliar = .Execute
                    
                    If Not adoAuxiliar.EOF Then
                        lblTipoDocID.Caption = Trim(adoAuxiliar("TipoIdentidad"))
                        lblProveedor.Caption = Trim(adoAuxiliar("DescripPersona"))
                        lblNumDocID.Caption = Trim(adoAuxiliar("NumIdentidad"))
                        lblDireccion.Caption = Trim(adoAuxiliar("Direccion"))
                    End If
                    adoAuxiliar.Close: Set adoAuxiliar = Nothing
                                        
                    txtDescripcion.Text = Trim(adoRegistro("DescripRegistro"))
                    
                    intRegistro = ObtenerItemLista(arrAfectacion(), adoRegistro("CodAfectacion"))
                    If intRegistro >= 0 Then cboAfectacion.ListIndex = intRegistro
                                                            
                    intRegistro = ObtenerItemLista(arrCreditoFiscal(), adoRegistro("CodCreditoFiscal"))
                    If intRegistro >= 0 Then cboCreditoFiscal.ListIndex = intRegistro
                                                                                                    
                    intRegistro = ObtenerItemLista(arrMoneda(), adoRegistro("CodMoneda"))
                    If intRegistro >= 0 Then cboMoneda.ListIndex = intRegistro
                    
                    txtPeriodoFiscal.Text = Valor_Caracter
                    txtSubTotal.Text = CStr(adoRegistro("Importe"))
                    txtIgv.Text = CStr(adoRegistro("ValorImpuesto"))
                    txtTotal.Text = CStr(adoRegistro("ValorTotal"))
                                                            
                    dtpFechaPago.Value = adoRegistro("FechaPago")
                    
                    intRegistro = ObtenerItemLista(arrDetraccion(), adoRegistro("CodDetraccionSiNo"))
                    If intRegistro >= 0 Then cboDetraccion.ListIndex = intRegistro
                    
                    If strDetraccionSiNo = Codigo_Respuesta_Si Then
                        dtpFechaTipoCambioPago.Value = adoRegistro("FechaComprobante")
                    Else
                        dtpFechaTipoCambioPago.Value = adoRegistro("FechaPago")
                    End If
                    
'                    intRegistro = ObtenerItemLista(arrFormaPagoUnico(), adoRegistro("CodFormaPago"))
'                    If intRegistro >= 0 Then cboFormaPagoUnico.ListIndex = intRegistro
                    
'                    intRegistro = ObtenerItemLista(arrCuentaFondoUnico(), adoRegistro("CodFile") + adoRegistro("CodAnalitica"))
'                    If intRegistro >= 0 Then cboCuentaFondoUnico.ListIndex = intRegistro
                                                                                
                    intRegistro = ObtenerItemLista(arrMonedaUnico(), adoRegistro("CodMonedaPago"))
                    If intRegistro >= 0 Then cboMonedaUnico.ListIndex = intRegistro
                    
                    txtMontoUnico.Text = CStr(adoRegistro("MontoPago"))
                                        
'                    intRegistro = ObtenerItemLista(arrFormaPagoDetraccion(), adoRegistro("CodFormaPagoDetraccion"))
'                    If intRegistro >= 0 Then cboFormaPagoDetraccion.ListIndex = intRegistro
                    
'                    intRegistro = ObtenerItemLista(arrCuentaFondoDetraccion(), adoRegistro("CodFileDetraccion") + adoRegistro("CodAnaliticaDetraccion"))
'                    If intRegistro >= 0 Then cboCuentaFondoDetraccion.ListIndex = intRegistro
                                                                                
                    intRegistro = ObtenerItemLista(arrMonedaDetraccion(), adoRegistro("CodMonedaDetraccion"))
                    If intRegistro >= 0 Then cboMonedaDetraccion.ListIndex = intRegistro
                    cboMonedaDetraccion.Enabled = False
                    
                    txtMontoDetraccion.Text = CStr(adoRegistro("MontoDetraccion"))
                    
                    intRegistro = ObtenerItemLista(arrTipoValorCambio(), adoRegistro("ClaseTipoCambio"))
                    If intRegistro >= 0 Then cboTipoValorCambio.ListIndex = intRegistro
                    
                    txtTipoCambioPago.Text = CStr(adoRegistro("TipoCambioPago"))
                    
                    
                    lblMontoTotal.Caption = CStr(adoRegistro("MontoTotal"))
                    
                    If adoPendientes1.RecordCount > 0 Then
'                     strCodFile = Trim(tdgPendientes.Columns(9).Value)
'                    Else
'                     strCodFile = Valor_Caracter
                    End If
                    
                    cboTipoComprobante.SetFocus
                                         
                                         
                                         
                    Dim adoRegistroTmp As ADODB.Recordset
                    Set adoRegistroTmp = New ADODB.Recordset
            
                    strSQL = "SELECT RTRIM(IndConfirma) AS IndConfirma FROM RegistroCompra RC JOIN FondoGasto FG " & _
                            "ON(RC.CodFondo=FG.CodFondo AND RC.CodAdministradora=FG.CodAdministradora " & _
                            "AND RC.NumGasto=FG.NumGasto) WHERE RC.NumRegistro='" & tdgConsulta.Columns(1) & "' "
                    
                    With adoComm
                    
                        .CommandText = strSQL
                        
                        Set adoRegistroTmp = .Execute
                    
                    End With
            
                    If Not adoRegistroTmp.EOF Then
                    
                        Do Until adoRegistroTmp.EOF
                            
                            If adoRegistroTmp("IndConfirma").Value = "X" Then
                            
        '                        MsgBox "Este registro ya fue contabilizado", vbOKOnly + vbCritical, Me.Caption
        '                        adoRegistroTmp.Close: Set adoRegistroTmp = Nothing
        '                        Exit Sub
                                
                                cmdAccion.Button(0).Enabled = False
                                cmdContabilizar.Enabled = False
                                lblMensaje.Caption = " : Este registro ya fue contabilizado : "
                                
                            End If
                            adoRegistroTmp.MoveNext
                        Loop
                        adoRegistroTmp.Close: Set adoRegistroTmp = Nothing
                    End If
                                                                   
                                         
                End If
               
                adoRegistro.Close: Set adoRegistro = Nothing
            End With
    End Select
    
End Sub
Public Sub Adicionar()
        
    frmMainMdi.stbMdi.Panels(3).Text = "Adicionar Registro..."
                
    If adoPendientes1.RecordCount > 0 Then
        tdgPendientes.SetFocus
    Else
        MsgBox "No existen gastos pendientes", vbCritical, Me.Caption
        tdgConsulta.SetFocus
        Exit Sub
    End If
    
    strEstado = Reg_Adicion
    LlenarFormulario strEstado
    cmdOpcion.Visible = False
    With tabRegistroCompras
        .TabEnabled(0) = False
        .Tab = 1
    End With
    
End Sub







Private Sub cboGasto_Click()

    Dim adoRegistro         As ADODB.Recordset
    Dim curDiferencia       As Currency
    Dim intRegistro         As Integer
        
    strCodGasto = Valor_Caracter: strCodAnalitica = Valor_Caracter
    strCodDetalleGasto = Valor_Caracter
    If cboGasto.ListIndex < 0 Then Exit Sub
    
    strCodGasto = Trim(Left(arrGasto(cboGasto.ListIndex), 10))
    strCodAnalitica = Right(arrGasto(cboGasto.ListIndex), 8)
    
    lblAnalitica.Caption = Trim(tdgPendientes.Columns(9).Value) & " - " & strCodAnalitica
        
'    Set adoRegistro = New ADODB.Recordset
'
'    With adoComm

'        txtDescripcion.Text = Valor_Caracter
'        lblMontoGasto.Caption = "0"
'        txtSubTotal.ToolTipText = Valor_Caracter
'        txtSubTotal.Text = "0"
        
'        .CommandText = "SELECT MontoGasto,MontoDevengo,DescripGasto,FechaFinal,CodCreditoFiscal,CodMoneda,CodTipoGasto " & _
'            "FROM FondoGasto " & _
'            "WHERE NumGasto=" & CInt(tdgPendientes.Columns(2).Value) & " AND CodCuenta='" & strCodGasto & "' AND " & _
'            "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
'        Set adoRegistro = .Execute
'
'        If Not adoRegistro.EOF Then
        
'            txtDescripcion.Text = Trim(adoRegistro("DescripGasto"))
'            lblMontoGasto.Caption = CStr(adoRegistro("MontoGasto"))
            
'            curDiferencia = adoRegistro("MontoGasto") - adoRegistro("MontoDevengo")
'            If curDiferencia > 0 Then
'                txtSubTotal.ToolTipText = "Faltan provisionar " & CStr(curDiferencia)
'            Else
'                txtSubTotal.ToolTipText = Valor_Caracter
'            End If
            
'            intRegistro = ObtenerItemLista(arrCreditoFiscal(), adoRegistro("CodCreditoFiscal"))
'            If intRegistro >= 0 Then cboCreditoFiscal.ListIndex = intRegistro
            
'            If strCodCreditoFiscal = Codigo_Tipo_Credito_RentaNoGravada Then
'                If a Then
'                    txtTotal.Text = CStr(adoRegistro("MontoGasto"))
'                Else
'                    txtSubTotal.Text = CStr(adoRegistro("MontoGasto"))
'                End If
'            ElseIf strCodCreditoFiscal = Codigo_Tipo_Credito_AdquisicionesNoGravada Then
'                txtTotal.Text = CStr(adoRegistro("MontoGasto"))
'            Else
'                txtSubTotal.Text = CStr(adoRegistro("MontoGasto"))
'            End If
            
'            If adoRegistro("CodTipoGasto") = Codigo_Aplica_Devengo_Inmediata Then
'                If CDate(adoRegistro("FechaFinal")) >= dtpFechaPago.Value Then
'                    dtpFechaPago.Value = adoRegistro("FechaFinal")
'                    'dtpFechaPago.MinDate = dtpFechaPago.Value 'acr
'                End If
'            Else
'                If DateAdd("d", 1, adoRegistro("FechaFinal")) >= dtpFechaPago.Value Then
'                    dtpFechaPago.Value = DateAdd("d", 1, adoRegistro("FechaFinal"))
'                    'dtpFechaPago.MinDate = dtpFechaPago.Value 'acr
'                End If
'            End If
            
'            intRegistro = ObtenerItemLista(arrMonedaUnico(), adoRegistro("CodMoneda"))
'            If intRegistro >= 0 Then cboMonedaUnico.ListIndex = intRegistro
'        End If
'        adoRegistro.Close
        
'        If Trim(tdgPendientes.Columns(9).Value) = "099" Or Trim(tdgPendientes.Columns(9).Value) <> "098" Then
'            .CommandText = "SELECT CodDetalleFile FROM InversionDetalleFile " & _
'                "WHERE CodFile='" & Trim(tdgPendientes.Columns(9).Value) & "' AND DescripDetalleFile='" & strCodGasto & "'"
'        Else
'            '.CommandText = "SELECT CodDetalleFile FROM InversionDetalleFile " & _
'            '    "WHERE CodFile='" & Trim(tdgPendientes.Columns(8).Value) & "' AND CodDetalleFile='" & strCodGasto & "'"
'            .CommandText = "SELECT CodDetalleFile FROM DinamicaContable " & _
'                "WHERE CodFile='" & Trim(tdgPendientes.Columns(9).Value) & "' AND CodCuenta='" & strCodGasto & "'"
'        End If
'        Set adoRegistro = .Execute
'
'        If Not adoRegistro.EOF Then
'            strCodDetalleGasto = adoRegistro("CodDetalleFile")
'        End If
'        adoRegistro.Close: Set adoRegistro = Nothing
'    End With
    
End Sub

Private Sub cboMoneda_Click()

    strCodMoneda = Valor_Caracter
    If cboMoneda.ListIndex < 0 Then Exit Sub
    
    strCodMoneda = arrMoneda(cboMoneda.ListIndex)
    
End Sub


Private Sub cboMonedaDetraccion_Click()

    strCodMonedaDetraccion = Valor_Caracter
    If cboMonedaDetraccion.ListIndex < 0 Then Exit Sub
    
    strCodMonedaDetraccion = arrMonedaDetraccion(cboMonedaDetraccion.ListIndex)
    
End Sub


Private Sub cboMonedaUnico_Click()

    strCodMonedaUnico = Valor_Caracter
    If cboMonedaUnico.ListIndex < 0 Then Exit Sub
    
    strCodMonedaUnico = arrMonedaUnico(cboMonedaUnico.ListIndex)
    
End Sub


Private Sub cboTipoComprobante_Click()

    Dim adoRegistro     As ADODB.Recordset
    Dim intRegistro     As Long
    
    strCodTipoComprobante = Valor_Caracter
    If cboTipoComprobante.ListIndex < 0 Then Exit Sub
    
    strCodTipoComprobante = arrTipoComprobante(cboTipoComprobante.ListIndex)
    
    Set adoRegistro = New ADODB.Recordset
    strIndImpuesto = Valor_Caracter: strIndRetencion = Valor_Caracter
    With adoComm
        .CommandText = "SELECT IndImpuesto,IndRetencion,DescripCampo1,DescripCampo2,DescripCampo3 " & _
            "FROM TipoComprobantePago WHERE CodTipoComprobantePago='" & strCodTipoComprobante & "'        "
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            strIndImpuesto = Trim(adoRegistro("IndImpuesto"))
            strIndRetencion = Trim(adoRegistro("IndRetencion"))
            lblDescrip(7).Caption = Trim(adoRegistro("DescripCampo1"))
            lblDescrip(9).Caption = Trim(adoRegistro("DescripCampo2"))
            lblDescrip(10).Caption = Trim(adoRegistro("DescripCampo3"))
            
            strCtaImpuesto = ObtenerCuentaAdministracion("025", "R")
            If strIndRetencion = Valor_Indicador Then
                strCtaImpuesto = ObtenerCuentaAdministracion("036", "R")
'                If strCodMoneda <> Codigo_Moneda_Local Then
'                    If (CCur(txtSubTotal.Text) * CDbl(txtTipoCambioPago.Text)) > gcurMontoMaximoRetencion Then
'                        intRegistro = ObtenerItemLista(arrAfectacion(), Codigo_Respuesta_Si)
'                    Else
'                        intRegistro = ObtenerItemLista(arrAfectacion(), Codigo_Respuesta_No)
'                    End If
'                Else
'                    If CCur(txtSubTotal.Text) > gcurMontoMaximoRetencion Then
'                        intRegistro = ObtenerItemLista(arrAfectacion(), Codigo_Respuesta_Si)
'                    Else
'                        intRegistro = ObtenerItemLista(arrAfectacion(), Codigo_Respuesta_No)
'                    End If
'                End If
'                cboAfectacion.ListIndex = intRegistro
            End If
        
        
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    'Call cboGasto_Click
    
    'Call cboDetraccion_Click
    
    Call Calculos
   
End Sub


Private Sub cboTipoValorCambio_Click()

    Dim datFechaConsulta    As Date
    
    strCodValorTipoCambio = Valor_Caracter
    If cboTipoValorCambio.ListIndex < 0 Then Exit Sub
    
    strCodValorTipoCambio = arrTipoValorCambio(cboTipoValorCambio.ListIndex)
    'datFechaConsulta = gdatFechaActual
    datFechaConsulta = dtpFechaPago.Value
    
    dtpFechaTipoCambioPago.Value = datFechaConsulta
    
    If strCodValorTipoCambio = Codigo_Valor_TipoCambioCompra Then
        'txtTipoCambioPago.Text = CStr(ObtenerTipoCambio(Codigo_TipoCambio_Sunat, Codigo_Valor_TipoCambioCompra, datFechaConsulta, strCodMoneda))
        'Para el caso de detraccion debe tomar el tipo de cambio de la fecha del comprobante
        'If strDetraccionSiNo = Codigo_Respuesta_Si And dtpFechaPago.Value <> dtpFechaComprobante.Value Then
            txtTipoCambioPago.Text = CStr(ObtenerTipoCambioMoneda(Codigo_TipoCambio_SBS, Codigo_Valor_TipoCambioCompra, dtpFechaComprobante.Value, strCodMoneda, Codigo_Moneda_Local))
            dtpFechaTipoCambioPago.Value = dtpFechaComprobante.Value
        'End If
    Else
        'txtTipoCambioPago.Text = CStr(ObtenerTipoCambio(Codigo_TipoCambio_Sunat, Codigo_Valor_TipoCambioVenta, datFechaConsulta, strCodMoneda))
        'Debe tomar el tipo de cambio de la fecha del comprobante si es factura
        'If strDetraccionSiNo = Codigo_Respuesta_Si And dtpFechaPago.Value <> dtpFechaComprobante.Value And strCodTipoComprobante = Codigo_Comprobante_Factura Then
        'If dtpFechaPago.Value <> dtpFechaComprobante.Value And strCodTipoComprobante = Codigo_Comprobante_Factura Then
            txtTipoCambioPago.Text = CStr(ObtenerTipoCambioMoneda(Codigo_TipoCambio_SBS, Codigo_Valor_TipoCambioVenta, dtpFechaComprobante.Value, strCodMoneda, Codigo_Moneda_Local))
            dtpFechaTipoCambioPago.Value = dtpFechaComprobante.Value
        'End If
    End If
    
    'cboDetraccion_Click
    
End Sub

Private Sub cmdAccion_Click()

End Sub

Private Sub cmdContabilizar_Click()
    
    Dim adoRegistro         As ADODB.Recordset
    Dim adoAuxiliar         As ADODB.Recordset
    Dim strNumCaja          As String
    Dim strCodDetalleFile   As String, strCodMonedaGasto        As String
    Dim strDescripGasto     As String, strSQLOrdenCajaDetalleI  As String
    Dim strSQLOrdenCaja     As String, strSQLOrdenCajaDetalle   As String
    Dim strSQLOrdenCajaMN   As String, strSQLOrdenCajaDetalleMN As String
    Dim strFechaAnterior    As String, strFechaSiguiente        As String
    Dim curSaldoProvision   As Currency, intCantRegistros       As Integer
    Dim dblTipCambio        As Double, dblTipoCambioGasto   As Double
    Dim datFechaFinPeriodo  As Date
    Dim strNumComprobante   As String
    Dim numRegistro         As Long
    
    
    
    
    
    
      If Not ValidarDinamicaGasto() Then Exit Sub
    
        If MsgBox("Desea contabilizar el gasto?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub
    
        Me.MousePointer = vbHourglass
        
        '*** Guardar ***
        With adoComm
    
            .CommandText = "{ call up_CNProcGastoProveedor( '" & gstrLogin & "','" & gstrFechaActual & _
             "','" & strCodFondo & "','" & gstrCodAdministradora & "','" & Trim(lblNumSecuencial.Caption) & "')}"
            adoComm.Execute .CommandText
    
        End With
    
        Me.MousePointer = vbDefault
                    
        MsgBox Mensaje_Adicion_Exitosa, vbExclamation
        
        frmMainMdi.stbMdi.Panels(3).Text = "Acción"
        
        cmdOpcion.Visible = True
        With tabRegistroCompras
            .TabEnabled(0) = True
            .TabEnabled(1) = False
            .TabEnabled(2) = False
            .Tab = 0
        End With
        Call CargarPendientes
        Call Buscar

    
    
    
    
    
    
    
    
    
    
'    If ValidarDinamicaGasto() = False Then Exit Sub
'
'     If MsgBox("Desea contabilizar el gasto?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub
'
'       Me.MousePointer = vbHourglass
'
'
''       strNumComprobante = Trim(txtSerieComprobante.Text) & "-" & Trim(txtNumComprobante.Text)
'
'        '*** Guardar ***
'        With adoComm
'
'            If strCodMoneda <> Codigo_Moneda_Local Then
'                dblTipoCambioGasto = ObtenerTipoCambioMoneda(Codigo_TipoCambio_Sunat, Codigo_Valor_TipoCambioVenta, dtpFechaComprobante.Value, strCodMoneda, Codigo_Moneda_Local)
'            Else
'                dblTipoCambioGasto = 1
'            End If
'      End With
'
'
'            '            '*** Generar Movimiento Contable de Impuesto ***
'            If strCodAplicacionDevengo = Codigo_Aplica_Devengo_Inmediata Then
'                Call GenerarAsientoGasto(strCodFile, strCodAnalitica, strCodFondo, gstrCodAdministradora, strCodDetalleGasto, Codigo_Dinamica_Gasto, CCur(txtSubTotal.Text), CCur(txtMontoNoGravado.Value), dblTipoCambioGasto, strCodMoneda, Trim(txtDescripcion.Text), frmMainMdi.Tag, strCodTipoComprobante, strNumComprobante, strCodAfectacion, Codigo_Tipo_Persona_Proveedor, Trim(lblCodProveedor.Caption))
'                Call GenerarAsientoGasto(strCodFile, strCodAnalitica, strCodFondo, gstrCodAdministradora, strCodDetalleGasto, Codigo_Dinamica_Gasto_Emitida, CCur(txtSubTotal.Text), CCur(txtMontoNoGravado.Value), dblTipoCambioGasto, strCodMoneda, Trim(txtDescripcion.Text), frmMainMdi.Tag, strCodTipoComprobante, strNumComprobante, strCodAfectacion, Codigo_Tipo_Persona_Proveedor, Trim(lblCodProveedor.Caption))
'            End If
'
'            If strCodAplicacionDevengo = Codigo_Aplica_Devengo_Periodica Then
'                Call GenerarAsientoGasto(strCodFile, strCodAnalitica, strCodFondo, gstrCodAdministradora, strCodDetalleGasto, Codigo_Dinamica_Gasto_Emitida, CCur(txtSubTotal.Text), CCur(txtMontoNoGravado.Value), dblTipoCambioGasto, strCodMoneda, Trim(txtDescripcion.Text), frmMainMdi.Tag, strCodTipoComprobante, strNumComprobante, strCodAfectacion, Codigo_Tipo_Persona_Proveedor, Trim(lblCodProveedor.Caption))
'            End If
'
'            '*** Generar Orden si no está generada o actualizar ***
'            Call GenerarOrdenGastosFondo(strCodDetalleGasto, strCodGasto, strCodFondo, CInt(tdgPendientes.Columns("NumGasto").Value), Trim(lblCodProveedor.Caption), numRegistro)
'
'
'            Me.MousePointer = vbDefault
'
'            MsgBox Mensaje_Adicion_Exitosa, vbExclamation
'
'            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
'
'            cmdOpcion.Visible = True
'            With tabRegistroCompras
'                .TabEnabled(0) = True
'                .TabEnabled(1) = False
'                .TabEnabled(2) = False
'                .Tab = 0
'            End With
    
End Sub

Private Sub cmdProveedor_Click()

    'gstrFormulario = "frmRegistroCompras"
    'frmBusquedaInstitucionPersona.lblTipoInstitucion = Codigo_Tipo_Persona_Proveedor
    'frmBusquedaInstitucionPersona.Caption = "Búsqueda de Proveedores"
    'frmBusquedaInstitucionPersona.Show vbModal
   
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
        .sSql = "{ call up_ACSelDatos(26) }"
        
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
            lblDireccion.Caption = .iParams(6).Valor
            lblCodProveedor.Caption = .iParams(1).Valor
        End If
            
       
    End With
    
    Set frmBus = Nothing

    
End Sub

Private Sub dtpFechaComprobante_Change()

    If dtpFechaComprobante.Value > gdatFechaActual Then
        MsgBox "La Fecha de comprobante debe ser igual o anterior a la fecha actual...se cambiará por la fecha actual !", vbInformation, Me.Caption
        dtpFechaComprobante.Value = gdatFechaActual
    End If
    
    Call cboDetraccion_Click

End Sub



Private Sub dtpFechaPago_Change()

    If Not EsDiaUtil(dtpFechaPago.Value) Then
        MsgBox "La Fecha no es un día útil...se cambiará por una fecha correcta !", vbInformation, Me.Caption
        If dtpFechaPago.Value >= gdatFechaActual Then
            dtpFechaPago.Value = AnteriorDiaUtil(dtpFechaPago.Value)
        Else
            dtpFechaPago.Value = ProximoDiaUtil(dtpFechaPago.Value)
        End If
    End If
    
    If dtpFechaPago.Value < dtpFechaRegistro.Value Then
        dtpFechaPago.Value = dtpFechaRegistro.Value
        MsgBox "La Fecha de Pago no debe ser anterior la Fecha del Sistema!", vbInformation, Me.Caption
        Exit Sub
    End If

    'Call cboDetraccion_Click
    
End Sub

Private Sub Form_Activate()

    Call CargarReportes
    
End Sub

Private Sub CargarReportes()

'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Registro de Compras - DAOT"
'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Text = "Registro de Ventas"
'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo3").Text = "Registro de Compras"
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo3").Visible = True
'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo4").Text = "Registro de Recibo por Honorarios"
'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo4").Visible = True
'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Text = "Registro de Compras - Parte2"
'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Visible = True
    
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
    
    ' /* */
    ' Bind True DBGrid Control to this XArray instance   - tdgPendientes - TDBGrid1
    Set tdgPendientes.Array = X
    
    ' Habilitamos los pies de columna ' footers
    '   TDBGrid1.ColumnFooters = True
    
    ' Mostramos las cabeceras y pies del grid como botones'    headers and footers as buttons
        Dim obcol As TrueOleDBGrid60.Column
        For Each obcol In tdgPendientes.Columns
            'obcol.ButtonFooter = True
            obcol.ButtonHeader = True
        Next obcol
    
    ' /* */
        
    Call ValidaExisteTipoCambio(Codigo_TipoCambio_SBS, gstrFechaActual)
    
End Sub


Private Sub DarFormato()

    Dim intCont As Integer
    Dim elemento As Object
    
    For intCont = 0 To (lblDescrip.Count - 1)
        Call FormatoEtiqueta(lblDescrip(intCont), vbLeftJustify)
    Next
    
    For intCont = 0 To (fraCompras.Count - 1)
        Call FormatoMarco(fraCompras(intCont))
    Next
    
    For Each elemento In Me.Controls
    
        If TypeOf elemento Is TDBGrid Then
            Call FormatoGrilla(elemento)
        End If
    
    Next
            
End Sub
Public Sub Buscar()
            
    Set adoConsulta1 = New ADODB.Recordset
            
    strSQL = "SELECT NumRegistro,CodTipoComprobante,CodProveedor,DescripRegistro,RC.CodMoneda,ValorTotal,RC.CodFileGasto, " & _
        "TCP.DescripTipoComprobantePago DescripTipoComprobante, CodSigno,FechaRegistro,DescripPersona DescripProveedor,RC.NumGasto " & _
        "FROM RegistroCompra RC JOIN TipoComprobantePago TCP ON(TCP.CodTipoComprobantePago=RC.CodTipoComprobante) " & _
        "JOIN Moneda MON ON(MON.CodMoneda=RC.CodMoneda) " & _
        "JOIN InstitucionPersona IP ON(IP.CodPersona=RC.CodProveedor AND IP.TipoPersona=RC.TipoProveedor) " & _
        "WHERE (FechaRegistro>='" & Convertyyyymmdd(dtpFechaDesde.Value) & "' AND FechaRegistro<'" & Convertyyyymmdd(DateAdd("d", 1, dtpFechaHasta.Value)) & "') AND " & _
        "CodAdministradora='" & gstrCodAdministradora & "' AND CodFondo='" & strCodFondo & "' AND RC.Estado='01' " & _
        "ORDER BY NumRegistro"

     strEstado = Reg_Defecto
    
    With adoConsulta1
'        .ConnectionString = gstrConnectConsulta
'        .RecordSource = strSQL
'        .Refresh
        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open strSQL
    End With
    
    tdgConsulta.DataSource = adoConsulta1
    
    Call AutoAjustarGrilla
    Call CargarPendientes
    
    tdgConsulta.Refresh


    If adoConsulta1.RecordCount > 0 Then strEstado = Reg_Consulta
    dtpFechaPago.MinDate = 0
            
End Sub
Private Sub CargarListas()
            
    '*** Fondos ***
    strSQL = "{ call up_ACSelDatosParametro(74,'" & gstrCodAdministradora & "','" & gstrCodFondoContable & "') }"
    CargarControlLista strSQL, cboFondo, arrFondo(), Valor_Caracter
    
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
    
    '*** Tipo de Comprobante Sunat ***
    strSQL = "SELECT CodTipoComprobantePago CODIGO,DescripTipoComprobantePago DESCRIP From TipoComprobantePago ORDER BY DescripTipoComprobantePago"
    CargarControlLista strSQL, cboTipoComprobante, arrTipoComprobante(), Sel_Defecto
            
    '*** Afectación ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP From AuxiliarParametro WHERE CodTipoParametro='AFEIMP' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboAfectacion, arrAfectacion(), Valor_Caracter
    
    '*** Tipo Crédito Fiscal ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP From AuxiliarParametro WHERE CodTipoParametro='CREFIS' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboCreditoFiscal, arrCreditoFiscal(), Sel_Defecto
    
    '*** Moneda ***
    strSQL = "{ call up_ACSelDatos(2) }"
    CargarControlLista strSQL, cboMoneda, arrMoneda(), Valor_Caracter
    CargarControlLista strSQL, cboMonedaUnico, arrMonedaUnico(), Valor_Caracter
    CargarControlLista strSQL, cboMonedaDetraccion, arrMonedaDetraccion(), Valor_Caracter
    
    '*** Detracción ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP From AuxiliarParametro WHERE CodTipoParametro='RESPSN' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboDetraccion, arrDetraccion(), ""
        
    '*** Forma de Pago ***
'    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='MEDPAG' ORDER BY DescripParametro"
'    CargarControlLista strSQL, cboFormaPagoUnico, arrFormaPagoUnico(), Valor_Caracter
'    CargarControlLista strSQL, cboFormaPagoDetraccion, arrFormaPagoDetraccion(), Valor_Caracter
    
    '*** Valor de Tipo de Cambio ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP From AuxiliarParametro WHERE CodTipoParametro='CLSVTC' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboTipoValorCambio, arrTipoValorCambio(), ""
    
End Sub
Private Sub InicializarValores()

    strEstado = Reg_Defecto
    tabRegistroCompras.Tab = 0
    
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    Set cmdOpcion.FormularioActivo = Me
    
End Sub
Private Sub Form_Unload(Cancel As Integer)

    Call OcultarReportes
    frmMainMdi.stbMdi.Panels(3).Text = "Acción"
    Set frmRegistroCompras = Nothing
    
End Sub


Private Sub lblMontoGasto_Change()

    Call FormatoMillarEtiqueta(lblMontoGasto, Decimales_Monto)

End Sub

Private Sub lblMontoTotal_Change()

    Call FormatoMillarEtiqueta(lblMontoTotal, Decimales_Monto)
    
End Sub

Private Sub tabRegistroCompras_Click(PreviousTab As Integer)

    Select Case tabRegistroCompras.Tab
        Case 1, 2
            If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vQuery)
            If strEstado = Reg_Defecto Then tabRegistroCompras.Tab = 0
        
    End Select
    
End Sub


Private Sub tdgConsulta_Click()

    tdgConsulta.HeadBackColor = &HFFC0C0
    tdgPendientes.HeadBackColor = &H8000000F
    
End Sub

Private Sub tdgConsulta_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)

    If ColIndex = 5 Then
        Call DarFormatoValor(Value, Decimales_Monto)
    End If
    
End Sub

Private Sub tdgPendientes_Click()

    tdgPendientes.HeadBackColor = &HFFC0C0
    tdgConsulta.HeadBackColor = &H8000000F
    
End Sub

Private Sub tdgPendientes_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)

    If ColIndex = 7 Or ColIndex = 8 Then
        Call DarFormatoValor(Value, Decimales_Monto)
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

    Call OrdenarDBGrid(ColIndex, adoConsulta1, tdgConsulta)
    
    numColindex = ColIndex
    
    '****
    strPrevColumTDB = strColNameTDB
    '***
    
End Sub

' /**/
'Private Sub tdgPendientes_HeadClick(ByVal ColIndex As Integer)
' Ascending sort
'    x.QuickSort x.LowerBound(1), x.UpperBound(1), ColIndex, XORDER_ASCEND, XTYPE_INTEGER
'    tdgPendientes.Refresh
'End Sub
' /**/


Private Sub txtIgv_Change()

    Call FormatoCajaTexto(txtIgv, Decimales_Monto)
    
End Sub

Private Sub txtIgv_KeyPress(KeyAscii As Integer)

'    Call ValidaCajaTexto(KeyAscii, "M", txtIgv, Decimales_Monto)
'    If KeyAscii = vbKeyReturn Then Call Calculos
    
End Sub

Private Sub txtMontoDetraccion_Change()

    Call FormatoCajaTexto(txtMontoDetraccion, Decimales_Monto)
    
End Sub

Private Sub txtMontoDetraccion_KeyPress(KeyAscii As Integer)

'    Call ValidaCajaTexto(KeyAscii, "M", txtMontoDetraccion, Decimales_Monto)
'    If KeyAscii = vbKeyReturn Then Call CalculosPago
    
End Sub

Private Sub txtMontoNoGravado_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then Call Calculos

End Sub

Private Sub txtMontoUnico_Change()

    Call FormatoCajaTexto(txtMontoUnico, Decimales_Monto)
    
End Sub

Private Sub txtMontoUnico_KeyPress(KeyAscii As Integer)

'    Call ValidaCajaTexto(KeyAscii, "M", txtMontoUnico, Decimales_Monto)
'    If KeyAscii = vbKeyReturn Then Call CalculosPago
    
End Sub

Private Sub txtSubTotal_Change()

    'Call FormatoCajaTexto(txtSubTotal, Decimales_Monto)
     Call Calculos
     
End Sub

Private Sub txtSubTotal_KeyPress(KeyAscii As Integer)

    'Call ValidaCajaTexto(KeyAscii, "M", txtSubTotal, Decimales_Monto)
    If KeyAscii = vbKeyReturn Then Call Calculos
    
End Sub

Private Sub txtTipoCambioPago_Change()

    Call FormatoCajaTexto(txtTipoCambioPago, Decimales_TipoCambio)
    
End Sub

Private Sub txtTipoCambioPago_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtTipoCambioPago, Decimales_TipoCambio)
    
End Sub

Private Sub txtTotal_Change()

    Call FormatoCajaTexto(txtTotal, Decimales_Monto)
    
End Sub

Private Sub txtTotal_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtTotal, Decimales_Monto)
    If KeyAscii = vbKeyReturn Then Call Calculos
    
End Sub

Private Function ValidarDinamicaGasto() As Boolean

    
    Dim strParametroFormulaXML As String
    Dim objParametroFormulaXML As DOMDocument60
    Dim strMsgError As String
    Dim adoRegistroAux As ADODB.Recordset
    Dim adoPendienteTmp As ADODB.Recordset

    ValidarDinamicaGasto = False

'    If strCodAfectacion = Codigo_Afecto Then
'        If strIndImpuesto = Valor_Indicador Then
'            If strDetraccionSiNo = Codigo_Respuesta_Si Then
'
'                adoRegistroAux.AddNew
'
'                adoRegistroAux.Fields("TipoOperacion") = "28"
'                adoRegistroAux.Fields("CodFile") = strCodFile
'                adoRegistroAux.Fields("CodDetalleFile") = strCodDetalleGasto
'                adoRegistroAux.Fields("CodSubDetalleFile") = 0
'                adoRegistroAux.Fields("CodMoneda") = strCodMoneda
'                adoRegistroAux.Fields("CodAfectacion") = strCodAfectacion
'                adoRegistroAux.Fields("CodDetraccionSiNo") = strDetraccionSiNo
'                adoRegistroAux.Fields("CodCreditoFiscal") = strCodCreditoFiscal
'                adoRegistroAux.Fields("TipoImpuesto") = Codigo_Impuesto_IGV
'
'                If strCodAplicacionDevengo = Codigo_Aplica_Devengo_Inmediata Then
'                    adoRegistroAux.Fields("DocumentoEmitido") = Codigo_Documento_Emitido
'                Else
'                    adoRegistroAux.Fields("DocumentoEmitido") = Codigo_Documento_No_Emitido
'                End If
'
'                Call XMLADORecordset(objParametroFormulaXML, "ParametroFormula", "Parametro", adoRegistroAux, strMsgError)
'                strParametroFormulaXML = objParametroFormulaXML.xml 'CrearXMLDetalle(objTipoCambioReemplazoXML)
'
'                If Not ExisteDinamica(gstrCodAdministradora, strParametroFormulaXML) Then
'                    MsgBox "No Existe Dinámica de Liquidación de Gasto con Credito Fiscal!", vbExclamation, Me.Caption
'                    Exit Function
'                End If
'
'                adoRegistroAux.MoveFirst
'                adoRegistroAux.Delete
'
'                adoRegistroAux.AddNew
'
'                adoRegistroAux.Fields("TipoOperacion") = Codigo_Dinamica_Liquidacion_Gasto_Detraccion
'                adoRegistroAux.Fields("CodFile") = strCodFile
'                adoRegistroAux.Fields("CodDetalleFile") = strCodDetalleGasto
'                adoRegistroAux.Fields("CodSubDetalleFile") = 0
'                adoRegistroAux.Fields("CodMoneda") = strCodMoneda
'                adoRegistroAux.Fields("CodAfectacion") = strCodAfectacion
'                adoRegistroAux.Fields("CodDetraccionSiNo") = strDetraccionSiNo
'                adoRegistroAux.Fields("CodCreditoFiscal") = strCodCreditoFiscal
'                adoRegistroAux.Fields("TipoImpuesto") = Codigo_Impuesto_IGV
'
'                If strCodAplicacionDevengo = Codigo_Aplica_Devengo_Inmediata Then
'                    adoRegistroAux.Fields("DocumentoEmitido") = Codigo_Documento_Emitido
'                Else
'                    adoRegistroAux.Fields("DocumentoEmitido") = Codigo_Documento_No_Emitido
'                End If
'
'
'                Call XMLADORecordset(objParametroFormulaXML, "ParametroFormula", "Parametro", adoRegistroAux, strMsgError)
'                strParametroFormulaXML = objParametroFormulaXML.xml 'CrearXMLDetalle(objTipoCambioReemplazoXML)
'
'                If Not ExisteDinamica(gstrCodAdministradora, strParametroFormulaXML) Then
'                    MsgBox "No Existe Dinámica de Liquidación de Detracción con Credito Fiscal!", vbExclamation, Me.Caption
'                    Exit Function
'                End If
'
'                adoRegistroAux.MoveFirst
'                adoRegistroAux.Delete
'            Else 'SIN DETRACCION
'
'                adoRegistroAux.AddNew
'
'                adoRegistroAux.Fields("TipoOperacion") = Codigo_Dinamica_Liquidacion_Gasto_Proveedor
'                adoRegistroAux.Fields("CodFile") = strCodFile
'                adoRegistroAux.Fields("CodDetalleFile") = strCodDetalleGasto
'                adoRegistroAux.Fields("CodSubDetalleFile") = 0
'                adoRegistroAux.Fields("CodMoneda") = strCodMoneda
'                adoRegistroAux.Fields("CodAfectacion") = strCodAfectacion
'                adoRegistroAux.Fields("CodDetraccionSiNo") = strDetraccionSiNo
'                adoRegistroAux.Fields("CodCreditoFiscal") = strCodCreditoFiscal
'                adoRegistroAux.Fields("TipoImpuesto") = Codigo_Impuesto_IGV
'
'                If strCodAplicacionDevengo = Codigo_Aplica_Devengo_Inmediata Then
'                    adoRegistroAux.Fields("DocumentoEmitido") = Codigo_Documento_Emitido
'                Else
'                    adoRegistroAux.Fields("DocumentoEmitido") = Codigo_Documento_No_Emitido
'                End If
'
'                Call XMLADORecordset(objParametroFormulaXML, "ParametroFormula", "Parametro", adoRegistroAux, strMsgError)
'                strParametroFormulaXML = objParametroFormulaXML.xml 'CrearXMLDetalle(objTipoCambioReemplazoXML)
'
'                If Not ExisteDinamica(gstrCodAdministradora, strParametroFormulaXML) Then
'                    MsgBox "No Existe Dinámica de Liquidación de Gasto con Credito Fiscal!", vbExclamation, Me.Caption
'                    Exit Function
'                End If
'
'                adoRegistroAux.MoveFirst
'                adoRegistroAux.Delete
'            End If
''       Else
''           If strIndRetencion = Valor_Indicador Then
''               If strCodDetraccionSiNo = Codigo_Respuesta_Si Then
''
''               Else
''
''               End If
''           Else
''
''           End If
'        End If
'    Else ' @CodAfectacion = Codigo_Inafecto
'
'        adoRegistroAux.AddNew
'
'        adoRegistroAux.Fields("TipoOperacion") = Codigo_Dinamica_Liquidacion_Gasto_Proveedor
'        adoRegistroAux.Fields("CodFile") = strCodFile
'        adoRegistroAux.Fields("CodDetalleFile") = strCodDetalleGasto
'        adoRegistroAux.Fields("CodSubDetalleFile") = 0
'        adoRegistroAux.Fields("CodMoneda") = strCodMoneda
'        adoRegistroAux.Fields("CodAfectacion") = strCodAfectacion
'        adoRegistroAux.Fields("CodDetraccionSiNo") = strDetraccionSiNo
'        adoRegistroAux.Fields("CodCreditoFiscal") = strCodCreditoFiscal
'        adoRegistroAux.Fields("TipoImpuesto") = Codigo_Impuesto_Sin_Impuesto
'
'        If strCodAplicacionDevengo = Codigo_Aplica_Devengo_Inmediata Then
'            adoRegistroAux.Fields("DocumentoEmitido") = Codigo_Documento_Emitido
'        Else
'            adoRegistroAux.Fields("DocumentoEmitido") = Codigo_Documento_No_Emitido
'        End If
'
'        Call XMLADORecordset(objParametroFormulaXML, "ParametroFormula", "Parametro", adoRegistroAux, strMsgError)
'        strParametroFormulaXML = objParametroFormulaXML.xml 'CrearXMLDetalle(objTipoCambioReemplazoXML)
'
'        If Not ExisteDinamica(gstrCodAdministradora, strParametroFormulaXML) Then
'            MsgBox "No Existe Dinámica de Liquidación de Gasto!", vbExclamation, Me.Caption
'            Exit Function
'        End If
'
'        adoRegistroAux.MoveFirst
'        adoRegistroAux.Delete
'
'
'    End If
'
'    ValidarDinamicaGasto = True
'
'


          strSQL = "SELECT RC.CodFileGasto FROM RegistroCompra RC  " & _
                            "WHERE RC.NumRegistro= '" & Trim(lblNumSecuencial.Caption) & "' "

                    With adoComm
                        .CommandText = strSQL
                        Set adoPendienteTmp = .Execute
                    End With


         strCodFile = adoPendienteTmp("CodFileGasto").Value

        If strCodAplicacionDevengo = Codigo_Aplica_Devengo_Inmediata Then
            If Not ExisteDinamica(strCodFile, strCodDetalleGasto, gstrCodAdministradora, Codigo_Dinamica_Gasto, strCodMoneda) Then Exit Function
            If Not ExisteDinamica(strCodFile, "000", gstrCodAdministradora, Codigo_Dinamica_Gasto_Emitida, strCodMoneda) Then Exit Function
        End If

        If strCodAplicacionDevengo = Codigo_Aplica_Devengo_Periodica Then
            If Not ExisteDinamica(strCodFile, "000", gstrCodAdministradora, Codigo_Dinamica_Gasto_Emitida, strCodMoneda) Then Exit Function
        End If


        If strDetraccionSiNo = Codigo_Respuesta_Si Then
            If strIndImpuesto = Valor_Indicador Then
                If CDec(txtMontoDetraccion.Text) = Round(txtMontoDetraccion.Text) Then
                    If Not ExisteDinamica(strCodFile, strCodDetalleGasto, gstrCodAdministradora, Codigo_Dinamica_Detraccion, strCodMoneda) Then Exit Function
                ElseIf CDec(txtMontoDetraccion.Text) > Round(txtMontoDetraccion.Text) Then
                    If Not ExisteDinamica(strCodFile, strCodDetalleGasto, gstrCodAdministradora, Codigo_Dinamica_Detraccion_Ajuste_Redondeo_Ganancia, strCodMoneda) Then Exit Function
                ElseIf CDec(txtMontoDetraccion.Text) < Round(txtMontoDetraccion.Text) Then
                    If Not ExisteDinamica(strCodFile, strCodDetalleGasto, gstrCodAdministradora, Codigo_Dinamica_Detraccion_Ajuste_Redondeo_Perdida, strCodMoneda) Then Exit Function
                End If
            ElseIf strIndRetencion = Valor_Indicador Then
                If Not ExisteDinamica(strCodFile, strCodDetalleGasto, gstrCodAdministradora, Codigo_Dinamica_Retencion, strCodMoneda) Then Exit Function
            End If
        End If

    ValidarDinamicaGasto = True


End Function


Private Sub AutoAjustarGrilla()

    Dim i As Integer, j As Integer

    If Not adoConsulta1.EOF Then
        If adoConsulta1.RecordCount > 0 Then
            For i = 1 To tdgConsulta.Columns.Count - 1
            tdgConsulta.Columns(i).AutoSize
            Next
            
            tdgConsulta.Columns(1).AutoSize
            tdgConsulta.Columns(7).AutoSize
        End If
    End If
    
    If Not adoPendientes1.EOF Then
        If adoPendientes1.RecordCount > 0 Then
            For i = 0 To tdgPendientes.Columns.Count - 1
            tdgPendientes.Columns(i).AutoSize
            Next
            
            tdgPendientes.Columns(0).AutoSize
            tdgPendientes.Columns(1).AutoSize
            tdgPendientes.Columns(9).AutoSize
        End If
    End If
    

End Sub

Private Sub tdgPendientes_HeadClick(ByVal ColIndex As Integer)
    
    Dim strColNameTDB  As String
    Static numColindex As Integer
    Static strPrevColumTDB As String
    '** agregar para que no se raye la seleccion de registro con ordenamiento
    strColNameTDB = tdgPendientes.Columns(ColIndex).DataField
    
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

    tdgPendientes.Splits(0).Columns(numColindex).HeadingStyle.ForegroundPicture = Null

    Call OrdenarDBGrid(ColIndex, adoPendientes1, tdgPendientes)
    
    numColindex = ColIndex
    
    '****
    strPrevColumTDB = strColNameTDB
    '***
    
End Sub
