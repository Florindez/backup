VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{830D5A06-9B70-4F8C-98B6-7A19C4E7760B}#1.0#0"; "TAMControls.ocx"
Begin VB.Form frmGenerarOrdenCobroProvision 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generar Orden Pago Provisi�n"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8430
   ScaleWidth      =   10545
   Begin TAMControls.ucBotonEdicion cmdContabilizar 
      Height          =   390
      Left            =   5820
      TabIndex        =   82
      Top             =   7290
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   688
      Caption0        =   "&Contabilizar"
      Tag0            =   "6"
      ToolTipText0    =   "Contabilizar"
      UserControlHeight=   390
      UserControlWidth=   1200
   End
   Begin TAMControls.ucBotonEdicion cmdAccion 
      Height          =   390
      Left            =   7350
      TabIndex        =   80
      Top             =   7290
      Visible         =   0   'False
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   688
      Buttons         =   2
      Caption0        =   "&Guardar"
      Tag0            =   "2"
      ToolTipText0    =   "Guardar"
      Caption1        =   "&Cancelar"
      Tag1            =   "8"
      ToolTipText1    =   "Cancelar"
      UserControlHeight=   390
      UserControlWidth=   2700
   End
   Begin TAMControls.ucBotonEdicion cmdSalir 
      Height          =   390
      Left            =   8580
      TabIndex        =   4
      Top             =   8010
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   688
      Caption0        =   "&Salir"
      Tag0            =   "9"
      ToolTipText0    =   "Salir"
      UserControlHeight=   390
      UserControlWidth=   1200
   End
   Begin TAMControls.ucBotonEdicion cmdOpcion 
      Height          =   390
      Left            =   690
      TabIndex        =   3
      Top             =   8010
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   688
      Buttons         =   4
      Caption0        =   "&Nuevo"
      Tag0            =   "0"
      ToolTipText0    =   "Nuevo"
      Caption1        =   "&Modificar"
      Tag1            =   "1"
      Enabled1        =   0   'False
      ToolTipText1    =   "Modificar"
      Caption2        =   "&Eliminar"
      Tag2            =   "4"
      ToolTipText2    =   "Eliminar"
      Caption3        =   "&Buscar"
      Tag3            =   "5"
      Enabled3        =   0   'False
      ToolTipText3    =   "Buscar"
      UserControlHeight=   390
      UserControlWidth=   5700
   End
   Begin MSAdodcLib.Adodc adoConsulta 
      Height          =   330
      Left            =   7050
      Top             =   8010
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
      Height          =   7920
      Left            =   0
      TabIndex        =   15
      Top             =   30
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   13970
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
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
      TabCaption(0)   =   "Lista"
      TabPicture(0)   =   "frmGenerarOrdenCobroProvision.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "gLista"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraCompras(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ucBotonNavegacion1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Datos Generales"
      TabPicture(1)   =   "frmGenerarOrdenCobroProvision.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraCompras(3)"
      Tab(1).Control(1)=   "fraCompras(1)"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Ordenes de Pago"
      TabPicture(2)   =   "frmGenerarOrdenCobroProvision.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Datos Tributarios"
      TabPicture(3)   =   "frmGenerarOrdenCobroProvision.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraCompras(2)"
      Tab(3).ControlCount=   1
      Begin VB.Frame fraCompras 
         Caption         =   "Definici�n de Pagos"
         Height          =   1605
         Index           =   3
         Left            =   -74640
         TabIndex        =   71
         Top             =   5280
         Width           =   9705
         Begin VB.ComboBox cboDetraccion 
            Height          =   315
            Left            =   7080
            Style           =   2  'Dropdown List
            TabIndex        =   75
            Tag             =   "0"
            Top             =   420
            Width           =   2295
         End
         Begin VB.TextBox txtPeriodoFiscal 
            Height          =   315
            Left            =   7080
            TabIndex        =   74
            Top             =   870
            Width           =   2295
         End
         Begin VB.ComboBox cboCreditoFiscal 
            Height          =   315
            ItemData        =   "frmGenerarOrdenCobroProvision.frx":0070
            Left            =   2520
            List            =   "frmGenerarOrdenCobroProvision.frx":0077
            Style           =   2  'Dropdown List
            TabIndex        =   73
            Top             =   885
            Width           =   2295
         End
         Begin VB.ComboBox cboAfectacion 
            Height          =   315
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   72
            Top             =   420
            Width           =   2295
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Periodo Registro Cr�dito Fiscal"
            Height          =   405
            Index           =   21
            Left            =   5400
            TabIndex        =   79
            Top             =   900
            Width           =   1455
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Cr�dito Fiscal"
            Height          =   195
            Index           =   20
            Left            =   360
            TabIndex        =   78
            Top             =   900
            Width           =   945
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Impuesto"
            Height          =   195
            Index           =   19
            Left            =   360
            TabIndex        =   77
            Top             =   495
            Width           =   645
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Retenci�n y/o Detracci�n"
            Height          =   405
            Index           =   16
            Left            =   5400
            TabIndex        =   76
            Top             =   390
            Width           =   1305
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Detalle del Pago"
         Height          =   6225
         Left            =   -74640
         TabIndex        =   59
         Top             =   720
         Width           =   9735
         Begin TrueOleDBGrid60.TDBGrid tdgGastos 
            Height          =   3465
            Left            =   720
            OleObjectBlob   =   "frmGenerarOrdenCobroProvision.frx":008D
            TabIndex        =   85
            Top             =   1110
            Width           =   8685
         End
         Begin VB.ComboBox cboGasto 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   69
            Top             =   480
            Width           =   7215
         End
         Begin VB.CommandButton cmdGasto 
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
            Height          =   285
            Left            =   9150
            TabIndex        =   68
            ToolTipText     =   "Buscar Proveedor"
            Top             =   480
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.TextBox txtSubTotal 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   7140
            TabIndex        =   62
            Top             =   4710
            Width           =   2295
         End
         Begin VB.TextBox txtIgv 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   7140
            TabIndex        =   61
            Top             =   5130
            Width           =   2295
         End
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   7140
            TabIndex        =   60
            Top             =   5520
            Width           =   2295
         End
         Begin DXDBGRIDLibCtl.dxDBGrid gGastos 
            Height          =   3405
            Left            =   780
            OleObjectBlob   =   "frmGenerarOrdenCobroProvision.frx":638B
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   1140
            Visible         =   0   'False
            Width           =   8640
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Orden de Pago"
            Height          =   195
            Index           =   14
            Left            =   270
            TabIndex        =   70
            Top             =   510
            Width           =   1080
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Valor de Venta"
            Height          =   195
            Index           =   7
            Left            =   5160
            TabIndex        =   67
            Top             =   4785
            Width           =   1050
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "IGV"
            Height          =   195
            Index           =   9
            Left            =   5160
            TabIndex        =   66
            Top             =   5160
            Width           =   270
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Precio de Venta"
            Height          =   195
            Index           =   10
            Left            =   5160
            TabIndex        =   65
            Top             =   5535
            Width           =   1140
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Anal�tica"
            Height          =   195
            Index           =   24
            Left            =   210
            TabIndex        =   64
            Top             =   6600
            Width           =   630
         End
         Begin VB.Label lblAnalitica 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   570
            TabIndex        =   63
            Top             =   6480
            Width           =   2295
         End
      End
      Begin VB.Frame fraCompras 
         Caption         =   "Definici�n de Obligaci�n"
         Height          =   5445
         Index           =   2
         Left            =   -74640
         TabIndex        =   40
         Top             =   720
         Width           =   9675
         Begin VB.ComboBox cboPorcenDetraccion 
            Height          =   315
            ItemData        =   "frmGenerarOrdenCobroProvision.frx":7F66
            Left            =   2280
            List            =   "frmGenerarOrdenCobroProvision.frx":7F79
            Style           =   2  'Dropdown List
            TabIndex        =   84
            Top             =   2550
            Width           =   2295
         End
         Begin VB.ComboBox cboMonedaUnico 
            Height          =   315
            Left            =   2280
            Style           =   2  'Dropdown List
            TabIndex        =   45
            Top             =   990
            Width           =   2295
         End
         Begin VB.ComboBox cboMonedaDetraccion 
            Height          =   315
            Left            =   2280
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   2040
            Width           =   2295
         End
         Begin VB.TextBox txtMontoUnico 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   7050
            TabIndex        =   43
            Top             =   990
            Width           =   2295
         End
         Begin VB.TextBox txtMontoDetraccion 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   7080
            TabIndex        =   42
            Top             =   2010
            Width           =   2295
         End
         Begin VB.ComboBox cboTipoValorCambio 
            Height          =   315
            Left            =   2280
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   3510
            Visible         =   0   'False
            Width           =   2295
         End
         Begin MSComCtl2.DTPicker dtpFechaPago 
            Height          =   345
            Left            =   2280
            TabIndex        =   46
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
            Left            =   7080
            TabIndex        =   47
            Top             =   2520
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
         Begin TAMControls.TAMTextBox txtTipoCambioPago 
            Height          =   315
            Left            =   2280
            TabIndex        =   81
            Top             =   3030
            Width           =   2265
            _ExtentX        =   3995
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
            Container       =   "frmGenerarOrdenCobroProvision.frx":7F8F
            Text            =   "0.00000000"
            Decimales       =   8
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   12648447
            EnterTab        =   -1  'True
            AceptaNegativos =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   8
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Porcentaje"
            Height          =   195
            Index           =   33
            Left            =   360
            TabIndex        =   83
            Top             =   2520
            Width           =   765
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            Height          =   195
            Index           =   22
            Left            =   360
            TabIndex        =   58
            Top             =   495
            Width           =   450
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Index           =   25
            Left            =   360
            TabIndex        =   57
            Top             =   990
            Width           =   585
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto"
            Height          =   195
            Index           =   26
            Left            =   5220
            TabIndex        =   56
            Top             =   1080
            Width           =   450
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Cambio"
            Height          =   195
            Index           =   27
            Left            =   360
            TabIndex        =   55
            Top             =   3030
            Width           =   1110
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Index           =   31
            Left            =   360
            TabIndex        =   54
            Top             =   2100
            Width           =   585
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto"
            Height          =   195
            Index           =   30
            Left            =   5220
            TabIndex        =   53
            Top             =   2070
            Width           =   450
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto Total"
            Height          =   195
            Index           =   28
            Left            =   5220
            TabIndex        =   52
            Top             =   3030
            Width           =   855
         End
         Begin VB.Line Line1 
            BorderColor     =   &H8000000C&
            BorderStyle     =   6  'Inside Solid
            X1              =   2400
            X2              =   9330
            Y1              =   1680
            Y2              =   1650
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Detracci�n / Retenci�n"
            Height          =   195
            Index           =   29
            Left            =   360
            TabIndex        =   51
            Top             =   1560
            Width           =   1680
         End
         Begin VB.Label lblMontoTotal 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7080
            TabIndex        =   50
            Top             =   3030
            Width           =   2295
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Valor Cambio Sunat"
            Height          =   195
            Index           =   23
            Left            =   360
            TabIndex        =   49
            Top             =   3570
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Tipo Cambio"
            Height          =   195
            Index           =   32
            Left            =   5220
            TabIndex        =   48
            Top             =   2580
            Width           =   1380
         End
      End
      Begin TAMControls.ucBotonNavegacion ucBotonNavegacion1 
         Height          =   30
         Left            =   5550
         TabIndex        =   38
         Top             =   5220
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   53
      End
      Begin VB.Frame fraCompras 
         Caption         =   "Definici�n del Comprobante"
         Height          =   4395
         Index           =   1
         Left            =   -74640
         TabIndex        =   21
         Top             =   720
         Width           =   9765
         Begin VB.TextBox txtSerieComprobante 
            Height          =   315
            Left            =   7200
            MaxLength       =   4
            TabIndex        =   37
            Top             =   1380
            Width           =   615
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
            Left            =   9120
            TabIndex        =   11
            ToolTipText     =   "Buscar Proveedor"
            Top             =   2385
            Width           =   375
         End
         Begin VB.ComboBox cboMoneda 
            Height          =   315
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   1875
            Width           =   2295
         End
         Begin VB.TextBox txtDescripcion 
            Height          =   315
            Left            =   2520
            MaxLength       =   100
            TabIndex        =   13
            Top             =   3915
            Width           =   6975
         End
         Begin VB.ComboBox cboTipoComprobante 
            Height          =   315
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   885
            Width           =   6975
         End
         Begin VB.TextBox txtNumComprobante 
            Height          =   315
            Left            =   7920
            MaxLength       =   10
            TabIndex        =   9
            Top             =   1380
            Width           =   1575
         End
         Begin MSComCtl2.DTPicker dtpFechaRegistro 
            Height          =   315
            Left            =   7200
            TabIndex        =   6
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
            TabIndex        =   8
            Top             =   1380
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
         Begin VB.Label lblTipoDocID 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2520
            TabIndex        =   39
            Top             =   2880
            Width           =   1815
         End
         Begin VB.Label lblNumDocID 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   4410
            TabIndex        =   36
            Top             =   2880
            Width           =   2655
         End
         Begin VB.Label lblMontoGasto 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   4800
            TabIndex        =   34
            Top             =   7440
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Label lblCodProveedor 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   5790
            TabIndex        =   33
            Top             =   2010
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Label lblDireccion 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2520
            TabIndex        =   12
            Top             =   3390
            Width           =   6960
         End
         Begin VB.Label lblProveedor 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2520
            TabIndex        =   10
            Top             =   2385
            Width           =   6600
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Comprobante"
            Height          =   195
            Index           =   18
            Left            =   360
            TabIndex        =   32
            Top             =   1410
            Width           =   1440
         End
         Begin VB.Label lblNumSecuencial 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2520
            TabIndex        =   5
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            Height          =   195
            Index           =   17
            Left            =   5400
            TabIndex        =   31
            Top             =   375
            Width           =   450
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Documento ID"
            Height          =   195
            Index           =   13
            Left            =   360
            TabIndex        =   29
            Top             =   2880
            Width           =   1035
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Index           =   12
            Left            =   360
            TabIndex        =   28
            Top             =   1905
            Width           =   585
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Comprobante"
            Height          =   195
            Index           =   11
            Left            =   360
            TabIndex        =   27
            Top             =   870
            Width           =   1305
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Direcci�n"
            Height          =   195
            Index           =   8
            Left            =   360
            TabIndex        =   26
            Top             =   3405
            Width           =   675
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Descripci�n"
            Height          =   195
            Index           =   6
            Left            =   360
            TabIndex        =   25
            Top             =   3930
            Width           =   840
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor"
            Height          =   195
            Index           =   5
            Left            =   360
            TabIndex        =   24
            Top             =   2400
            Width           =   735
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Num. Registro"
            Height          =   195
            Index           =   4
            Left            =   360
            TabIndex        =   23
            Top             =   375
            Width           =   1005
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Num. Comprobante"
            Height          =   195
            Index           =   3
            Left            =   5400
            TabIndex        =   22
            Top             =   1410
            Width           =   1365
         End
      End
      Begin VB.Frame fraCompras 
         Caption         =   "Criterios de b�squeda"
         Height          =   1335
         Index           =   0
         Left            =   360
         TabIndex        =   17
         Top             =   720
         Width           =   9705
         Begin MSComCtl2.DTPicker dtpFechaDesde 
            Height          =   345
            Left            =   3600
            TabIndex        =   1
            Top             =   840
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   609
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
            Height          =   345
            Left            =   7200
            TabIndex        =   2
            Top             =   840
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   609
            _Version        =   393216
            Format          =   50397185
            CurrentDate     =   39042
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Registro"
            Height          =   195
            Index           =   15
            Left            =   840
            TabIndex        =   30
            Top             =   930
            Width           =   1305
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            Height          =   195
            Index           =   2
            Left            =   6000
            TabIndex        =   20
            Top             =   900
            Width           =   420
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            Height          =   195
            Index           =   1
            Left            =   2520
            TabIndex        =   19
            Top             =   900
            Width           =   465
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fondo"
            Height          =   195
            Index           =   0
            Left            =   840
            TabIndex        =   18
            Top             =   360
            Width           =   450
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gLista 
         Height          =   4785
         Left            =   360
         OleObjectBlob   =   "frmGenerarOrdenCobroProvision.frx":7FAB
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   2280
         Width           =   9690
      End
   End
End
Attribute VB_Name = "frmGenerarOrdenCobroProvision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
Dim strEstadoRegCompra          As String, strCodCuenta                 As String
Dim strCodComisionista  As String
Dim adoRegistro                 As ADODB.Recordset
Dim adoRegistroAux              As ADODB.Recordset
Dim numContadorGastos           As Integer
Dim dblPorcenDetraccion         As Double
Dim dblMontoSubtotal            As Double
Dim strNumOrdenPagoLista        As String
Dim strNumOrdenPago             As String
Dim adoField                    As ADODB.Field
Dim adoRegistroAuxGastos        As ADODB.Recordset
Dim rs                          As ADODB.Recordset


Private Sub Calculos()

    Dim intRegistro As Integer
    
    If Trim(txtSubTotal.Text) = Valor_Caracter Or Trim(txtIgv.Text) = Valor_Caracter Or Trim(txtTotal.Text) = Valor_Caracter Then Exit Sub
    
    Call cboTipoValorCambio_Click
    
    If strCodAfectacion = Codigo_Afecto Then
        If strIndImpuesto = Valor_Indicador Then
            'txtTotal.Text = lblMontoGasto.Caption
            'txtIgv.Text = CStr(CCur(txtTotal.Text) * gdblTasaIgv / (1 + gdblTasaIgv))
            txtSubTotal.Text = dblMontoSubtotal 'CStr(CCur(txtTotal.Text) - CCur(txtIgv.Text))
            txtIgv.Text = CStr(Round(dblMontoSubtotal * gdblTasaIgv, 2))
            txtTotal.Text = CDbl(txtIgv.Text) + CDbl(txtSubTotal.Text)
        ElseIf strIndRetencion = Valor_Indicador Then
            txtSubTotal.Text = lblMontoGasto.Caption
            If strCodMoneda <> Codigo_Moneda_Local Then
'                If (CCur(txtSubTotal.Text) * CDbl(txtTipoCambioPago.Text)) > gcurMontoMaximoRetencion Then
                txtIgv.Text = CStr(CCur(txtSubTotal.Text) * gdblTasaRetencion)
'                Else
'                    txtIgv.Text = "0"
'                End If
            Else
                If CCur(txtSubTotal.Text) > gcurMontoMaximoRetencion Then
                    txtIgv.Text = CStr(CCur(txtSubTotal.Text) * gdblTasaRetencion)
'                    cboDetraccion.Tag = "1"
'                    intRegistro = ObtenerItemLista(arrDetraccion(), Codigo_Respuesta_Si)
'                    txtIgv.Text = CStr(CCur(txtSubTotal.Text) * gdblTasaRetencion)
                Else
                    txtIgv.Text = "0"
'                    cboDetraccion.Tag = "1"
'                    intRegistro = ObtenerItemLista(arrDetraccion(), Codigo_Respuesta_No)
'                    If intRegistro >= 0 Then cboDetraccion.ListIndex = intRegistro
                End If
            End If
            txtTotal.Text = CStr(CCur(txtSubTotal.Text) - CCur(txtIgv.Text))
        Else
            txtSubTotal.Text = lblMontoGasto.Caption
            txtIgv.Text = "0"
            txtTotal.Text = txtSubTotal.Text
        End If
    Else
        If strIndImpuesto = Valor_Indicador Then
            txtTotal.Text = lblMontoGasto.Caption
            txtSubTotal.Text = txtTotal.Text
            txtIgv.Text = "0"
        ElseIf strIndRetencion = Valor_Indicador Then
            txtSubTotal.Text = lblMontoGasto.Caption
        
            'If strCodMoneda <> Codigo_Moneda_Local Then
                txtIgv.Text = CStr(CCur(txtSubTotal.Text) * gdblTasaRetencion / 100)
            'End If
            txtTotal.Text = CStr(CCur(txtSubTotal.Text) - CCur(txtIgv.Text))
        Else
            txtSubTotal.Text = lblMontoGasto.Caption
            txtTotal.Text = txtSubTotal.Text
        End If
        'txtIgv.Text = "0"
    End If
    
    If strDetraccionSiNo = Codigo_Respuesta_Si Then
        If strIndImpuesto = Valor_Indicador Then
            txtMontoDetraccion.Text = CStr(Round(CCur(txtTotal.Text) * dblPorcenDetraccion, 2))
            txtMontoUnico.Text = CStr(CCur(txtTotal.Text) - CCur(txtMontoDetraccion.Text))
            If strCodMoneda <> Codigo_Moneda_Local Then
                txtMontoDetraccion.Text = CStr(Round(txtMontoDetraccion.Text * CDbl(txtTipoCambioPago.Text), 2))
            End If
            
        ElseIf strIndRetencion = Valor_Indicador Then
            'If strCodMoneda <> Codigo_Moneda_Local Then
            '    txtMontoDetraccion.Text = CStr(CCur(txtSubTotal.Text) * gdblTasaRetencion / 100 * CDbl(txtTipoCambioPago.Text))
            'Else
                txtMontoDetraccion.Text = CStr(CCur(txtSubTotal.Text) * gdblTasaRetencion / 100)
            'End If
            txtMontoUnico.Text = CStr(CCur(txtSubTotal.Text) - CCur(txtMontoDetraccion.Text))
        End If
        lblMontoTotal.Caption = CStr(CCur(txtTotal.Text)) 'CStr(CCur(txtMontoUnico.Text) + CCur(txtMontoDetraccion.Text))
    Else
        txtMontoDetraccion.Text = "0"
        If strIndImpuesto = Valor_Indicador Then
            txtMontoUnico.Text = CStr(CCur(txtTotal.Text))
        ElseIf strIndRetencion = Valor_Indicador Then
            txtMontoUnico.Text = CStr(CCur(txtSubTotal.Text))
        Else
            txtMontoUnico.Text = CStr(CCur(txtSubTotal.Text))
        End If
        lblMontoTotal.Caption = CStr(CCur(txtMontoUnico.Text))
    End If
    
    
End Sub

Private Sub CalculosPago()

    If Trim(txtMontoUnico.Text) = Valor_Caracter Or Trim(txtMontoDetraccion.Text) = Valor_Caracter Then Exit Sub
                
    lblMontoTotal.Caption = CStr(CCur(txtMontoUnico.Text) + CCur(txtMontoDetraccion.Text))
    
End Sub

'''Private Sub CargarPendientes()
''
'''    strSQL = "SELECT FG.CodGasto,FG.CodDetalleGasto,NumGasto,DCG.CodAnalitica,CG.DescripConcepto,DCG.DescripGasto,MontoGasto " & _
'''        "FROM FondoGasto FG JOIN DetalleConceptoGasto DCG ON(DCG.CodDetalleGasto=FG.CodDetalleGasto AND DCG.CodGasto=FG.CodGasto) " & _
'''        "JOIN ConceptoGasto CG ON(CG.CodGasto=DCG.CodGasto) " & _
'''        "WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND IndConfirma='' AND IndVigente='X' " & _
'''        "ORDER BY DCG.DescripGasto"
''
''    strSQL = "SELECT FG.CodCuenta,NumGasto,CodFile,CodAnalitica,DescripCuenta,DescripGasto,CodTipoGasto,MontoGasto,MontoDevengo " & _
''        "FROM FondoGasto FG JOIN FondoConceptoGasto FCG ON(FCG.CodCuenta=FG.CodCuenta AND FCG.CodAdministradora=FG.CodAdministradora AND FCG.CodFondo=FG.CodFondo) " & _
''        "JOIN PlanContable PCG ON(PCG.CodCuenta=FCG.CodCuenta) " & _
''        "WHERE FG.CodFondo='" & strCodFondo & "' AND FG.CodAdministradora='" & gstrCodAdministradora & "' AND IndConfirma=''"
''
'''    strSQL = "SELECT FG.CodCuenta,NumGasto,CodFile,CodAnalitica,DescripCuenta,DescripGasto,MontoGasto,CodTipoGasto " & _
'''        "FROM FondoGasto FG JOIN FondoConceptoGasto FCG ON(FCG.CodCuenta=FG.CodCuenta AND FCG.CodAdministradora=FG.CodAdministradora AND FCG.CodFondo=FG.CodFondo) " & _
'''        "JOIN PlanContable PCG ON(PCG.CodCuenta=FCG.CodCuenta) " & _
'''        "WHERE FG.CodFondo='" & strCodFondo & "' AND FG.CodAdministradora='" & gstrCodAdministradora & "' AND IndConfirma='' AND (FG.IndVigente='X' OR FG.IndVigente='')"
''
''    strEstado = Reg_Defecto
''    With adoPendientes
''        .ConnectionString = gstrConnectConsulta
''        .RecordSource = strSQL
''        .Refresh
''    End With
''
''    tdgPendientes.Refresh

'''End Sub

Private Sub Deshabilita()

    strIndDetraccion = Valor_Caracter
'    cboFormaPagoDetraccion.Enabled = False
'    cboCuentaFondoDetraccion.Enabled = False
    
'    Call ColorControlDeshabilitado(cboFormaPagoDetraccion)
'    Call ColorControlDeshabilitado(cboCuentaFondoDetraccion)
    
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

    strCodAfectacion = Valor_Caracter
    If cboAfectacion.ListIndex < 0 Then Exit Sub
    
    strCodAfectacion = arrAfectacion(cboAfectacion.ListIndex)
    
    Call Calculos
    
End Sub


Private Sub cboCreditoFiscal_Click()

    strCodCreditoFiscal = Valor_Caracter
    If cboCreditoFiscal.ListIndex < 0 Then Exit Sub
    
    strCodCreditoFiscal = arrCreditoFiscal(cboCreditoFiscal.ListIndex)
    
    Call Calculos
    
End Sub





Private Sub cboDetraccion_Click()

    strDetraccionSiNo = Valor_Caracter
    If cboDetraccion.ListIndex < 0 Then Exit Sub
         
    strDetraccionSiNo = Trim(arrDetraccion(cboDetraccion.ListIndex))
    
'    If cboDetraccion.Tag = "1" Then Exit Sub
    
    Call Calculos
    
    'cboDetraccion.Tag = "0"
    
End Sub


Private Sub cboFondo_Click()

    Dim adoRegistro As ADODB.Recordset
    
    strCodFondo = Valor_Caracter
    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(arrFondo(cboFondo.ListIndex))
    
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Fecha Vigente, Valor Cuota, Valor Cuota T-1, Moneda y Cantidad Inicial de Cuotas del Fondo ***
        .CommandText = "{ call up_ACSelDatosParametro(23,'" & strCodFondo & "','" & gstrCodAdministradora & "','000') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            gdatFechaActual = adoRegistro("FechaCuota")
            gdblTipoCambio = adoRegistro("ValorTipoCambio")
            gstrCodMoneda = adoRegistro("CodMoneda")
            dtpFechaDesde.Value = gdatFechaActual
            dtpFechaHasta.Value = dtpFechaDesde.Value
            
            gstrFechaActual = Convertyyyymmdd(adoRegistro("FechaCuota"))
            gstrMesActual = Format(Month(gdatFechaActual), "00")
            gstrPeriodoActual = Format(Year(gdatFechaActual), "0000")
            
            gdblTipoCambio = ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, gdatFechaActual, Codigo_Moneda_Local, gstrCodMoneda)
            If gdblTipoCambio = 0 Then gdblTipoCambio = ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, DateAdd("d", -1, gdatFechaActual), Codigo_Moneda_Local, gstrCodMoneda)
                        
            frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
            
            'If strEstadoRegCompra <> Estado_Registro_Contabilizado Then
            '    Call CargarOrdenesPago
            'End If
            
'            Call Buscar
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
End Sub

Private Sub CargarOrdenesPago()
'*** Ordenes de pago del Fondo ***
strSQL = "SELECT op.NumOrdenPago CODIGO, (RTRIM(fg.DescripGasto)) + ' - ' + CONVERT(varchar(20),op.MontoOrdenPago) + ' ' + MO.CodSigno AS DESCRIP " & _
         "FROM OrdenPago op " & _
         "INNER JOIN FondoGasto fg ON (op.CodFondo = fg.CodFondo AND op.CodAdministradora = fg.CodAdministradora AND " & _
         "op.NumGasto = fg.NumGasto) " & _
         "JOIN Moneda MO ON (MO.CodMoneda = op.CodMoneda) " & _
         "WHERE op.CodFondo='" & strCodFondo & "' " & _
           "AND op.CodAdministradora='" & gstrCodAdministradora & "' " & _
           "AND fg.CodProveedor = '" & lblCodProveedor.Caption & "' " & _
           "AND op.CodMoneda = '" & strCodMoneda & "' " & _
           "AND op.Estado = '01'" & _
           "AND op.NumOrdenPago NOT IN (" & strNumOrdenPagoLista & ")" '(SELECT RTRIM(LTRIM(item)) FROM dbo.fnSplit('" & strNumOrdenPagoLista & "',','))"

        'ACD.CodCuenta                       IN  (SELECT RTRIM(LTRIM(item)) FROM dbo.fnSplit(@CodCuentaBusqueda,','))

'            strSQL = "SELECT (FCG.CodCuenta + CodAnalitica) CODIGO,(RTRIM(DescripCuenta)) DESCRIP " & _
'                "FROM FondoConceptoGasto FCG JOIN PlanContable PCG ON(PCG.CodCuenta=FCG.CodCuenta AND PCG.CodAdministradora=FCG.CodAdministradora) " & _
'                "JOIN FondoGasto FG ON(FG.CodCuenta=FCG.CodCuenta AND FG.CodAdministradora=FCG.CodAdministradora AND FG.CodFondo=FCG.CodFondo) " & _
'                "WHERE (FG.CodFile='099' OR FG.CodFile<>'098') AND FCG.CodFondo='" & strCodFondo & "' AND FCG.CodAdministradora='" & gstrCodAdministradora & "' " & _
'                "UNION " & _
'                "SELECT (CodCuenta + CodAnalitica) CODIGO,(RTRIM(DescripGasto)) DESCRIP " & _
'                "FROM FondoGasto " & _
'                "WHERE CodFile='098' AND CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
CargarControlLista strSQL, cboGasto, arrGasto(), Sel_Defecto
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
        Case vPrint
            Call Contabilizar
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
'    Call Buscar
    
End Sub

Public Sub Contabilizar()

'Dim strMsgError As String
'
'On Error GoTo err
'
''Validamos si el registro de compra ya fue enviado a comtabilidad
'If strEstado = Reg_Edicion Then
'
'    'strEstadoRegCompra = traerCampo("RegistroCompra", "Estado", "NumRegistro", gLista.Columns.ColumnByFieldName("NumRegistro").Value, " CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' ")
'
'    If strEstadoRegCompra = Estado_Registro_Contabilizado Then
'        strMsgError = "El Registro de Compras ya fue enviado a Contabilidad"
'        GoTo err
'    End If
'
'    If MsgBox("�Seguro de contabilizar el Registro de Compras?", vbQuestion + vbYesNo, App.Title) = vbYes Then
'        '*** Generar Orden si no est� generada o actualizar ***
'        Call ContabilizarRegistroCompra(CInt(gLista.Columns.ColumnByFieldName("NumRegistro").Value), strCodFondo, Trim(lblCodProveedor.Caption), strMsgError)
'        If strMsgError <> "" Then GoTo err
'        MsgBox "Registro de Compras contabilizado con exito", vbInformation, App.Title
'    End If
'
'    Call Cancelar
'Else
'    MsgBox "Grabe los datos del Registro de Compras antes de Contabilizarlo!", vbInformation, App.Title
'End If
'
'Exit Sub
'
'err:
'If strMsgError = "" Then strMsgError = err.Description
'MsgBox strMsgError, vbInformation, App.Title
End Sub

Public Sub Grabar()
    
    Dim adoRegistro         As ADODB.Recordset
    Dim adoAuxiliar         As ADODB.Recordset
    Dim strNumCaja          As String
    Dim strCodDetalleFile   As String, strCodMonedaGasto        As String
    Dim strDescripGasto     As String, strSQLOrdenCajaDetalleI  As String
    Dim strSQLOrdenCaja     As String, strSQLOrdenCajaDetalle   As String
    Dim strSQLOrdenCajaMN   As String, strSQLOrdenCajaDetalleMN As String
    Dim strFechaAnterior    As String, strFechaSiguiente        As String
    Dim curSaldoProvision   As Currency, intCantRegistros       As Integer
    Dim dblTipCambio        As Double, strNuevoMod              As String
    Dim datFechaFinPeriodo  As Date
    
    Dim xmlDocGastos As DOMDocument60 'JCB
    Dim strMsgError As String 'JCB
    
    If strEstado = Reg_Consulta Then Exit Sub
    
    If Not TodoOK() Then Exit Sub
    
'    XMLDetalleGrid xmlDocGastos, "DetGastos", gGastos, "Item,DescripGasto,NumOrdenPago,CodMoneda,MontoSubTotal,MontoImpuesto,TasaImpuesto,MontoGasto,CodFile,CodDetalleFile,CodAnalitica,NumGasto", strMsgError 'JCB
'''    If strMsgError <> "" Then GoTo CtrlError 'JCB
    
    Dim objElem As MSXML2.IXMLDOMElement
    Dim objParent As MSXML2.IXMLDOMElement
    Dim lngPos As Long, lngParent As Long
    Dim i As Integer, j As Integer, aux As Integer
    Dim lblnSuccess As Boolean
    Dim NomCampos() As String, ArrayCols() As String
    Dim indCumpleCondicion As Boolean
    Dim strNomCampos As String
    Dim objXML As DOMDocument60
    Dim strNomEntidad As String
    strNomEntidad = "DetGastos"
    strNomCampos = "DescripParticipe,CodAnalitica,MontoMovimiento"
   
    
    NomCampos = Split(strNomCampos, ",")
    'ArrayCols = Split(strNumColumnas, ",")
    ' iniciando el documento xml
    If objXML Is Nothing Then
        Set objXML = New MSXML2.DOMDocument60
        Set objXML.documentElement = objXML.createElement("ROOT")
    'Else
    '    lblnSuccess = objXML.loadXML(xmlDoc)
    End If
    Set objParent = objXML.documentElement
    'Recorriendo todas las filas de una rejilla
'    If tdgGastos.Count > 0 Then
 
                
        tdgGastos.MoveFirst
        Do While Not tdgGastos.EOF
            indCumpleCondicion = True
'            If strCampoCond <> "" Then
'                    If g.Columns.ColumnByFieldName(strCampoCond).Value <> strDatoCond Then indCumpleCondicion = False
'            End If
            
            If indCumpleCondicion Then
                Set objElem = objParent.appendChild(objXML.createElement(strNomEntidad))
                For j = 0 To UBound(NomCampos)
                    'a�adiendo los atributos, solo para las columnas especificadas
                    objElem.setAttribute NomCampos(j), "" & Trim(tdgGastos.Columns(j + 1).Value)
                Next
            End If
            
            tdgGastos.MoveNext
        Loop
'        g.Dataset.EnableControls
'    End If
    
    Set objElem = Nothing
    Set objParent = Nothing
    strCodFile = "000"
    strCodAnalitica = "00000000"
    
'''        strCodFile = Trim(tdgPendientes.Columns(9).Value) JCB de donde saco este dato

'''        If strCodCreditoFiscal = Codigo_Tipo_Credito_RentaGravada Then
'''            If Not ExisteDinamica(strCodFile, strCodDetalleGasto, gstrCodAdministradora, Codigo_Dinamica_Impuesto) Then Exit Sub
'''        End If
        
        Me.MousePointer = vbHourglass
        
        strNuevoMod = "I"
        If strEstado = Reg_Edicion Then strNuevoMod = "U"
        
        '*** Guardar ***
        
        
        
        With adoComm
            If strCodComisionista <> "00000008" Then
            '*** Adicionar registro ***
                .CommandText = "{ call up_CNProcGeneraOrdenCobroProveedores('" & _
                    strCodFondo & "','" & gstrCodAdministradora & "','" & Convertyyyymmdd(dtpFechaDesde.Value) & "'," & _
                    "'" & Convertyyyymmdd(dtpFechaHasta.Value) & "','" & gstrCodMoneda & "'," & _
                     "'" & strCodComisionista & "','" & objXML.xml & "') }"
                adoConn.Execute .CommandText
            Else
                .CommandText = "{ call up_CNProcGeneraOrdenCobroSafi('" & _
                    strCodFondo & "','" & gstrCodAdministradora & "','" & Convertyyyymmdd(dtpFechaDesde.Value) & "'," & _
                    "'" & Convertyyyymmdd(dtpFechaHasta.Value) & "','" & gstrCodMoneda & "'," & _
                     "'" & strCodComisionista & "','" & objXML.xml & "') }"
                adoConn.Execute .CommandText
            End If
        End With
                                    
        Me.MousePointer = vbDefault
                    
        If strNuevoMod = "I" Then
            MsgBox Mensaje_Adicion_Exitosa, vbExclamation
        Else
            MsgBox Mensaje_Edicion_Exitosa, vbExclamation
        End If
        
        frmMainMdi.stbMdi.Panels(3).Text = "Acci�n"
        
        cmdOpcion.Visible = True
        With tabRegistroCompras
            .TabEnabled(0) = True
            .Tab = 0
        End With

'        Call Buscar
'''    End If
    
'''    If strEstado = Reg_Edicion Then
'''        Me.MousePointer = vbHourglass
                    
'''        '*** Guardar ***
'''        With adoComm
'''            '*** Actualizar registro ***
'''            .CommandText = "{ call up_CNManRegistroCompra('" & _
'''                strCodFondo & "','" & gstrCodAdministradora & "','" & Convertyyyymmdd(dtpFechaRegistro.Value) & "'," & _
'''                CInt(lblNumSecuencial.Caption) & ",'" & strCodTipoComprobante & "','" & Convertyyyymmdd(dtpFechaComprobante.Value) & "','" & _
'''                Trim(txtSerieComprobante.Text) & "-" & Trim(txtNumComprobante.Text) & "','" & strCodGasto & "','" & Trim(lblCodProveedor.Caption) & "','" & _
'''                Trim(txtDescripcion.Text) & "','" & strCodAfectacion & "','" & strCodCreditoFiscal & "','" & Trim(txtPeriodoFiscal.Text) & "','" & _
'''                strCodMoneda & "'," & CDec(txtSubTotal.Text) & "," & CDec(txtIgv.Text) & "," & CDec(txtTotal.Text) & ",'" & strDetraccionSiNo & "','" & _
'''                strCodFile & "','" & strCodAnalitica & "','" & Convertyyyymmdd(dtpFechaPago.Value) & "','" & strCodFormaPagoUnico & "','" & _
'''                strCodMonedaUnico & "','" & strCodFileUnico & "','" & strCodAnaliticaUnico & "'," & _
'''                CDec(txtMontoUnico.Text) & ",'" & strCodFormaPagoDetraccion & "','" & strCodMonedaDetraccion & "','" & _
'''                strCodFileDetraccion & "','" & strCodAnaliticaDetraccion & "'," & CDec(txtMontoDetraccion.Text) & ",'" & _
'''                strCodValorTipoCambio & "'," & CDec(txtTipoCambioPago.Text) & "," & CDec(lblMontoTotal.Caption) & ",'" & Convertyyyymmdd(Valor_Fecha) & "','" & _
'''                0 & "','" & Estado_Activo & "','" & CrearXMLDetalle(xmlDocGastos) & "','U') }"
'''            adoConn.Execute .CommandText
            
            'JCB FechaConfirma, de donde saco este dato?
'''            .CommandText = "SELECT FechaConfirma FROM FondoGasto " & _
'''                "WHERE CodCuenta='" & strCodGasto & "' AND " & _
'''                "NumGasto=" & CInt(tdgConsulta.Columns(7).Value) & " AND CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND IndVigente=''"
'''            Set adoRegistro = .Execute
'''
'''            If Not adoRegistro.EOF Then
'''                If adoRegistro("FechaConfirma") <> Valor_Fecha Then
'''                    MsgBox "Orden de cobro ya fu� generada", vbCritical, Me.Caption
'''                    adoRegistro.Close: Set adoRegistro = Nothing
'''                    Me.MousePointer = vbDefault
'''                    Exit Sub
'''                Else
'''                    Call GenerarOrdenGastosFondo(strCodDetalleGasto, CInt(lblNumSecuencial.Caption), strCodFondo, CInt(lblNumSecuencial.Caption), Trim(lblCodProveedor.Caption))
                    
'                    '*** Actualizar Registro del Gasto ***
'                    .CommandText = "UPDATE FondoGasto SET IndConfirma='X' " & _
'                        "WHERE CodCuenta='" & strCodGasto & "' AND " & _
'                        "NumGasto=" & CInt(tdgConsulta.Columns(7).Value) & " AND CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
'                    adoConn.Execute .CommandText
'''                End If
'''            Else
'''                adoConn.Execute .CommandText
'''            End If
'''            adoRegistro.Close: Set adoRegistro = Nothing
                                            
'''        End With
'''
'''        Me.MousePointer = vbDefault
'''
'''        MsgBox Mensaje_Edicion_Exitosa, vbExclamation
'''
'''        frmMainMdi.stbMdi.Panels(3).Text = "Acci�n"
'''
'''        cmdOpcion.Visible = True
'''        With tabRegistroCompras
'''            .TabEnabled(0) = True
'''            .Tab = 0
'''        End With
'''
'''        Call CargarOrdenesPago
'''        Call Buscar
'''    End If
    
End Sub

Private Function TodoOK() As Boolean
        
    TodoOK = False
            
'    If cboTipoComprobante.ListIndex <= 0 Then
'        MsgBox "Seleccione el tipo de comprobante", vbCritical, Me.Caption
'        If cboTipoComprobante.Enabled Then cboTipoComprobante.SetFocus
'        Exit Function
'    End If

     
        
'    If Trim(txtSerieComprobante.Text) = Valor_Caracter Then
'        MsgBox "Ingrese el n�mero de serie, si no lo tiene ingrese cero", vbCritical, Me.Caption
'        If txtSerieComprobante.Enabled Then txtSerieComprobante.SetFocus
'        Exit Function
'    End If
'
'    If Trim(txtNumComprobante.Text) = Valor_Caracter Then
'        MsgBox "Ingrese el n�mero de comprobante", vbCritical, Me.Caption
'        If txtNumComprobante.Enabled Then txtNumComprobante.SetFocus
'        Exit Function
'    End If
'
'    If cboGasto.ListIndex <= 0 Then
'        MsgBox "Seleccione el gasto relacionado", vbCritical, Me.Caption
'        If cboGasto.Enabled Then cboGasto.SetFocus
'        Exit Function
'    End If

    
'    If Trim(lblProveedor.Caption) = Valor_Caracter Then
'        MsgBox "Seleccione el Proveedor", vbCritical, Me.Caption
'        If cmdProveedor.Enabled Then cmdProveedor.SetFocus
'        Exit Function
'    End If
'
'    If strDetraccionSiNo = Codigo_Respuesta_Si And strCodMoneda <> Codigo_Moneda_Local Then
'        If CDbl(txtTipoCambioPago.Text) = 0 Then
'            MsgBox "Tipo de Cambio SUNAT NO REGISTRADO...", vbCritical, Me.Caption
'            If cboTipoValorCambio.Enabled Then cboTipoValorCambio.SetFocus
'            Exit Function
'        End If
'    End If
    
    If adoRegistroAuxGastos.RecordCount = 0 Then
        MsgBox "", vbCritical, Me.Caption
        Exit Function
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
    
    Dim strEstadoRegCompra As String
    
    If strEstado = Reg_Edicion Then
        strEstadoRegCompra = traerCampo("RegistroCompra", "Estado", "NumRegistro", gLista.Columns.ColumnByFieldName("NumRegistro").Value, " CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' ")

        If strEstadoRegCompra = "04" Then
            MsgBox "El Registro de Compras ya fue enviado a Contabilidad, no se puede modificar", vbInformation, App.Title
            Exit Function
        End If
    End If
    
    'JAFR 10/12/2010:
    
'    If numContadorGastos = 0 Then
'        MsgBox "No hay �rdenes de pago asociadas al comprobante de pago", vbCritical, Me.Caption
'        Exit Function
'    End If
'
    If cboCreditoFiscal.ListIndex = 0 Then
        MsgBox "Seleccione el cr�dito fiscal", vbCritical, Me.Caption
        Exit Function
    End If
    
    'fin JAFR
    
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
        Case 1, 2
            If Index = 1 Then gstrNameRepo = "RegistroComprasParte1_1"
            If Index = 2 Then gstrNameRepo = "RegistroComprasParte2"
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

'Validar que no este contabilizado el "RegistroCompra"  --Estado = '04'

'Anular el movimiento de "RegistroCompra"               --Estado = '03'

'Activar las "OrdenPago" asociadas                      --Estado = '01'


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
    Dim strMsgError     As String
    
    
    tdgGastos.DataSource = Nothing
    txtSubTotal.Text = CStr(0)
    txtIgv.Text = CStr(0)
    txtTotal.Text = CStr(0)
    
    Select Case strModo
        Case Reg_Adicion
            fraCompras(1).Caption = "Definici�n del Registro - Fondo : " & Trim(cboFondo.Text)
            fraCompras(2).Caption = "Definici�n de Obligaci�n - Fondo : " & Trim(cboFondo.Text)
            
            dblMontoSubtotal = 0
            
            strNumOrdenPagoLista = "''"
            Call ConfiguraRecordsetAuxiliarGastos
            
'            Set adoRegistro = New ADODB.Recordset
'            With adoComm
'
'                .CommandText = "SELECT * FROM InstitucionPersona WHERE TipoPersona = '10' "
'                Set adoRegistro = .Execute
                

'                .CommandText = "SELECT PCO.DescripParticipe, PS.FechaSolicitud,PCS.CodFile,PCS.CodAnalitica,PCS.CodMoneda, PCS.SaldoFinalME * -1 as MontoMovimiento" & _
'                               "SUM(PCS.SaldoFinalME) * -1 AS MontoMovimiento FROM PartidaContableSaldos PCS JOIN ParticipeComisionista PC ON " & _
'                               "(PC.CodFondo = PCS.CodFondo AND PC.CodFile = PCS.CodFile AND PC.CodAnalitica = PCS.CodAnalitica)" & _
'                               "JOIN ParticipeSolicitud PS ON (PS.CodFondo = PC.CodFondo AND PS.NumSolicitud = PC.NumSolicitud)JOIN ParticipeContrato PCO" & _
'                               "ON (PS.CodParticipe = PCO.CodParticipe) Where PCS.CodFondo = '" & strCodFondo & "' AND PCS.CodFile = '098' and " & _
'                               "PC.CodComisionista = '00000002' and PCS.FechaSaldo = @FechaFin and PCS.CodCuenta = '421122' "
'                Set adoRegistro = .Execute
'
'                If Not adoRegistro.EOF Then
'                    If IsNull(adoRegistro("NumRegistro")) Then
'                        lblNumSecuencial.Caption = "1"
'                    Else
'                        lblNumSecuencial.Caption = CStr(adoRegistro("NumRegistro") + 1)
'                    End If
'                Else
'                    lblNumSecuencial.Caption = "1"
'                End If
'                adoRegistro.Close
'
'                strEstadoRegCompra = Estado_Registro_Ingresado
'
'                dtpFechaRegistro.Value = gdatFechaActual
'                dtpFechaComprobante.Value = gdatFechaActual
'                txtSerieComprobante.Text = Valor_Caracter
'                txtNumComprobante.Text = Valor_Caracter
'                If cboGasto.ListCount > 0 Then cboGasto.ListIndex = 0
'                lblCodProveedor.Caption = Valor_Caracter
'                lblProveedor.Caption = Valor_Caracter
'                lblDireccion.Caption = Valor_Caracter
'                txtDescripcion.Text = Valor_Caracter
'                lblAnalitica.Caption = Valor_Caracter
'                cboAfectacion.Enabled = True
'
'                cboAfectacion.Enabled = True
'                intRegistro = ObtenerItemLista(arrAfectacion(), Codigo_Afecto)
'                If intRegistro >= 0 Then cboAfectacion.ListIndex = intRegistro
'
'                If cboCreditoFiscal.ListCount > 0 Then cboCreditoFiscal.ListIndex = 0
'
'                txtPeriodoFiscal.Text = Valor_Caracter
'                txtSubTotal.Text = "0": txtIgv.Text = "0"
'                txtTotal.Text = "0"
'
'                dtpFechaPago.Value = gdatFechaActual
'                dtpFechaTipoCambioPago.Value = gdatFechaActual
'
'                If cboMonedaUnico.ListCount > 0 Then cboMonedaUnico.ListIndex = 0
'
'                intRegistro = ObtenerItemLista(arrMonedaUnico(), gstrCodMoneda)
'                If intRegistro >= 0 Then cboMonedaUnico.ListIndex = intRegistro
'
'                txtMontoUnico.Text = "0"
'
'                intRegistro = ObtenerItemLista(arrMoneda(), gstrCodMoneda)
'                If intRegistro >= 0 Then cboMoneda.ListIndex = intRegistro
'
'                If cboTipoComprobante.ListCount > 0 Then cboTipoComprobante.ListIndex = 0
'
'                If cboTipoValorCambio.ListCount > 0 Then cboTipoValorCambio.ListIndex = 1
'
'                If cboMonedaDetraccion.ListCount > 0 Then cboMonedaDetraccion.ListIndex = 0
'
'                If cboPorcenDetraccion.ListCount > 0 Then cboPorcenDetraccion.ListIndex = 4
'
'                intRegistro = ObtenerItemLista(arrMonedaDetraccion(), Codigo_Moneda_Local)
'                If intRegistro >= 0 Then cboMonedaDetraccion.ListIndex = intRegistro
'                cboMonedaDetraccion.Enabled = False
'
'                If cboDetraccion.ListCount > 0 Then cboDetraccion.ListIndex = 0
'                intRegistro = ObtenerItemLista(arrDetraccion(), Codigo_Respuesta_No)
'                If intRegistro >= 0 Then cboDetraccion.ListIndex = intRegistro
'
'                txtMontoDetraccion.Text = "0"
'                txtTipoCambioPago.Text = gdblTipoCambio
'
'                lblMontoTotal.Caption = "0"
'
'                strCodFile = "000"
'
'                Me.Refresh
'
'            End With
'
'            cboTipoComprobante.SetFocus
'
'        Case Reg_Edicion
'
'
'            Set adoRegistro = New ADODB.Recordset
'
'            With adoComm
'                .CommandText = "SELECT * FROM RegistroCompra " & _
'                    "WHERE NumRegistro=" & gLista.Columns.ColumnByFieldName("NumRegistro").Value & " AND CodFondo='" & _
'                    strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
'                Set adoRegistro = .Execute
'
'                If Not adoRegistro.EOF Then
'                    fraCompras(1).Caption = "Definici�n del Registro - Fondo : " & Trim(cboFondo.Text)
'                    fraCompras(2).Caption = "Definici�n de Obligaci�n - Fondo : " & Trim(cboFondo.Text)
'
'                    numContadorGastos = 0
'
'                    'Carga la lista de ordenes de pago
'                    strNumOrdenPagoLista = "''"
'
'                    Set adoAuxiliar = New ADODB.Recordset
'
'                    .CommandText = "SELECT NumOrdenPago FROM RegistroCompraDetalle " & _
'                                   "WHERE NumRegistro=" & gLista.Columns.ColumnByFieldName("NumRegistro").Value & " " & _
'                                     "AND CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
'                    Set adoAuxiliar = .Execute
'
'                    If Not adoAuxiliar.EOF Then
'                        Do Until adoAuxiliar.EOF
'                            If Trim(strNumOrdenPagoLista) = "''" Then
'                                strNumOrdenPagoLista = "'" & adoAuxiliar("NumOrdenPago") & "'"  'adoRegistro("NumGasto")
'                            Else
'                                strNumOrdenPagoLista = strNumOrdenPagoLista & ",'" & adoAuxiliar("NumOrdenPago") & "'"  'adoRegistro("NumGasto")
'                            End If
'                            adoAuxiliar.MoveNext
'                        Loop
'                    End If
'                    adoAuxiliar.Close: Set adoAuxiliar = Nothing
'
'                    strEstadoRegCompra = adoRegistro("Estado")
'
'                    lblNumSecuencial.Caption = CStr(adoRegistro("NumRegistro"))
'                    dtpFechaRegistro.Value = adoRegistro("FechaRegistro")
'
'                    dtpFechaComprobante.Value = adoRegistro("FechaComprobante")
'
'                    intRegistro = InStr(1, adoRegistro("NumComprobante"), "-")
'                    If intRegistro > 0 Then txtSerieComprobante.Text = Left(adoRegistro("NumComprobante"), intRegistro - 1)
'                    txtNumComprobante.Text = Mid(adoRegistro("NumComprobante"), intRegistro + 1)
'
'                    strIndDetraccion = Valor_Caracter
'                    If CCur(adoRegistro("MontoDetraccion")) > 0 Then strIndDetraccion = Valor_Indicador
'
''                    intRegistro = ObtenerItemLista(arrGasto(), adoRegistro("NumOrdenPago"))
''                    If intRegistro >= 0 Then cboGasto.ListIndex = intRegistro
'
'                    lblProveedor.Caption = Valor_Caracter
'                    lblDireccion.Caption = Valor_Caracter
'                    lblCodProveedor.Caption = adoRegistro("CodProveedor")
'
'                    Set adoAuxiliar = New ADODB.Recordset
'                    .CommandText = "SELECT IP.NumIdentidad, IP.DescripPersona, IP.Direccion1 + IP.Direccion2 Direccion, AP.DescripParametro TipoIdentidad " & _
'                        "FROM InstitucionPersona IP " & _
'                        "JOIN AuxiliarParametro AP ON (AP.CodParametro = IP.TipoIdentidad AND AP.CodTipoParametro = 'TIPIDE')" & _
'                        "WHERE CodPersona='" & lblCodProveedor.Caption & "' AND TipoPersona='" & Codigo_Tipo_Persona_Proveedor & "'"
'                    Set adoAuxiliar = .Execute
'
'                    If Not adoAuxiliar.EOF Then
'                        lblTipoDocID.Caption = Trim(adoAuxiliar("TipoIdentidad"))
'                        lblProveedor.Caption = Trim(adoAuxiliar("DescripPersona"))
'                        lblNumDocID.Caption = Trim(adoAuxiliar("NumIdentidad"))
'                        lblDireccion.Caption = Trim(adoAuxiliar("Direccion"))
'                    End If
'                    adoAuxiliar.Close: Set adoAuxiliar = Nothing
'
'                    txtDescripcion.Text = Trim(adoRegistro("DescripRegistro"))
'
'                    cboAfectacion.Enabled = True
'                    intRegistro = ObtenerItemLista(arrAfectacion(), adoRegistro("CodAfectacion"))
'                    If intRegistro >= 0 Then cboAfectacion.ListIndex = intRegistro
'
'                    intRegistro = ObtenerItemLista(arrCreditoFiscal(), adoRegistro("CodCreditoFiscal"))
'                    If intRegistro >= 0 Then cboCreditoFiscal.ListIndex = intRegistro
'
'                    intRegistro = ObtenerItemLista(arrMoneda(), adoRegistro("CodMoneda"))
'                    If intRegistro >= 0 Then cboMoneda.ListIndex = intRegistro
'
'                    dtpFechaPago.Value = adoRegistro("FechaPago")
'
'                    intRegistro = ObtenerItemLista(arrDetraccion(), adoRegistro("CodDetraccionSiNo"))
'                    If intRegistro >= 0 Then cboDetraccion.ListIndex = intRegistro
'
'                    If strDetraccionSiNo = Codigo_Respuesta_Si Then
'                        dtpFechaTipoCambioPago.Value = adoRegistro("FechaComprobante")
'                    Else
'                        dtpFechaTipoCambioPago.Value = adoRegistro("FechaPago")
'                    End If
'
'
'                    intRegistro = ObtenerItemLista(arrMonedaUnico(), adoRegistro("CodMonedaPago"))
'                    If intRegistro >= 0 Then cboMonedaUnico.ListIndex = intRegistro
'
'                    txtMontoUnico.Text = CStr(adoRegistro("MontoPago"))
'
'                    lblMontoGasto.Caption = adoRegistro("Importe")
'
'
'                    intRegistro = ObtenerItemLista(arrMonedaDetraccion(), adoRegistro("CodMonedaDetraccion"))
'                    If intRegistro >= 0 Then cboMonedaDetraccion.ListIndex = intRegistro
'                    cboMonedaDetraccion.Enabled = False
'
'                    txtMontoDetraccion.Text = CStr(adoRegistro("MontoDetraccion"))
'
'                    intRegistro = ObtenerItemLista(arrTipoValorCambio(), adoRegistro("ClaseTipoCambio"))
'                    If intRegistro >= 0 Then cboTipoValorCambio.ListIndex = intRegistro
'
'                    txtTipoCambioPago.Text = CStr(adoRegistro("TipoCambioPago"))
'
'                    lblMontoTotal.Caption = CStr(adoRegistro("MontoTotal"))
'                    strCodFile = adoRegistro("CodFile") 'Trim(tdgPendientes.Columns(9).Value)
'
'                    intRegistro = ObtenerItemLista(arrTipoComprobante(), adoRegistro("CodTipoComprobante"))
'                    If intRegistro >= 0 Then cboTipoComprobante.ListIndex = intRegistro
'
'                    cboTipoComprobante.SetFocus
'
'                    txtPeriodoFiscal.Text = adoRegistro("DescripPeriodoCredito")
'                    txtSubTotal.Text = CStr(adoRegistro("Importe"))
'                    txtIgv.Text = CStr(adoRegistro("ValorImpuesto"))
'                    txtTotal.Text = CStr(adoRegistro("ValorTotal"))
'
'                    'Si el registro de compras no ha sido contabilizado aun...
''                    If adoRegistro("Estado") <> "04" Then
''                        Call CargarOrdenesPago 'carga las ordenes de pago pendientes
''                    End If
'
'                    .CommandText = "SELECT ISNULL(COUNT(*),0) AS NumReg FROM RegistroCompraDetalle " & _
'                    "WHERE NumRegistro=" & gLista.Columns.ColumnByFieldName("NumRegistro").Value & " AND CodFondo='" & _
'                    strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
'                    Set adoAuxiliar = .Execute
'
'                    If Not adoAuxiliar.EOF Then
'                        numContadorGastos = adoAuxiliar("NumReg")
'                    Else
'                        numContadorGastos = 0
'                    End If
'
'                    adoAuxiliar.Close
'
'                    Set adoAuxiliar = Nothing
'
'                    'Muestro el detalle de las series
''                    Set adoRegistro = DataProcedimiento("up_GNSelFondo", strMsgError, 1, strCodFondo, gstrCodAdministradora)
'                    .CommandText = "SELECT SecRegistroDetalle AS Item, NumOrdenPago, CodFile, CodDetalleFile, CodAnalitica,DescripRegistroDetalle AS DescripGasto, " & _
'                                          "CodMoneda, MontoSubtotal, MontoImpuesto, TasaImpuesto, NumGasto, MontoTotal AS MontoGasto " & _
'                                   "FROM RegistroCompraDetalle " & _
'                                   "WHERE NumRegistro=" & gLista.Columns.ColumnByFieldName("NumRegistro").Value & " " & _
'                                     "AND CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
'
'                    Set adoRegistro = .Execute
'
'                    FormatoGrillaGastos strMsgError
'                    mostrarDatosGridRS gGastos, adoRegistro, strMsgError
'
'
'                End If
'                adoRegistro.Close: Set adoRegistro = Nothing
'            End With
    End Select
    
End Sub

Public Sub Adicionar()
Dim strMsgError As String

On Error GoTo err

    frmMainMdi.stbMdi.Panels(3).Text = "Adicionar Registro..."
    
'    FormatoGrillaGastos strMsgError
'    If strMsgError <> "" Then GoTo err
                
'    If adoPendientes.Recordset.RecordCount > 0 Then
'        tdgPendientes.SetFocus
'    Else
'        MsgBox "No existen gastos pendientes", vbCritical, Me.Caption
'        tdgConsulta.SetFocus
'        Exit Sub
'    End If
    
    strEstado = Reg_Adicion
    LlenarFormulario strEstado
    cmdOpcion.Visible = False
    With tabRegistroCompras
        .TabEnabled(0) = False
        .TabEnabled(1) = False
        .TabEnabled(3) = False
        .Tab = 2
    End With
    
Exit Sub
err:
If strMsgError = "" Then strMsgError = err.Description
MsgBox strMsgError, vbInformation, App.Title
End Sub

Private Sub cboGasto_Click()

    Dim adoRegistro         As ADODB.Recordset
    Dim curDiferencia       As Currency
    Dim intRegistro         As Integer
    Dim strSQL              As String
    Dim dblSubTotal         As Double
    Dim dblIGV              As Double
    Dim dblTotal            As Double
    Dim i As Integer
    strNumOrdenPago = Valor_Caracter
    strCodComisionista = Valor_Caracter
    strCodDetalleGasto = Valor_Caracter
    Set rs = New ADODB.Recordset
    If cboGasto.ListIndex < 0 Then Exit Sub
    
    strCodComisionista = arrGasto(cboGasto.ListIndex)
    
    
'                Set adoRegistro = New ADODB.Recordset
'            With adoComm
                
              If strCodComisionista <> "00000008" Then
                
                If gstrCodMoneda <> Codigo_Moneda_Local Then
                                     
                    strSQL = "SELECT '' AS Item,PCO.DescripParticipe, PS.FechaSolicitud,PCS.CodFile,PCS.CodAnalitica,PCS.CodMoneda, PCS.SaldoFinalME * -1 as MontoMovimiento " & _
                            "FROM PartidaContableSaldos PCS JOIN ParticipeComisionista PC ON " & _
                            "(PC.CodFondo = PCS.CodFondo AND PC.CodFile = PCS.CodFile AND PC.CodAnalitica = PCS.CodAnalitica) " & _
                            "JOIN ParticipeSolicitud PS ON (PS.CodFondo = PC.CodFondo AND PS.NumSolicitud = PC.NumSolicitud)JOIN ParticipeContrato PCO " & _
                            "ON (PS.CodParticipe = PCO.CodParticipe) Where PCS.CodFondo = '" & strCodFondo & "' AND PCS.CodFile = '098' and " & _
                            "PC.CodComisionista = '" & strCodComisionista & "' and PCS.FechaSaldo = '" & Convertyyyymmdd(DateAdd("d", 0, dtpFechaHasta.Value)) & "' and PCS.CodCuenta = '421122' ORDER BY PC.CodAnalitica "
'                    Set adoRegistro = .Execute
                Else
                    strSQL = "SELECT '' AS Item,PCO.DescripParticipe, PS.FechaSolicitud,PCS.CodFile,PCS.CodAnalitica,PCS.CodMoneda, PCS.SaldoFinalMN * -1 as MontoMovimiento " & _
                             "FROM PartidaContableSaldos PCS JOIN ParticipeComisionista PC ON " & _
                             "(PC.CodFondo = PCS.CodFondo AND PC.CodFile = PCS.CodFile AND PC.CodAnalitica = PCS.CodAnalitica) " & _
                             "JOIN ParticipeSolicitud PS ON (PS.CodFondo = PC.CodFondo AND PS.NumSolicitud = PC.NumSolicitud)JOIN ParticipeContrato PCO " & _
                             "ON (PS.CodParticipe = PCO.CodParticipe) Where PCS.CodFondo = '" & strCodFondo & "' AND PCS.CodFile = '098' and " & _
                             "PC.CodComisionista = '" & strCodComisionista & "' and PCS.FechaSaldo = '" & Convertyyyymmdd(DateAdd("d", 0, dtpFechaHasta.Value)) & "' and PCS.CodCuenta = '421121' ORDER BY PC.CodAnalitica "
'                    Set adoRegistro = .Execute
                End If
'            End With

        Call ConfiguraRecordsetAuxiliarGastos
            
        With rs
            .ActiveConnection = gstrConnectConsulta
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Open strSQL
            If .RecordCount > 0 Then
                    .MoveFirst
                       i = 1
            Do While Not .EOF
                 
                    adoRegistroAuxGastos.AddNew
                    For Each adoField In adoRegistroAuxGastos.Fields
                            If (adoField.Name = "Item") Then
                                adoRegistroAuxGastos.Fields(adoField.Name) = i
                            Else
                                adoRegistroAuxGastos.Fields(adoField.Name) = rs.Fields(adoField.Name)
                            End If
                            
                    Next
                    dblSubTotal = dblSubTotal + rs.Fields("MontoMovimiento")
                    dblIGV = Round((dblSubTotal * gdblTasaIgv), 2)
                    dblTotal = dblSubTotal + dblIGV
                    
                    i = i + 1
                    adoRegistroAuxGastos.Update
                    rs.MoveNext
                  
            Loop
            adoRegistroAuxGastos.MoveFirst
            End If
        End With
                
        Else
        
        
'          If gstrCodMoneda <> Codigo_Moneda_Local Then
                                     
                    strSQL = "SELECT '' AS Item,'BLANCO SOCIEDAD DE INVERSIONES' AS DescripParticipe,'00000001' AS CodAnalitica,0 as MontoMovimiento"
'                    Set adoRegistro = .Execute
'                Else
'                    strSQL = "SELECT '' AS Item,'BLANCO SOCIEDAD DE INVERSIONES' AS DescripParticipe,'00000001' AS CodAnalitica,0 as MontoMovimiento"
'                    Set adoRegistro = .Execute
'                End If
'            End With

        Call ConfiguraRecordsetAuxiliarGastos
            
        With rs
            .ActiveConnection = gstrConnectConsulta
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Open strSQL
            If .RecordCount > 0 Then
                    .MoveFirst
                       i = 1
            Do While Not .EOF
                 
                    adoRegistroAuxGastos.AddNew
                    For Each adoField In adoRegistroAuxGastos.Fields
                            If (adoField.Name = "Item") Then
                                adoRegistroAuxGastos.Fields(adoField.Name) = i
                            Else
                                adoRegistroAuxGastos.Fields(adoField.Name) = rs.Fields(adoField.Name)
                            End If
                            
                    Next
                    dblSubTotal = dblSubTotal + rs.Fields("MontoMovimiento")
                    dblIGV = Round((dblSubTotal * gdblTasaIgv), 2)
                    dblTotal = dblSubTotal + dblIGV
                    
                    i = i + 1
                    adoRegistroAuxGastos.Update
                    rs.MoveNext
                  
            Loop
            adoRegistroAuxGastos.MoveFirst
            End If
        End With
        
        End If
        
                
        
        tdgGastos.DataSource = adoRegistroAuxGastos
        tdgGastos.Refresh
        
        txtSubTotal.Text = CStr(dblSubTotal)
        txtIgv.Text = CStr(dblIGV)
        txtTotal.Text = CStr(dblTotal)
'
'      With gGastos
'        .DefaultFields = False
'        .Dataset.ADODataset.ConnectionString = gstrConnectConsulta
'        .Dataset.ADODataset.CursorLocation = clUseClient
'        .Dataset.Active = False
'        .Dataset.ADODataset.CommandText = strSQL
'        .Dataset.DisableControls
'        .Dataset.Active = True
'        .KeyField = "DescripParticipe"
'     End With
    
    
'    txtSubTotal.Text = gGastos.Columns.ColumnByFieldName("MontoMovimiento").SummaryFooterValue
    
'    With adoComm
'
'        strSQL = "SELECT fg.NumGasto, fg.CodAnalitica, fg.CodCuenta " & _
'             "FROM OrdenPago op " & _
'             "INNER JOIN FondoGasto fg ON (op.CodFondo = fg.CodFondo AND op.CodAdministradora = fg.CodAdministradora AND " & _
'             "op.NumGasto = fg.NumGasto) " & _
'             "JOIN Moneda MO ON (MO.CodMoneda = op.CodMoneda) " & _
'             "WHERE op.CodFondo='" & strCodFondo & "' " & _
'               "AND op.CodAdministradora='" & gstrCodAdministradora & "' " & _
'               "AND fg.CodProveedor = '" & lblCodProveedor.Caption & "' " & _
'               "AND op.CodMoneda = '" & strCodMoneda & "' " & _
'               "AND op.Estado = '01'" & _
'               "AND op.NumOrdenPago = '" & strNumOrdenPago & "'"
'
'        Set adoRegistro = .Execute
'
'        If Not adoRegistro.EOF Then
'
'            strCodAnalitica = adoRegistro("CodAnalitica")
'            strCodGasto = adoRegistro("NumGasto")
'            strCodCuenta = adoRegistro("CodCuenta")
'        End If
'
'    End With
    
'''    lblAnalitica.Caption = Trim(tdgPendientes.Columns(9).Value) & " - " & strCodAnalitica
        
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
    
    If strEstadoRegCompra <> Estado_Registro_Contabilizado Then
        Call CargarOrdenesPago
    End If
    
    
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


Private Sub cboPorcenDetraccion_Click()
    
    dblPorcenDetraccion = CDbl(cboPorcenDetraccion.Text) * 0.01
    
    Call Calculos
    
End Sub

Private Sub cboTipoComprobante_Click()

    Dim adoRegistro     As ADODB.Recordset, intRegistro As Long
    
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
            
            'ACC 12/07/2010
            'Poner en un valor consistente al combo Afecto
            If strIndImpuesto = Valor_Indicador Then
                cboAfectacion.ListIndex = ObtenerItemLista(arrAfectacion(), Codigo_Afecto)
                cboAfectacion.Enabled = False
            Else
                cboAfectacion.Enabled = True
            End If
            
            If strCodTipoComprobante = Codigo_Comprobante_Recibo_Honorarios Then
                If cboPorcenDetraccion.ListCount > 0 Then cboPorcenDetraccion.ListIndex = 2 '10%
                
                intRegistro = ObtenerItemLista(arrMonedaDetraccion(), strCodMoneda)
                If intRegistro >= 0 Then cboMonedaDetraccion.ListIndex = intRegistro
                
                cboPorcenDetraccion.Enabled = False
                cboMonedaDetraccion.Enabled = True
            Else
                intRegistro = ObtenerItemLista(arrMonedaDetraccion(), Codigo_Moneda_Local)
                If intRegistro >= 0 Then cboMonedaDetraccion.ListIndex = intRegistro
                cboPorcenDetraccion.Enabled = True
                cboMonedaDetraccion.Enabled = False
            End If
            
'''            strCtaImpuesto = ObtenerCuentaAdministracion("025", "R")
'''            If strIndRetencion = Valor_Indicador Then strCtaImpuesto = ObtenerCuentaAdministracion("036", "R")
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
    datFechaConsulta = gdatFechaActual
    datFechaConsulta = dtpFechaPago.Value
'    If Not EsDiaUtil(datFechaConsulta) Then
'        datFechaConsulta = AnteriorDiaUtil(datFechaConsulta)
'    End If
    
    dtpFechaTipoCambioPago.Value = datFechaConsulta
    
    If strCodValorTipoCambio = Codigo_Valor_TipoCambioCompra Then
        'txtTipoCambioPago.Text = CStr(ObtenerTipoCambio(Codigo_TipoCambio_Sunat, Codigo_Valor_TipoCambioCompra, datFechaConsulta, strCodMoneda))
        'Para el caso de detraccion debe tomar el tipo de cambio de la fecha del comprobante
        'If strDetraccionSiNo = Codigo_Respuesta_Si And dtpFechaPago.Value <> dtpFechaComprobante.Value Then
            txtTipoCambioPago.Text = CStr(ObtenerTipoCambioMoneda(Codigo_TipoCambio_SBS, Codigo_Valor_TipoCambioCompra, dtpFechaComprobante.Value, Codigo_Moneda_Local, strCodMoneda))
            dtpFechaTipoCambioPago.Value = dtpFechaComprobante.Value
        'End If
    Else
        'txtTipoCambioPago.Text = CStr(ObtenerTipoCambio(Codigo_TipoCambio_Sunat, Codigo_Valor_TipoCambioVenta, datFechaConsulta, strCodMoneda))
        'Debe tomar el tipo de cambio de la fecha del comprobante si es factura
        'If strDetraccionSiNo = Codigo_Respuesta_Si And dtpFechaPago.Value <> dtpFechaComprobante.Value And strCodTipoComprobante = Codigo_Comprobante_Factura Then
        'If dtpFechaPago.Value <> dtpFechaComprobante.Value And strCodTipoComprobante = Codigo_Comprobante_Factura Then
            txtTipoCambioPago.Text = CStr(ObtenerTipoCambioMoneda(Codigo_TipoCambio_SBS, Codigo_Valor_TipoCambioVenta, dtpFechaComprobante.Value, Codigo_Moneda_Local, strCodMoneda))
            dtpFechaTipoCambioPago.Value = dtpFechaComprobante.Value
        'End If
    End If

End Sub

Private Sub cmdAdicionarGasto_Click()
Dim strMsgError As String
Dim strPorDefecto As String

'On Error GoTo err

    If cboGasto.ListIndex <= 0 Then
        strMsgError = "Debe seleccionar un Gasto."
        GoTo err
    End If
    
    If gGastos.Columns.ColumnByFieldName("DescripGasto").Value <> "" Or gGastos.Count = 0 Then
        gGastos.Dataset.Insert
    End If
    
    gGastos.Dataset.Edit
   
    gGastos.Columns.ColumnByFieldName("item").Value = gGastos.Count
    gGastos.Columns.ColumnByFieldName("NumOrdenPago").Value = strNumOrdenPago
    gGastos.Columns.ColumnByFieldName("CodFile").Value = ""
    gGastos.Columns.ColumnByFieldName("CodAnalitica").Value = ""
    gGastos.Columns.ColumnByFieldName("DescripGasto").Value = "" 'cboGasto.Text
    gGastos.Columns.ColumnByFieldName("CodMoneda").Value = ""
    gGastos.Columns.ColumnByFieldName("MontoSubTotal").Value = 0
    gGastos.Columns.ColumnByFieldName("MontoImpuesto").Value = 0
    gGastos.Columns.ColumnByFieldName("TasaImpuesto").Value = 0
    gGastos.Columns.ColumnByFieldName("MontoGasto").Value = 0
    gGastos.Columns.ColumnByFieldName("CodDetalleFile").Value = ""
    gGastos.Columns.ColumnByFieldName("NumGasto").Value = ""
    
    Set adoRegistro = New ADODB.Recordset

    With adoComm

        'el CodCreditoFiscal lo jalamos de la tabla fondo gasto o del form JCB?
        .CommandText = "SELECT fg.CodCreditoFiscal, fg.DescripGasto, op.MontoOrdenPago, op.CodMoneda, fg.CodFile, fg.CodCuenta, fg.CodAnalitica, fg.NumGasto " & _
            "FROM OrdenPago op INNER JOIN FondoGasto fg ON op.CodFondo = fg.CodFondo AND op.CodAdministradora = fg.CodAdministradora AND op.NumGasto = fg.NumGasto " & _
            "WHERE op.NumOrdenPago = " & strNumOrdenPago & " " & _
              "AND op.CodFondo='" & strCodFondo & "' AND op.CodAdministradora='" & gstrCodAdministradora & "'"
        Set adoRegistro = .Execute

        If Not adoRegistro.EOF Then
        
            gGastos.Columns.ColumnByFieldName("DescripGasto").Value = adoRegistro("DescripGasto")
            
            gGastos.Columns.ColumnByFieldName("NumGasto").Value = adoRegistro("NumGasto")
      
            If Trim(strNumOrdenPagoLista) = "''" Then
                strNumOrdenPagoLista = "'" & strNumOrdenPago & "'" 'adoRegistro("NumGasto")
            Else
                strNumOrdenPagoLista = strNumOrdenPagoLista & ",'" & strNumOrdenPago & "'" 'adoRegistro("NumGasto")
            End If
       
            If strCodAfectacion <> Codigo_Afecto Then
            'If adoRegistro("CodCreditoFiscal") = Codigo_Tipo_Credito_RentaNoGravada Or adoRegistro("CodCreditoFiscal") = Codigo_Tipo_Credito_AdquisicionesNoGravada Then
                gGastos.Columns.ColumnByFieldName("MontoSubTotal").Value = adoRegistro("MontoOrdenPago")
                gGastos.Columns.ColumnByFieldName("MontoImpuesto").Value = 0
                gGastos.Columns.ColumnByFieldName("MontoGasto").Value = adoRegistro("MontoOrdenPago")
            Else
                gGastos.Columns.ColumnByFieldName("MontoSubTotal").Value = Round(adoRegistro("MontoOrdenPago"), 2) '/ (gdblTasaIgv + 1), 2)
                gGastos.Columns.ColumnByFieldName("MontoImpuesto").Value = Round(adoRegistro("MontoOrdenPago") * (gdblTasaIgv), 2) ' adoRegistro("MontoOrdenPago") - gGastos.Columns.ColumnByFieldName("MontoSubTotal").Value
                gGastos.Columns.ColumnByFieldName("MontoGasto").Value = adoRegistro("MontoOrdenPago") + gGastos.Columns.ColumnByFieldName("MontoImpuesto").Value 'adoRegistro("MontoOrdenPago")
            End If
            
            dblMontoSubtotal = dblMontoSubtotal + adoRegistro("MontoOrdenPago")
    
            gGastos.Columns.ColumnByFieldName("TasaImpuesto").Value = gdblTasaIgv
            
            gGastos.Columns.ColumnByFieldName("CodMoneda").Value = adoRegistro("CodMoneda")
            gGastos.Columns.ColumnByFieldName("CodAnalitica").Value = adoRegistro("CodAnalitica")
        End If
        
        gGastos.Columns.ColumnByFieldName("CodFile").Value = Trim(adoRegistro("CodFile"))
        
        .CommandText = "SELECT CodDetalleFile FROM InversionDetalleFile " & _
            "WHERE CodFile='" & Trim(adoRegistro("CodFile")) & "' AND DescripDetalleFile='" & adoRegistro("CodCuenta") & "'"
        Set adoRegistro = .Execute

        If Not adoRegistro.EOF Then
            'Aqui guardo el CodFile o el CodDetalleFile
            gGastos.Columns.ColumnByFieldName("CodDetalleFile").Value = adoRegistro("CodDetalleFile")
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With

    gGastos.Dataset.Post
    
    gGastos.Dataset.Refresh
            
    lblMontoGasto.Caption = gGastos.Columns.ColumnByFieldName("MontoGasto").SummaryFooterValue
    
    Call Calculos
    
    If strEstadoRegCompra <> Estado_Registro_Contabilizado Then
        Call CargarOrdenesPago
    End If
   
    cboGasto.ListIndex = 0
    numContadorGastos = numContadorGastos + 1
    
Exit Sub
err:
If strMsgError = "" Then strMsgError = err.Description
MsgBox strMsgError, vbInformation, App.Title
End Sub

Private Sub cmdEliminarGasto_Click()
Dim strMsgError As String
Dim i As Integer

On Error GoTo err

    If gGastos.Count = 1 Then
    
    
        'Elimina de la lista de elementos seleccionados (strNumOrdenPagoLista) el elemento que se esta sacando de la grilla
        If InStr(1, strNumOrdenPagoLista, gGastos.Columns.ColumnByFieldName("NumOrdenPago").Value) > 0 Then
            'Es el ultimo elemento
            If InStr(1, strNumOrdenPagoLista, "'" & gGastos.Columns.ColumnByFieldName("NumOrdenPago").Value & "'") + Len(gGastos.Columns.ColumnByFieldName("NumOrdenPago").Value & "'") = Len(strNumOrdenPagoLista) Then
                If Len(gGastos.Columns.ColumnByFieldName("NumOrdenPago").Value) = Len(strNumOrdenPagoLista) - 2 Then 'hay solo un elemento
                    strNumOrdenPagoLista = "''" 'Replace(strNumOrdenPagoLista, gGastos.Columns.ColumnByFieldName("NumOrdenPago").Value, Valor_Caracter)
                Else
                    strNumOrdenPagoLista = Replace(strNumOrdenPagoLista, ",'" & gGastos.Columns.ColumnByFieldName("NumOrdenPago").Value & "'", Valor_Caracter)
                End If
            Else 'no es el ultimo elemento
                strNumOrdenPagoLista = Replace(strNumOrdenPagoLista, "'" & gGastos.Columns.ColumnByFieldName("NumOrdenPago").Value & "',", Valor_Caracter)
            End If
        End If
        
        gGastos.Dataset.Edit
        
        gGastos.Columns.ColumnByFieldName("Item").Value = 1
        gGastos.Columns.ColumnByFieldName("NumOrdenPago").Value = ""
        gGastos.Columns.ColumnByFieldName("CodFile").Value = ""
        gGastos.Columns.ColumnByFieldName("CodAnalitica").Value = ""
        gGastos.Columns.ColumnByFieldName("DescripGasto").Value = ""
        gGastos.Columns.ColumnByFieldName("CodMoneda").Value = ""
        gGastos.Columns.ColumnByFieldName("MontoSubTotal").Value = 0
        gGastos.Columns.ColumnByFieldName("MontoImpuesto").Value = 0
        gGastos.Columns.ColumnByFieldName("TasaImpuesto").Value = 0
        gGastos.Columns.ColumnByFieldName("MontoGasto").Value = 0
        gGastos.Columns.ColumnByFieldName("NumGasto").Value = 0
        
        gGastos.Dataset.Post
        
    Else
        
        'Elimina de la lista de elementos seleccionados (strNumOrdenPagoLista) el elemento que se esta sacando de la grilla
        If InStr(1, strNumOrdenPagoLista, gGastos.Columns.ColumnByFieldName("NumOrdenPago").Value) > 0 Then
            'Es el ultimo elemento
            If InStr(1, strNumOrdenPagoLista, "'" & gGastos.Columns.ColumnByFieldName("NumOrdenPago").Value & "'") + Len(gGastos.Columns.ColumnByFieldName("NumOrdenPago").Value & "'") = Len(strNumOrdenPagoLista) Then
                If "'" & Len(gGastos.Columns.ColumnByFieldName("NumOrdenPago").Value) & "'" = Len(strNumOrdenPagoLista) Then 'hay solo un elemento
                    strNumOrdenPagoLista = Replace(strNumOrdenPagoLista, "'" & gGastos.Columns.ColumnByFieldName("NumOrdenPago").Value & "'", Valor_Caracter)
                Else
                    strNumOrdenPagoLista = Replace(strNumOrdenPagoLista, ",'" & gGastos.Columns.ColumnByFieldName("NumOrdenPago").Value & "'", Valor_Caracter)
                End If
            Else 'no es el ultimo elemento
                strNumOrdenPagoLista = Replace(strNumOrdenPagoLista, "'" & gGastos.Columns.ColumnByFieldName("NumOrdenPago").Value & "',", Valor_Caracter)
            End If
        End If
        
        gGastos.Dataset.Delete
                    
        gGastos.Dataset.First
        Do While Not gGastos.Dataset.EOF
            
            If gGastos.Columns.ColumnByFieldName("Item").Value > 0 Then
                i = i + 1
                gGastos.Dataset.Edit
                gGastos.Columns.ColumnByFieldName("Item").Value = i
                gGastos.Dataset.Post
            End If
            
            gGastos.Dataset.Next
        Loop
        If gGastos.Dataset.State = dsEdit Or gGastos.Dataset.State = dsInsert Then
            gGastos.Dataset.Post
        End If
    
    End If
    
    lblMontoGasto.Caption = gGastos.Columns.ColumnByFieldName("MontoGasto").SummaryFooterValue
    
    Call Calculos
    
    If strEstadoRegCompra <> Estado_Registro_Contabilizado Then
        
        If strEstado = Reg_Edicion Then
            adoComm.CommandText = "UPDATE OrdenPago " & _
                     "SET Estado = '01' " & _
                     "WHERE CodFondo='" & strCodFondo & "' " & _
                       "AND CodAdministradora='" & gstrCodAdministradora & "' " & _
                       "AND CodMoneda = '" & strCodMoneda & "' " & _
                       "AND Estado = '04' " & _
                       "AND NumOrdenPago = '" & strNumOrdenPago & "'"
            adoConn.Execute adoComm.CommandText
        End If
        
        Call CargarOrdenesPago
    End If
    
'    txtSubTotal.Text = gGastos.Columns.ColumnByFieldName("MontoSubTotal").SummaryFooterValue
'    txtIgv.Text = gGastos.Columns.ColumnByFieldName("MontoImpuesto").SummaryFooterValue
'    txtTotal.Text = gGastos.Columns.ColumnByFieldName("MontoGasto").SummaryFooterValue
    If numContadorGastos > 0 Then
        numContadorGastos = numContadorGastos - 1
    End If
    
Exit Sub
err:
If strMsgError = "" Then strMsgError = err.Description
MsgBox strMsgError, vbInformation, App.Title
End Sub

Private Sub cmdProveedor_Click()

    'gstrFormulario = "frmGenerarOrdenCobroProvision"
    'frmBusquedaInstitucionPersona.lblTipoInstitucion = Codigo_Tipo_Persona_Proveedor
    'frmBusquedaInstitucionPersona.Caption = "B�squeda de Proveedores"
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
        
        frmBus.Caption = " Relaci�n de Proveedores"
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
    
    If strEstadoRegCompra <> Estado_Registro_Contabilizado Then
        Call CargarOrdenesPago
    End If
   
End Sub

Private Sub dtpFechaComprobante_Change()

    If dtpFechaComprobante.Value > gdatFechaActual Then
        MsgBox "La Fecha de comprobante debe ser igual o anterior a la fecha actual...se cambiar� por la fecha actual !", vbInformation, Me.Caption
        dtpFechaComprobante.Value = gdatFechaActual
    End If
    
    Call cboDetraccion_Click

End Sub



Private Sub dtpFechaPago_Change()

    If Not EsDiaUtil(dtpFechaPago.Value) Then
        MsgBox "La Fecha no es un d�a �til...se cambiar� por una fecha correcta !", vbInformation, Me.Caption
        If dtpFechaPago.Value >= gdatFechaActual Then
            dtpFechaPago.Value = AnteriorDiaUtil(dtpFechaPago.Value)
        Else
            dtpFechaPago.Value = ProximoDiaUtil(dtpFechaPago.Value)
        End If
    End If

    Call cboDetraccion_Click
    
End Sub



Private Sub Form_Activate()

    Call CargarReportes
    
End Sub

Private Sub CargarReportes()

    'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Registro de Compras"
    'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
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
    
    Call ValidarPermisoUsoControl(Trim(gstrLogin), Me, Trim(App.Title) + Separador_Codigo_Objeto + _
    gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)

    CentrarForm Me
    
    Call ValidaExisteTipoCambio(Codigo_TipoCambio_SBS, gstrFechaActual)
    
    
End Sub


Private Sub DarFormato()

    Dim intCont As Integer
    
    For intCont = 0 To (lblDescrip.Count - 1)
        Call FormatoEtiqueta(lblDescrip(intCont), vbLeftJustify)
    Next
    
    For intCont = 0 To (fraCompras.Count - 1)
        Call FormatoMarco(fraCompras(intCont))
    Next
            
End Sub

Public Sub Buscar()
            
    strSQL = "SELECT NumRegistro,CodTipoComprobante,CodProveedor,DescripRegistro,RC.CodMoneda,ValorTotal, " & _
        "TCP.DescripTipoComprobantePago DescripTipoComprobante, CodSigno,FechaRegistro,DescripPersona DescripProveedor,RC.NumGasto " & _
        "FROM RegistroCompra RC JOIN TipoComprobantePago TCP ON(TCP.CodTipoComprobantePago=RC.CodTipoComprobante) " & _
        "JOIN Moneda MON ON(MON.CodMoneda=RC.CodMoneda) " & _
        "LEFT JOIN InstitucionPersona IP ON(IP.CodPersona=RC.CodProveedor AND IP.TipoPersona='" & Codigo_Tipo_Persona_Proveedor & "') " & _
        "WHERE (FechaRegistro>='" & Convertyyyymmdd(dtpFechaDesde.Value) & "' AND FechaRegistro<'" & Convertyyyymmdd(DateAdd("d", 1, dtpFechaHasta.Value)) & "') AND " & _
        "CodAdministradora='" & gstrCodAdministradora & "' AND CodFondo='" & strCodFondo & "' AND RC.Estado NOT IN ('03')" & _
        "ORDER BY NumRegistro"

    strEstado = Reg_Defecto
    
    With gLista
        .DefaultFields = False
        .Dataset.ADODataset.ConnectionString = gstrConnectConsulta
        .Dataset.ADODataset.CursorLocation = clUseClient
        .Dataset.Active = False
        .Dataset.ADODataset.CommandText = strSQL
        .Dataset.DisableControls
        .Dataset.Active = True
        .KeyField = "NumRegistro"
    End With


    If gLista.Count > 0 Then strEstado = Reg_Consulta
    dtpFechaPago.MinDate = 0
            
End Sub
Private Sub CargarListas()
            
    '*** Fondos ***
    strSQL = "{ call up_ACSelDatosParametro(29,'" & gstrCodAdministradora & "') }"
    CargarControlLista strSQL, cboFondo, arrFondo(), Valor_Caracter
    
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
    
    '*** Tipo de Comprobante Sunat ***
    strSQL = "SELECT CodTipoComprobantePago CODIGO,DescripTipoComprobantePago DESCRIP From TipoComprobantePago ORDER BY DescripTipoComprobantePago"
    CargarControlLista strSQL, cboTipoComprobante, arrTipoComprobante(), Sel_Defecto
            
    '*** Afectaci�n ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP From AuxiliarParametro WHERE CodTipoParametro='AFEIMP' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboAfectacion, arrAfectacion(), Valor_Caracter
    
    '*** Tipo Cr�dito Fiscal ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP From AuxiliarParametro WHERE CodTipoParametro='CREFIS' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboCreditoFiscal, arrCreditoFiscal(), Sel_Defecto
    
    '*** Moneda ***
    strSQL = "{ call up_ACSelDatos(2) }"
    CargarControlLista strSQL, cboMoneda, arrMoneda(), Valor_Caracter
    CargarControlLista strSQL, cboMonedaUnico, arrMonedaUnico(), Valor_Caracter
    CargarControlLista strSQL, cboMonedaDetraccion, arrMonedaDetraccion(), Valor_Caracter
    
    '*** Detracci�n ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP From AuxiliarParametro WHERE CodTipoParametro='RESPSN' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboDetraccion, arrDetraccion(), ""
    
    '*** Comisionistas ***
    strSQL = "SELECT CodPersona CODIGO,RazonSocial DESCRIP From InstitucionPersona WHERE TipoPersona='10' ORDER BY CodPersona"
    CargarControlLista strSQL, cboGasto, arrGasto(), ""
    
        
    '*** Forma de Pago ***
'    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='MEDPAG' ORDER BY DescripParametro"
'    CargarControlLista strSQL, cboFormaPagoUnico, arrFormaPagoUnico(), Valor_Caracter
'    CargarControlLista strSQL, cboFormaPagoDetraccion, arrFormaPagoDetraccion(), Valor_Caracter
    
    '*** Valor de Tipo de Cambio ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP From AuxiliarParametro WHERE CodTipoParametro='CLSVTC' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboTipoValorCambio, arrTipoValorCambio(), ""
    
End Sub
Private Sub ConfiguraRecordsetAuxiliar()

    Set adoRegistroAux = New ADODB.Recordset

    With adoRegistroAux
       .CursorLocation = adUseClient
       .Fields.Append "CodFondo", adVarChar, 3
       .Fields.Append "CodAdministradora", adVarChar, 3
       .Fields.Append "NumComprobantePago", adVarChar, 10
       .Fields.Append "SecComprobantePago", adInteger, 4
       .Fields.Append "NumOrdenPago", adVarChar, 10
       .Fields.Append "NumGasto", adInteger, 4
       .Fields.Append "DescripGasto", adVarChar, 60
       .Fields.Append "CodCuenta", adVarChar, 10
       .Fields.Append "CodFile", adVarChar, 3
       .Fields.Append "CodAnalitica", adVarChar, 8
       .Fields.Append "FechaPago", adDate, 8
       .Fields.Append "CodMoneda", adVarChar, 2
       .Fields.Append "MontoOrdenPago", adDecimal, 19
'       .CursorType = adOpenStatic
       .LockType = adLockBatchOptimistic
    End With

    With adoRegistroAux.Fields.Item("MontoOrdenPago")
        .Precision = 19
        .NumericScale = 2
    End With
    
' ComprobantePagoDetalle.CodFondo                                   CodigoCorto     NOT NULL,
' ComprobantePagoDetalle.CodAdministradora                          CodigoCorto     NOT NULL,
' ComprobantePagoDetalle.NumComprobantePago                         CodigoMediano   NOT NULL,
' ComprobantePagoDetalle.SecComprobantePago                         int             NOT NULL,
' ComprobantePagoDetalle.NumOrdenPago   (OrdenPago.NumOrdenPago)    CodigoMediano   NOT NULL,

' OrdenPago.NumGasto   (FondoGasto.NumGasto)

' FondoGasto.DescripGasto                                           varchar(60)     NOT NULL,
' FondoGasto.CodCuenta                                              CodigoMediano   NOT NULL,
' FondoGasto.CodFile                                                CodigoCorto     NOT NULL,
' FondoGasto.CodAnalitica                                           char(8)         NOT NULL,

' OrdenPago.FechaPago                                               datetime        NOT NULL,

' ComprobantePagoDetalle.CodMoneda (OrdenPago.CodMoneda)            Codigo          NOT NULL,
' ComprobantePagoDetalle.MontoComprobanteDetalle (OrdenPago.MontoOrdenPago)     decimal(19,2)   NOT NULL, --del archivo original
'

End Sub
Private Sub ConfiguraRecordsetAuxiliarGastos()

    Set adoRegistroAuxGastos = New ADODB.Recordset

    With adoRegistroAuxGastos
       .CursorLocation = adUseClient
       .Fields.Append "Item", adInteger, 2
       .Fields.Append "DescripParticipe", adChar, 80
       .Fields.Append "CodAnalitica", adChar, 10
       .Fields.Append "MontoMovimiento", adDecimal, 10
       .LockType = adLockBatchOptimistic
    End With

    With adoRegistroAuxGastos.Fields.Item("MontoMovimiento")
        .Precision = 19
        .NumericScale = 8
    End With
    
    adoRegistroAuxGastos.Open
    
End Sub
Private Sub InicializarValores()

    strEstado = Reg_Defecto
    tabRegistroCompras.Tab = 0
    strNumOrdenPagoLista = "''"
    txtTipoCambioPago.Text = 0
    
    dtpFechaDesde.Value = gdatFechaActual
    dtpFechaHasta.Value = gdatFechaActual
    
    ConfGrid gLista, False, False, False, False
    ConfGrid gGastos, True, False, False, False
    
        With tabRegistroCompras
            .TabEnabled(1) = False
            .TabEnabled(3) = False
        End With
    
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    Set cmdOpcion.FormularioActivo = Me
    Set cmdContabilizar.FormularioActivo = Me
    
End Sub
Private Sub Form_Unload(Cancel As Integer)

    Call OcultarReportes
    frmMainMdi.stbMdi.Panels(3).Text = "Acci�n"
    Set frmGenerarOrdenCobroProvision = Nothing
    
End Sub

Private Sub lblMontoTotal_Change()

    Call FormatoMillarEtiqueta(lblMontoTotal, Decimales_Monto)
    
End Sub

Private Sub tabRegistroCompras_Click(PreviousTab As Integer)
    cmdAccion.Visible = False
    cmdContabilizar.Visible = False
    Select Case tabRegistroCompras.Tab
        Case 1, 2, 3
            cmdAccion.Visible = True
            cmdContabilizar.Visible = False
            If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vQuery)
            If strEstado = Reg_Defecto Then tabRegistroCompras.Tab = 0
    End Select
    
End Sub


Private Sub tdgGastos_AfterColEdit(ByVal ColIndex As Integer)
            Dim dblSubTotal As Double
            Dim dblIGV As Double
            Dim dblTotal As Double
            
            adoRegistroAuxGastos.MoveFirst
            Do While Not adoRegistroAuxGastos.EOF
                    dblSubTotal = dblSubTotal + adoRegistroAuxGastos.Fields("MontoMovimiento")
                    adoRegistroAuxGastos.MoveNext
            Loop
        
            dblIGV = Round((dblSubTotal * gdblTasaIgv), 2)
            dblTotal = dblSubTotal + dblIGV
            
            txtSubTotal.Text = CStr(dblSubTotal)
            txtIgv.Text = CStr(dblIGV)
            txtTotal.Text = CStr(dblTotal)

End Sub


'''Private Sub tdgConsulta_Click()
'''
'''    tdgConsulta.HeadBackColor = &HFFC0C0
'''    tdgPendientes.HeadBackColor = &H8000000F
'''
'''End Sub
'''
'''Private Sub tdgConsulta_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
'''
'''    If ColIndex = 5 Then
'''        Call DarFormatoValor(Value, Decimales_Monto)
'''    End If
'''
'''End Sub

'''Private Sub tdgPendientes_Click()
'''
'''    tdgPendientes.HeadBackColor = &HFFC0C0
'''    tdgConsulta.HeadBackColor = &H8000000F
'''
'''End Sub
'''
'''Private Sub tdgPendientes_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
'''
'''    If ColIndex = 7 Or ColIndex = 8 Then
'''        Call DarFormatoValor(Value, Decimales_Monto)
'''    End If
'''
'''End Sub

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

Private Sub txtMontoUnico_Change()

    Call FormatoCajaTexto(txtMontoUnico, Decimales_Monto)
    
End Sub

Private Sub txtMontoUnico_KeyPress(KeyAscii As Integer)

'    Call ValidaCajaTexto(KeyAscii, "M", txtMontoUnico, Decimales_Monto)
'    If KeyAscii = vbKeyReturn Then Call CalculosPago
    
End Sub

Private Sub txtSubTotal_Change()

    Call FormatoCajaTexto(txtSubTotal, Decimales_Monto)
                
End Sub

Private Sub txtSubTotal_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtSubTotal, Decimales_Monto)
    If KeyAscii = vbKeyReturn Then Call Calculos
    
End Sub

Private Sub txtTipoCambioPago_Change()

    Call Calculos
    
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

Private Sub FormatoGrillaGastos2(ByRef strMsgError As String) 'JCB
Dim rsGastos As New ADODB.Recordset
On Error GoTo err
    '********FORMATO GRILLA DE GASTOS
'    rsGastos.Fields.Append "Item", adInteger, , adFldRowID
'    rsGastos.Fields.Append "NumOrdenPago", adVarChar, 10, adFldIsNullable
    rsGastos.Fields.Append "DescripParticipe", adVarChar, 15, adFldIsNullable
'    rsGastos.Fields.Append "CodFile", adVarChar, 10, adFldIsNullable
    rsGastos.Fields.Append "CodAnalitica", adVarChar, 10, adFldIsNullable
'    rsGastos.Fields.Append "DescripGasto", adVarChar, 100, adFldIsNullable
'    rsGastos.Fields.Append "CodMoneda", adVarChar, 2, adFldIsNullable
'    rsGastos.Fields.Append "MontoSubTotal", adDouble, , adFldIsNullable
'    rsGastos.Fields.Append "MontoImpuesto", adDouble, , adFldIsNullable
'    rsGastos.Fields.Append "TasaImpuesto", adDouble, , adFldIsNullable
'    rsGastos.Fields.Append "MontoGasto", adDouble, , adFldIsNullable
'    rsGastos.Fields.Append "CodDetalleFile", adVarChar, 10, adFldIsNullable
    rsGastos.Fields.Append "MontoMovimiento", adInteger, 10, adFldIsNullable
    

    rsGastos.Open
    rsGastos.AddNew

'    rsGastos.Fields("Item") = 1
'    rsGastos.Fields("NumOrdenPago") = ""
    rsGastos.Fields("DescripParticipe") = ""
'    rsGastos.Fields("CodFile") = ""
    rsGastos.Fields("CodAnalitica") = ""
'    rsGastos.Fields("DescripGasto") = ""
'    rsGastos.Fields("CodMoneda") = ""
'    rsGastos.Fields("MontoSubTotal") = 0
'    rsGastos.Fields("MontoImpuesto") = 0
'    rsGastos.Fields("TasaImpuesto") = 0
'    rsGastos.Fields("MontoGasto") = 0
'    rsGastos.Fields("CodDetalleFile") = ""
    rsGastos.Fields("MontoMovimiento") = 0
    
    
    Set gGastos.DataSource = Nothing
    mostrarDatosGridSQL gGastos, rsGastos, strMsgError
    If strMsgError <> "" Then GoTo err

Exit Sub
err:
If strMsgError = "" Then strMsgError = err.Description
End Sub






