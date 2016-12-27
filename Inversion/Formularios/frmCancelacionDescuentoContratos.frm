VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{830D5A06-9B70-4F8C-98B6-7A19C4E7760B}#1.0#0"; "TAMControls.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmCancelacionDescuentoContratos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancelacion de Descuentos de Contratos Futuros"
   ClientHeight    =   8865
   ClientLeft      =   7080
   ClientTop       =   3615
   ClientWidth     =   14415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   14415
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboTipoOrden 
      Height          =   315
      Left            =   3180
      Style           =   2  'Dropdown List
      TabIndex        =   77
      Top             =   9270
      Visible         =   0   'False
      Width           =   4725
   End
   Begin VB.ComboBox cboOrigen 
      Height          =   315
      Left            =   9240
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   75
      Top             =   9060
      Visible         =   0   'False
      Width           =   4455
   End
   Begin TAMControls2.ucBotonEdicion2 cmdOpcion 
      Height          =   735
      Left            =   720
      TabIndex        =   74
      Top             =   8010
      Width           =   4200
      _ExtentX        =   7408
      _ExtentY        =   1296
      Buttons         =   3
      Caption0        =   "&Nuevo"
      Tag0            =   "0"
      ToolTipText0    =   "Nuevo"
      Caption1        =   "&Eliminar"
      Tag1            =   "4"
      ToolTipText1    =   "Eliminar"
      Caption2        =   "&Buscar"
      Tag2            =   "5"
      ToolTipText2    =   "Buscar"
      UserControlWidth=   4200
   End
   Begin TAMControls2.ucBotonEdicion2 cmdAccion 
      Height          =   735
      Left            =   5160
      TabIndex        =   73
      Top             =   8040
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
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
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
      Left            =   12690
      Picture         =   "frmCancelacionDescuentoContratos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   8010
      Width           =   1200
   End
   Begin MSAdodcLib.Adodc adoConsulta 
      Height          =   330
      Left            =   9390
      Top             =   8220
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
   Begin TabDlg.SSTab tabRFCortoPlazo 
      Height          =   7845
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14355
      _ExtentX        =   25321
      _ExtentY        =   13838
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
      TabPicture(0)   =   "frmCancelacionDescuentoContratos.frx":0582
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraCriterio"
      Tab(0).Control(1)=   "tdgConsulta"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Datos Orden Inversión"
      TabPicture(1)   =   "frmCancelacionDescuentoContratos.frx":059E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraDatosTitulo"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraDatosBasicos"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "fraResumen"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.Frame fraResumen 
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
         Height          =   3615
         Left            =   120
         TabIndex        =   56
         Top             =   4080
         Width           =   14085
         Begin DXDBGRIDLibCtl.dxDBGrid gDetalle 
            Height          =   2535
            Left            =   120
            OleObjectBlob   =   "frmCancelacionDescuentoContratos.frx":05BA
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   270
            Width           =   13815
         End
         Begin TAMControls.TAMTextBox txtDeudaFecha 
            Height          =   315
            Left            =   11910
            TabIndex        =   60
            Top             =   2880
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   556
            BackColor       =   16777215
            ForeColor       =   0
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
            Locked          =   -1  'True
            Container       =   "frmCancelacionDescuentoContratos.frx":5EA1
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   2
         End
         Begin TAMControls.TAMTextBox txtMontoRecibido 
            Height          =   315
            Left            =   11910
            TabIndex        =   62
            Top             =   3210
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   556
            BackColor       =   12632319
            ForeColor       =   0
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
            Container       =   "frmCancelacionDescuentoContratos.frx":5EBD
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   2
         End
         Begin TAMControls.TAMTextBox txtSaldoDeuda 
            Height          =   315
            Left            =   1410
            TabIndex        =   64
            Top             =   2880
            Visible         =   0   'False
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   556
            BackColor       =   16777215
            ForeColor       =   0
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
            Locked          =   -1  'True
            Container       =   "frmCancelacionDescuentoContratos.frx":5ED9
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   2
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Deuda"
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
            Left            =   180
            TabIndex        =   65
            Top             =   2940
            Visible         =   0   'False
            Width           =   1110
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto recibido"
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
            Left            =   10260
            TabIndex        =   63
            Top             =   3270
            Width           =   1275
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Deuda a la fecha"
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
            Left            =   10260
            TabIndex        =   61
            Top             =   2940
            Width           =   1485
         End
      End
      Begin VB.Frame fraDatosBasicos 
         Caption         =   "Datos Básicos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2235
         Left            =   120
         TabIndex        =   33
         Top             =   420
         Width           =   14085
         Begin VB.CommandButton cmdBuscarSolicitud 
            Caption         =   "..."
            Height          =   300
            Left            =   4260
            TabIndex        =   66
            Top             =   1800
            Width           =   315
         End
         Begin VB.ComboBox cboEmisor 
            Height          =   315
            Left            =   9300
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   360
            Width           =   4455
         End
         Begin VB.TextBox txtNumOperacionOrig 
            Height          =   315
            Left            =   1830
            TabIndex        =   45
            Top             =   1800
            Width           =   2385
         End
         Begin VB.ComboBox cboLineaCliente 
            Height          =   315
            Left            =   9300
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   1800
            Width           =   4455
         End
         Begin VB.ComboBox cboOperacion 
            Height          =   315
            Left            =   9300
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   1455
            Width           =   4455
         End
         Begin VB.ComboBox cboGestor 
            Height          =   315
            Left            =   9300
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   1080
            Width           =   4455
         End
         Begin VB.ComboBox cboObligado 
            Height          =   315
            Left            =   9300
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   720
            Visible         =   0   'False
            Width           =   4455
         End
         Begin VB.ComboBox cboSubClaseInstrumento 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1830
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   1455
            Width           =   4725
         End
         Begin VB.ComboBox cboFondoOrden 
            Height          =   315
            Left            =   1830
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   360
            Width           =   4725
         End
         Begin VB.ComboBox cboTipoInstrumentoOrden 
            Height          =   315
            Left            =   1830
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   720
            Width           =   4725
         End
         Begin VB.ComboBox cboTitulo 
            Height          =   315
            Left            =   9300
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   360
            Width           =   4185
         End
         Begin VB.ComboBox cboClaseInstrumento 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1830
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   1095
            Width           =   4725
         End
         Begin VB.CheckBox chkTitulo 
            Height          =   255
            Left            =   13560
            TabIndex        =   34
            ToolTipText     =   "Seleccionar Título"
            Top             =   360
            Width           =   255
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Nro. Solicitud"
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
            Index           =   126
            Left            =   360
            TabIndex        =   55
            Top             =   1860
            Width           =   1170
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Especificar Línea"
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
            Index           =   115
            Left            =   7170
            TabIndex        =   54
            Top             =   1860
            Width           =   1500
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Liquidación Operación"
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
            Index           =   28
            Left            =   7170
            TabIndex        =   53
            Top             =   1530
            Width           =   1905
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Gestor"
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
            Index           =   88
            Left            =   7170
            TabIndex        =   52
            Top             =   1185
            Width           =   570
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Obligado"
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
            Index           =   87
            Left            =   7170
            TabIndex        =   51
            Top             =   810
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Clase"
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
            TabIndex        =   50
            Top             =   1185
            Width           =   480
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
            Index           =   0
            Left            =   360
            TabIndex        =   49
            Top             =   435
            Width           =   540
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Instrumento"
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
            TabIndex        =   48
            Top             =   810
            Width           =   1005
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Emisor"
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
            Left            =   7170
            TabIndex        =   47
            Top             =   435
            Width           =   570
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "SubClase"
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
            TabIndex        =   46
            Top             =   1530
            Width           =   810
         End
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
         Height          =   1935
         Left            =   -74880
         TabIndex        =   14
         Top             =   420
         Width           =   14085
         Begin VB.CommandButton cmdEnviar 
            Caption         =   "En&viar"
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
            Left            =   12180
            Picture         =   "frmCancelacionDescuentoContratos.frx":5EF5
            Style           =   1  'Graphical
            TabIndex        =   71
            ToolTipText     =   "Enviar a BackOffice"
            Top             =   1080
            Width           =   1200
         End
         Begin VB.ComboBox cboLineaClienteLista 
            Height          =   315
            Left            =   8880
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   1080
            Width           =   3090
         End
         Begin VB.ComboBox cboEstado 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   1080
            Width           =   4785
         End
         Begin VB.ComboBox cboTipoInstrumento 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   720
            Width           =   4785
         End
         Begin VB.ComboBox cboFondo 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   360
            Width           =   4785
         End
         Begin MSComCtl2.DTPicker dtpFechaOrdenDesde 
            Height          =   315
            Left            =   9600
            TabIndex        =   19
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
            CheckBox        =   -1  'True
            Format          =   175702017
            CurrentDate     =   38785
         End
         Begin MSComCtl2.DTPicker dtpFechaOrdenHasta 
            Height          =   315
            Left            =   11955
            TabIndex        =   20
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
            CheckBox        =   -1  'True
            Format          =   175702017
            CurrentDate     =   38785
         End
         Begin MSComCtl2.DTPicker dtpFechaLiquidacionDesde 
            Height          =   315
            Left            =   9600
            TabIndex        =   21
            Top             =   720
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
            CheckBox        =   -1  'True
            Format          =   175702017
            CurrentDate     =   38785
         End
         Begin MSComCtl2.DTPicker dtpFechaLiquidacionHasta 
            Height          =   315
            Left            =   11955
            TabIndex        =   22
            Top             =   720
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
            CheckBox        =   -1  'True
            Format          =   175702017
            CurrentDate     =   38785
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Especificar Línea"
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
            Index           =   116
            Left            =   7200
            TabIndex        =   32
            Top             =   1155
            Width           =   1500
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
            Index           =   23
            Left            =   240
            TabIndex        =   31
            Top             =   1155
            Width           =   600
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Instrumento"
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
            Left            =   240
            TabIndex        =   30
            Top             =   795
            Width           =   1005
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
            Index           =   21
            Left            =   11280
            TabIndex        =   29
            Top             =   375
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
            Index           =   20
            Left            =   8880
            TabIndex        =   28
            Top             =   435
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
            Index           =   19
            Left            =   240
            TabIndex        =   27
            Top             =   435
            Width           =   540
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Orden"
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
            Index           =   43
            Left            =   7200
            TabIndex        =   26
            Top             =   435
            Width           =   1110
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Liquidación"
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
            Index           =   44
            Left            =   7200
            TabIndex        =   25
            Top             =   795
            Width           =   1560
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
            Index           =   45
            Left            =   8880
            TabIndex        =   24
            Top             =   795
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
            Index           =   46
            Left            =   11280
            TabIndex        =   23
            Top             =   795
            Width           =   510
         End
      End
      Begin VB.Frame fraDatosTitulo 
         Caption         =   "Datos de la Orden"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1425
         Left            =   120
         TabIndex        =   1
         Top             =   2640
         Width           =   14085
         Begin VB.TextBox txtTasa 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   5340
            Locked          =   -1  'True
            MaxLength       =   45
            TabIndex        =   69
            Text            =   "0.0000"
            Top             =   630
            Width           =   1470
         End
         Begin VB.TextBox txtCuotasPago 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   13200
            MaxLength       =   45
            TabIndex        =   68
            Text            =   "1"
            Top             =   240
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.TextBox txtObservacion 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   675
            Left            =   9270
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   630
            Width           =   4470
         End
         Begin VB.ComboBox cboMoneda 
            Height          =   315
            Left            =   1800
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   630
            Width           =   2220
         End
         Begin VB.TextBox txtDescripOrden 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   45
            TabIndex        =   2
            Top             =   990
            Width           =   5010
         End
         Begin MSComCtl2.DTPicker dtpFechaOrden 
            Height          =   315
            Left            =   1800
            TabIndex        =   5
            Top             =   240
            Width           =   1485
            _ExtentX        =   2619
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
            Format          =   175702017
            CurrentDate     =   38776
         End
         Begin MSComCtl2.DTPicker dtpFechaLiquidacion 
            Height          =   315
            Left            =   5340
            TabIndex        =   6
            Top             =   240
            Width           =   1485
            _ExtentX        =   2619
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
            Format          =   175702017
            CurrentDate     =   38776
         End
         Begin MSComCtl2.DTPicker dtpFechaPago 
            Height          =   315
            Left            =   9270
            TabIndex        =   7
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
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
            Format          =   175702017
            CurrentDate     =   38776
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tasa (%)"
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
            Left            =   4200
            TabIndex        =   70
            Top             =   690
            Width           =   750
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Núm. Cuotas a Pagar"
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
            Left            =   11310
            TabIndex        =   67
            Top             =   300
            Visible         =   0   'False
            Width           =   1800
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Observación"
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
            Index           =   30
            Left            =   7200
            TabIndex        =   13
            Top             =   690
            Width           =   1065
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Pago"
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
            Index           =   86
            Left            =   7200
            TabIndex        =   12
            Top             =   300
            Width           =   450
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
            Height          =   195
            Index           =   3
            Left            =   360
            TabIndex        =   11
            Top             =   690
            Width           =   690
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Descripción"
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
            Index           =   7
            Left            =   360
            TabIndex        =   10
            Top             =   1050
            Width           =   1005
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Liquidación"
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
            Left            =   4200
            TabIndex        =   9
            Top             =   285
            Width           =   975
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Orden"
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
            TabIndex        =   8
            Top             =   300
            Width           =   525
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmCancelacionDescuentoContratos.frx":6450
         Height          =   5265
         Left            =   -74880
         OleObjectBlob   =   "frmCancelacionDescuentoContratos.frx":646A
         TabIndex        =   57
         Top             =   2430
         Width           =   14100
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   35
         Left            =   -67920
         TabIndex        =   58
         Top             =   5100
         Visible         =   0   'False
         Width           =   405
      End
   End
   Begin VB.Label lblMercadoNegociación 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mercado Negociación"
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
      Left            =   6960
      TabIndex        =   76
      Top             =   9210
      Visible         =   0   'False
      Width           =   1860
   End
End
Attribute VB_Name = "frmCancelacionDescuentoContratos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrFondo()             As String, arrFondoOrden()              As String
Dim arrTipoInstrumento()   As String, arrTipoInstrumentoOrden()    As String
Dim arrEstado()            As String, arrTipoOrden()               As String
Dim arrOperacion()         As String, arrNegociacion()             As String
Dim arrEmisor()            As String, arrMoneda()                  As String
Dim arrObligado()          As String, arrGestor()                  As String
Dim arrBaseAnual()         As String, arrTipoTasa()                As String
Dim arrOrigen()            As String, arrClaseInstrumento()        As String

Dim arrTitulo()            As String, arrSubClaseInstrumento()     As String
Dim arrConceptoCosto()     As String, arrFiador()                 As String
Dim arrLineaCliente()      As String
Dim arrLineaClienteLista() As String
Dim arrResponsablePago()   As String, arrResponsablePagoCancel()   As String
Dim arrViaCobranza()       As String
Dim strCodFondo            As String, strCodFondoOrden             As String
Dim strCodTipoInstrumento  As String, strCodTipoInstrumentoOrden   As String
Dim strCodEstado           As String, strCodTipoOrden              As String
Dim strCodOperacion        As String, strCodNegociacion            As String
Dim strCodEmisor           As String, strCodMoneda                 As String
Dim strCodObligado         As String, strCodGestor                 As String
Dim strCodBaseAnual        As String, strCodTipoTasa               As String
Dim strCodOrigen           As String, strCodClaseInstrumento       As String
Dim strCodTitulo           As String, strCodSubClaseInstrumento    As String
Dim strCodTituloOrigen     As String
Dim strCodConcepto         As String, strCodReportado              As String
Dim strCodGarantia         As String, strCodAgente                 As String
Dim strEstado              As String, strSQL                       As String
Dim strCodFiador           As String, strIndGarantia               As String
Dim strLineaCliente        As String
Dim strLineaClienteLista   As String
Dim strResponsablePago     As String, strResponsablePagoCancel     As String
Dim arrPagoInteres()       As String

Dim strCodFile             As String, strCodAnalitica              As String
Dim strCodAnaliticaOrig    As String
Dim strCodGrupo            As String, strCodCiiu                   As String
Dim strEstadoOrden         As String, strCodCategoria              As String
Dim strCodRiesgo           As String, strCodSubRiesgo              As String
Dim strCalcVcto            As String, strCodSector                 As String
Dim strCodTipoCostoBolsa   As String, strCodTipoCostoConasev       As String
Dim strCodTipoCostoFondo   As String, strCodTipoCavali             As String
Dim strIndCuponCero        As String, strIndPacto                  As String
Dim strIndNegociable       As String, strCodigosFile               As String
Dim strCodIndiceInicial    As String, strCodIndiceFinal            As String
Dim strCodTipoAjuste       As String, strCodPeriodoPago            As String
Dim strCodCobroInteres     As String, strViaCobranza               As String
Dim dblTipoCambio          As Double, dblTasaCuponNormal           As Double
Dim dblComisionBolsa       As Double, dblComisionConasev           As Double
Dim dblComisionFondo       As Double, dblComisionCavali            As Double
Dim intBaseCalculo         As Integer, dblFactorDiarioNormal       As Double
Dim SwCalculo              As Boolean

Dim indInicializaGrilla    As Boolean
Dim indCargaPantalla       As Boolean
Dim rsg                    As New ADODB.Recordset
Dim rsgVcto                As New ADODB.Recordset
Dim indInserta             As Boolean
Dim indActualizaFP         As Boolean
Dim dblSumaFPCnt           As Double
Dim dblSumaFPVto           As Double
Dim strIndCuentaCorriente, strIndCuentaAhorros As String
Dim dblBkpMontoFPago          As Double
Dim strCodMonedaParEvaluacion As String
Dim strCodMonedaParPorDefecto As String
Dim strNumAnexo               As String
Dim strCodLimiteSel           As String
Dim strCodEstructuraSel       As String
Dim strCodPersonaLim          As String
Dim strTipoPersonaLim         As String
Dim intDiasAdicionales        As Integer
Dim strCodComisionista        As String
Dim intSecuencialComisionista As Integer

Dim datFechaVctoAdicional     As Date
Dim intPlazoConAdic           As Integer
Dim blnCargadoDesdeCartera    As Boolean
Dim blnCargarCabeceraAnexo    As Boolean
Dim blnCancelaPrepago         As Boolean
Dim dblInteresesCorridosHOY   As Double
Dim intDiasTranscurridos      As Integer
Dim dblComisionOperacion      As Double          'Comisión que le corresponde a cada operación
Dim strCodMonedaComision      As String          'Moneda de expresión de la comisión
Dim strPersonalizaComision    As String
Dim dblPorcDescuento          As Double

Public Sub Adicionar()
    Dim strMsgError As String

    On Error GoTo err

    If cboTipoInstrumento.ListCount > 1 Then
        frmMainMdi.stbMdi.Panels(3).Text = "Adicionar Orden..."
        
        strEstado = Reg_Adicion
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        cmdAccion.Visible = True
        
        If blnCargarCabeceraAnexo = False Then Call HabilitaCombos(True)

        With tabRFCortoPlazo
            .TabEnabled(0) = False
            .TabEnabled(1) = True
            .Tab = 1
        End With
        
    Else
        MsgBox "Acceso a Negociación Denegada", vbCritical, Me.Caption
    End If
    
    Exit Sub
err:

    If strMsgError = "" Then strMsgError = err.Description
    MsgBox strMsgError, vbInformation, App.Title
End Sub

Private Sub LlenarFormulario(strModo As String)

    Dim adoRecord       As ADODB.Recordset
    Dim strSQL          As String
    Dim intRegistro     As Integer
    Dim strCambiarTCOpe As String
  
    Select Case strModo

        Case Reg_Adicion
        
            If blnCargarCabeceraAnexo = False Then  'si no he precargado datos

                chkTitulo.Value = vbUnchecked
                intRegistro = ObtenerItemLista(arrFondoOrden(), strCodFondo)

                If intRegistro >= 0 Then cboFondoOrden.ListIndex = intRegistro
            
                cboTipoInstrumentoOrden.ListIndex = -1

                If cboTipoInstrumentoOrden.ListCount > 0 Then cboTipoInstrumentoOrden.ListIndex = 0
                                        
                cboTipoOrden.ListIndex = -1

                If cboTipoOrden.ListCount > 0 Then cboTipoOrden.ListIndex = 0
        
                cboOperacion.ListIndex = -1

                If cboOperacion.ListCount > 0 Then cboOperacion.ListIndex = 0
                
                cboEmisor.ListIndex = -1

                If cboEmisor.ListCount > 0 Then cboEmisor.ListIndex = 0
                
                cboGestor.ListIndex = -1

                If cboGestor.ListCount > 0 Then cboGestor.ListIndex = 0
            
                intRegistro = ObtenerItemLista(arrOrigen(), Codigo_Negociacion_Local)

                If intRegistro >= 0 Then cboOrigen.ListIndex = intRegistro

            End If
           
            cboObligado.ListIndex = -1

            If cboObligado.ListCount > 0 Then cboObligado.ListIndex = 0

            intRegistro = ObtenerItemLista(arrMoneda(), strCodMoneda)

            If intRegistro >= 0 Then cboMoneda.ListIndex = intRegistro
            
            txtNumOperacionOrig.Text = ""
            dtpFechaOrden.Value = gdatFechaActual
            dtpFechaLiquidacion.Value = dtpFechaOrden.Value
            
            txtDescripOrden.Text = Valor_Caracter
            txtObservacion.Text = Valor_Caracter
            
            SwCalculo = False 'prepara la interfase para el ingreso de datos

    End Select
    
    'Obteniendo el parámetro que indica si se puede cambiar el TC en la operación
    Set adoRecord = New ADODB.Recordset
    
    adoComm.CommandText = "SELECT upper(ValorParametro) AS CambiarTCOpe FROM ParametroGeneral WHERE CodParametro = '21'"
    Set adoRecord = adoComm.Execute
 
    If Not (adoRecord.EOF) Then
        strCambiarTCOpe = Trim(adoRecord("CambiarTCOpe"))
    End If
    
    adoRecord.Close: Set adoRecord = Nothing
    
End Sub

Public Sub Cancelar()

    cmdOpcion.Visible = True
    cmdAccion.Visible = False

    With tabRFCortoPlazo
        .TabEnabled(0) = True
        .Tab = 0
    End With

    Call Buscar
    
End Sub

Public Sub Eliminar()

    If strEstado = Reg_Consulta Or strEstado = Reg_Edicion Then
        Dim strMensaje As String
        
        
        'verificar si la orden no está ya anulada
        
        If strCodEstado <> Estado_Orden_Anulada And strCodEstado <> Estado_Orden_Procesada Then
        
            strMensaje = "Se procederá a eliminar la ORDEN " & tdgConsulta.Columns(0) & " por la " & tdgConsulta.Columns(3) & vbNewLine & vbNewLine & vbNewLine & "¿ Seguro de continuar ?"
        
            If MsgBox(strMensaje, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
        
                '*** Anular Orden ***
                adoComm.CommandText = "UPDATE InversionOrden SET EstadoOrden='" & Estado_Orden_Anulada & "' WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & "CodTitulo='" & Trim(tdgConsulta.Columns(2)) & "' AND NumOrden='" & Trim(tdgConsulta.Columns(0)) & "'"
                    
                adoConn.Execute adoComm.CommandText
                
                '*** Anular Título si corresponde ***
                adoComm.CommandText = "UPDATE InstrumentoInversion SET IndVigente='' WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & "CodTitulo='" & Trim(tdgConsulta.Columns(2)) & "'"
                    
                adoConn.Execute adoComm.CommandText
                
                MsgBox Mensaje_Eliminacion_Exitosa, vbExclamation, Me.Caption
                
                tabRFCortoPlazo.TabEnabled(0) = True
                tabRFCortoPlazo.Tab = 0
                Call Buscar
                
                Exit Sub
            End If
            
        Else

            If strCodEstado = Estado_Orden_Anulada Then
                MsgBox "La orden " & Trim(tdgConsulta.Columns(0)) & " ya ha sido anulada.", vbExclamation, "Anular Orden"
            Else
                MsgBox "La orden " & Trim(tdgConsulta.Columns(0)) & " ya ha sido procesada." & vbNewLine & "No se puede anular.", vbCritical, "Anular Orden"
            End If
        End If

        

    End If
    
End Sub

Public Sub Grabar()

    Call Accion(vSave)

End Sub

Public Sub GrabarNew()

    Dim adoRegistro            As ADODB.Recordset
    Dim adoTemporal            As ADODB.Recordset
    Dim strCodTipoOrden        As String
    Dim strFechaOrden          As String, strFechaLiquidacion      As String
    Dim strFechaEmision        As String, strFechaVencimiento      As String
    Dim strFechaPago           As String
    Dim strFechaVctoDcto       As String
    Dim strFechaInteresAdic    As String
    Dim strMensaje             As String, strIndTitulo             As String
    Dim intRegistro            As Integer, intAccion               As Integer
    Dim lngNumError            As Long
    Dim dblTasaInteres         As Double
    
    Dim i                      As Integer
    Dim xmlDocFPIni            As DOMDocument60 'JCB
    Dim xmlDocCancelacion      As DOMDocument60 'JCB
    Dim strMsgError            As String 'JCB
    Dim intDiasAdicionalesVcto As Integer
    Dim strIndDevolucion       As String
    
    Dim dblSumaPrincipalCuota           As Double
    Dim dblSumaInteresCuota             As Double
    Dim dblSumaInteresAdicionalCuota    As Double
    Dim dblSumaIGVInteresCuota          As Double
    Dim dblSumaIGVInteresAdicionalCuota As Double
    Dim dblSumaTotalCuota               As Double
    
    On Error GoTo CtrlError
      
    If strEstado = Reg_Consulta Then Exit Sub
    
    If strEstado = Reg_Adicion Then
        If TodoOK() Then
        
            XMLDetalleGrid xmlDocCancelacion, "DetCancelacion", gDetalle, "NumCupon,FechaEmision,FechaVencimiento,PagoPrincipal,PagoIntereses,PagoIGVIntereses,PagoInteresAdicional,PagoIGVInteresAdicional", strMsgError  'JCB

            dblSumaPrincipalCuota = CDec(gDetalle.Columns.ColumnByFieldName("PagoPrincipal").SummaryFooterValue)
            dblSumaInteresCuota = CDec(gDetalle.Columns.ColumnByFieldName("PagoIntereses").SummaryFooterValue)
            dblSumaInteresAdicionalCuota = CDec(gDetalle.Columns.ColumnByFieldName("PagoInteresAdicional").SummaryFooterValue)
            dblSumaIGVInteresCuota = CDec(gDetalle.Columns.ColumnByFieldName("PagoIGVIntereses").SummaryFooterValue)
            dblSumaIGVInteresAdicionalCuota = CDec(gDetalle.Columns.ColumnByFieldName("PagoIGVInteresAdicional").SummaryFooterValue)
            
            dblSumaTotalCuota = dblSumaPrincipalCuota + dblSumaInteresCuota + dblSumaInteresAdicionalCuota + dblSumaIGVInteresCuota + dblSumaIGVInteresAdicionalCuota
            
            If strMsgError <> "" Then GoTo CtrlError
        
            strEstadoOrden = Estado_Orden_Ingresada

            strMensaje = "_____________________________________________________" & Chr(vbKeyReturn) & Chr(vbKeyReturn) & Space(8) & "<<<<<     " & Trim(UCase(cboFondoOrden.Text)) & "     >>>>>" & Chr(vbKeyReturn) & "_____________________________________________________" & Chr(vbKeyReturn) & Chr(vbKeyReturn) & "Para proceder al Registro de la Orden Confirme lo siquiente : " & Chr(vbKeyReturn) & Chr(vbKeyReturn) & "Fecha de Operación        " & Space(3) & ">" & Space(2) & CStr(dtpFechaOrden.Value) & Chr(vbKeyReturn) & "Fecha de Liquidación      " & Space(3) & ">" & Space(2) & CStr(dtpFechaLiquidacion.Value) & Chr(vbKeyReturn) & Chr(vbKeyReturn) & "Fecha de Pago             " & Space(3) & ">" & Space(2) & CStr(dtpFechaPago.Value) & Chr(vbKeyReturn) & Chr(vbKeyReturn) & Chr(vbKeyReturn) & Chr(vbKeyReturn) & "¿ Seguro de continuar ?"

            If MsgBox(strMensaje, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                Me.Refresh: Exit Sub
            End If
        
            Me.MousePointer = vbHourglass
            
            strFechaOrden = Convertyyyymmdd(dtpFechaOrden.Value)
            strFechaLiquidacion = Convertyyyymmdd(dtpFechaLiquidacion.Value)
            strFechaPago = Convertyyyymmdd(dtpFechaPago.Value)
           
            Set adoRegistro = New ADODB.Recordset
            
            '*** Guardar Orden de Cancelación ***
            With adoComm
                strIndTitulo = Valor_Caracter
                

                strIndTitulo = Valor_Indicador
              
                strCodTitulo = strCodTituloOrigen
                strCodGarantia = Valor_Caracter
               
                strCodReportado = Valor_Caracter
    
                
                .CommandText = "select FechaVencimiento from InversionOrden where CodFondo = '" & strCodFondoOrden & "' and CodAdministradora = '" & gstrCodAdministradora & "' and CodFile = '" & strCodFile & "' and CodAnalitica = '" & strCodAnaliticaOrig & "'"
                
                Set adoTemporal = .Execute
            
                If Not adoTemporal.EOF Then
                    strFechaVencimiento = Convertyyyymmdd(adoTemporal("FechaVencimiento"))
                    strFechaVctoDcto = strFechaVencimiento
                End If

                adoTemporal.Close
                
                Set adoTemporal = New ADODB.Recordset
                           
                .CommandText = "select ModoCobroInteres from InversionOrden where CodFondo = '" & strCodFondoOrden & "' and CodAdministradora = '" & gstrCodAdministradora & "' and CodFile = '" & strCodFile & "' and CodAnalitica = '" & strCodAnaliticaOrig & "'"
                
                Set adoTemporal = .Execute
        
                If Not adoTemporal.EOF Then
                    strCodCobroInteres = adoTemporal("ModoCobroInteres")
                End If

                If CDec(txtMontoRecibido.Text) < CDec(txtDeudaFecha.Text) Then
                    strCodTipoOrden = Codigo_Orden_Prepago
                Else
                    strCodTipoOrden = Codigo_Orden_PagoCancelacion
                End If

                adoTemporal.Close
                Dim dblAmortizacion As Double
                 dblAmortizacion = CDec(txtMontoRecibido.Text) - dblSumaInteresCuota - dblSumaInteresAdicionalCuota - dblSumaIGVInteresCuota - dblSumaIGVInteresAdicionalCuota
              
                .CommandText = "{ call up_IVAdicInversionOrden('" & strCodFondoOrden & "','" & gstrCodAdministradora & _
                                "','','" & strFechaOrden & "','" & strCodTitulo & "','" & Trim(txtDescripOrden.Text) & _
                                "','" & gstrPeriodoActual & "','" & gstrMesActual & "','','" & strEstadoOrden & "','" & _
                                strCodAnaliticaOrig & "','" & strCodFile & "','" & strCodAnaliticaOrig & "','" & strCodClaseInstrumento & "','" & _
                                strCodSubClaseInstrumento & "','" & strCodTipoOrden & "','" & strCodOperacion & "','" & _
                                Codigo_Negociacion_Local & "','" & strCodOrigen & "','" & Trim(txtDescripOrden.Text) & "','" & _
                                strCodEmisor & "','" & strCodAgente & "','" & strCodGarantia & "','" & strCodComisionista & "'," & intSecuencialComisionista & ",'" & strFechaPago & "','" & _
                                strFechaVencimiento & "','" & strFechaLiquidacion & "','" & strFechaEmision & "','" & _
                                strCodMoneda & "'," & dblSumaPrincipalCuota & ",'','" & strCodMoneda & "','" & strCodMoneda & "'," & _
                                dblAmortizacion & ",1,1," & dblSumaPrincipalCuota & ",100," & dblSumaPrincipalCuota & ",1,1," & _
                                dblAmortizacion & "," & dblSumaInteresCuota & ",0,0,0,0,0,0,0,0,0," & _
                                CDec(txtMontoRecibido.Value) & ",0,0,0,0,0,0,0,0,0,0,0,0,0,0," & dblSumaTotalCuota & _
                                ",0,'','','','','','" & strCodReportado & "','" & strCodEmisor & "','" & strCodEmisor & "','" & _
                                strCodEmisor & "','" & strCodGestor & "','" & strCodFiador & "',0,'','X','" & strIndTitulo & "','" & strCodTipoTasa & _
                                "','" & strCodBaseAnual & "'," & CDec(dblTasaInteres) & ",'05','X','07',''," & CDec(dblTasaInteres) & "," & CDec(dblTasaInteres) & _
                                "," & CDec(dblTasaInteres) & ",'" & strCodRiesgo & "','" & strCodSubRiesgo & "','" & Trim(txtObservacion.Text) & "','" & _
                                gstrLogin & "','" & gstrFechaActual & "','" & gstrLogin & "','" & gstrFechaActual & "','" & strCodTituloOrigen & "','" & _
                                strCodCobroInteres & "'," & dblSumaInteresCuota & ",0," & dblSumaInteresAdicionalCuota & ",0,0,'01'," & gdblTasaIgv * 100 & "," & _
                                dblSumaIGVInteresCuota & "," & dblSumaIGVInteresAdicionalCuota & "," & gdblTasaIgv * 100 & ",0,0,0,0,0,0,0,'','','','','" & strLineaCliente & _
                                "','" & Codigo_LimiteRE_Cliente & "','" & strCodPersonaLim & "','" & strTipoPersonaLim & "','" & strResponsablePagoCancel & _
                                "','" & strViaCobranza & "',0,0,0," & CDec(txtMontoRecibido.Value) & ") }"
                
                adoConn.Execute .CommandText
               
            End With
            
            MsgBox Mensaje_Adicion_Exitosa, vbExclamation
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
            cmdOpcion.Visible = True
            cmdAccion.Visible = False
            
            With tabRFCortoPlazo
                .TabEnabled(0) = True
                .Tab = 0
            End With

            Call Buscar
        End If
    End If

    Exit Sub
        
CtrlError:

    Me.MousePointer = vbDefault
   
    If Left(err.Description, 14) <> "Excede Limites" Then
        strMsgError = "Error " & Str(err.Number) & vbNewLine
        strMsgError = strMsgError & err.Description
        MsgBox strMsgError, vbCritical, "Error"
      
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

    Else
        strMsgError = strMsgError & err.Description
        MsgBox strMsgError, vbCritical, "Limites"
        
    End If
        
End Sub

Private Function TodoOK() As Boolean
        
    Dim adoRegistro   As ADODB.Recordset
    Dim strFechaDesde As String, strFechaHasta        As String
    
    TodoOK = False
          
    If cboTipoInstrumentoOrden.ListIndex <= 0 Then
        MsgBox "Debe seleccionar el Tipo de Instrumento de Corto Plazo.", vbCritical, Me.Caption

        If cboTipoInstrumentoOrden.Enabled Then cboTipoInstrumentoOrden.SetFocus
        Exit Function
    End If
    
    If cboClaseInstrumento.ListIndex <= 0 Then
        MsgBox "Debe seleccionar la Clase de Instrumento de Corto Plazo.", vbCritical, Me.Caption

        If cboClaseInstrumento.Enabled Then cboClaseInstrumento.SetFocus
        Exit Function
    End If
                          
    If chkTitulo.Value Then
        If cboTitulo.ListIndex <= 0 Then
            MsgBox "Debe seleccionar el Título.", vbCritical, Me.Caption

            If cboTitulo.Enabled Then cboTitulo.SetFocus
            Exit Function
        End If

    Else

        If cboEmisor.ListIndex <= 0 Then
            MsgBox "Debe seleccionar el Emisor.", vbCritical, Me.Caption

            If cboEmisor.Enabled Then cboEmisor.SetFocus
            Exit Function
        End If
        
        Set adoRegistro = New ADODB.Recordset
        
        strFechaDesde = Convertyyyymmdd(dtpFechaOrden.Value)
        strFechaHasta = Convertyyyymmdd(DateAdd("d", 1, dtpFechaOrden.Value))
    End If
        
    If cboLineaCliente.ListIndex < 0 Then
        MsgBox "Debe seleccionar la Línea a afectar.", vbCritical, Me.Caption

        If cboLineaCliente.Enabled Then cboLineaCliente.SetFocus
        Exit Function
    End If
        
    If Trim(txtDescripOrden.Text) = Valor_Caracter Then
        MsgBox "Debe indicar la Descripción de la ORDEN.", vbCritical, Me.Caption

        If txtDescripOrden.Enabled Then txtDescripOrden.SetFocus
        Exit Function
    End If
    
    If CVDate(dtpFechaOrden.Value) > CVDate(dtpFechaLiquidacion.Value) Then
        MsgBox "La Fecha de Liquidación debe ser mayor o igual a la Fecha de la ORDEN.", vbCritical, Me.Caption

        If dtpFechaLiquidacion.Enabled Then dtpFechaLiquidacion.SetFocus
        Exit Function
    End If
    
    '*** Si todo paso OK ***
    TodoOK = True
  
End Function

Public Sub SubImprimir(Index As Integer)

    Dim frmReporte As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    Dim strFechaDesde        As String, strFechaHasta        As String
    Dim strSeleccionRegistro As String

    If tabRFCortoPlazo.Tab = 1 Then Exit Sub
    
    Select Case Index

        Case 1
            gstrNameRepo = "InversionOrden"
            
            strSeleccionRegistro = "{InversionOrden.FechaOrden} IN 'Fch1' TO 'Fch2'"
            gstrSelFrml = strSeleccionRegistro
            frmRangoFecha.Show vbModal
                        
            If gstrSelFrml <> "0" Then
                Set frmReporte = New frmVisorReporte
    
                ReDim aReportParamS(5)
                ReDim aReportParamFn(5)
                ReDim aReportParamF(5)
                            
                aReportParamFn(0) = "Usuario"
                aReportParamFn(1) = "FechaDesde"
                aReportParamFn(2) = "FechaHasta"
                aReportParamFn(3) = "Hora"
                aReportParamFn(4) = "Fondo"
                aReportParamFn(5) = "NombreEmpresa"
                            
                aReportParamF(0) = gstrLogin
                aReportParamF(1) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10)
                aReportParamF(2) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10)
                aReportParamF(3) = Format(Time(), "hh:mm:ss")
                aReportParamF(4) = Trim(cboFondo.Text)
                aReportParamF(5) = gstrNombreEmpresa & Space(1)
                            
                aReportParamS(0) = strCodFondo
                aReportParamS(1) = gstrCodAdministradora
                aReportParamS(2) = Convertyyyymmdd(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10))
                aReportParamS(3) = Convertyyyymmdd(DateAdd("d", 1, CVDate(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10))))
                aReportParamS(4) = strCodMoneda
                aReportParamS(5) = strCodTipoInstrumento
            End If

        Case 2
            gstrNameRepo = "PapeletaInversion"
            
            strSeleccionRegistro = "{InversionOrden.FechaOrden} IN 'Fch1' TO 'Fch2'"
            gstrSelFrml = strSeleccionRegistro
            frmRangoFecha.Show vbModal
                        
            If gstrSelFrml <> "0" Then
                Set frmReporte = New frmVisorReporte
    
                ReDim aReportParamS(5)
                ReDim aReportParamFn(1)
                ReDim aReportParamF(1)
                            
                aReportParamFn(0) = "Fondo"
                aReportParamFn(1) = "NombreEmpresa"
                            
                aReportParamF(0) = Trim(cboFondo.Text)
                aReportParamF(1) = gstrNombreEmpresa & Space(1)
                            
                aReportParamS(0) = strCodFondo
                aReportParamS(1) = gstrCodAdministradora
                aReportParamS(2) = Convertyyyymmdd(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10))
                aReportParamS(3) = Convertyyyymmdd(DateAdd("d", 1, CVDate(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10))))
                aReportParamS(4) = strCodMoneda
                aReportParamS(5) = strCodTipoInstrumento
            End If
            
        Case 3, 4

            If Index = 3 Then
                gstrNameRepo = "Anexo"
            Else
                gstrNameRepo = "AnexoCliente"
            End If
            
            Set frmReporte = New frmVisorReporte
            
            ReDim aReportParamS(9)
            ReDim aReportParamFn(3)
            ReDim aReportParamF(3)
            
            aReportParamFn(0) = "Usuario"
            aReportParamFn(1) = "Hora"
            aReportParamFn(2) = "Fondo"
            aReportParamFn(3) = "NombreEmpresa"
            
            aReportParamF(0) = gstrLogin
            aReportParamF(1) = Format(Time(), "hh:mm:ss")
            aReportParamF(2) = Trim(cboFondo.Text)
            aReportParamF(3) = gstrNombreEmpresa & Space(1)
                        
            aReportParamS(0) = strCodFondo
            aReportParamS(1) = gstrCodAdministradora
            aReportParamS(2) = tdgConsulta.Columns(11)
            aReportParamS(3) = tdgConsulta.Columns(12)
            aReportParamS(4) = tdgConsulta.Columns(13)
            aReportParamS(5) = tdgConsulta.Columns(21)
            aReportParamS(6) = tdgConsulta.Columns(23)
            aReportParamS(7) = tdgConsulta.Columns(24)
            aReportParamS(8) = tdgConsulta.Columns(25)
            aReportParamS(9) = tdgConsulta.Columns(9)
            
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

Public Sub Salir()

    Unload Me
    
End Sub

Private Sub cboClaseInstrumento_Click()

    strCodClaseInstrumento = Valor_Caracter

    If cboClaseInstrumento.ListIndex < 0 Then Exit Sub
    
    strCodClaseInstrumento = Trim(arrClaseInstrumento(cboClaseInstrumento.ListIndex))
    
    'En caso de documentos cambiarios podría usarse más de una línea dependiendo del destino
    'de uso del capital. Por tanto se mostrarán las líneas para su elección
    If strCodTipoInstrumentoOrden = "015" Then   'Documentos cambiarios
        strSQL = "SELECT CodLimite CODIGO, DescripLimite DESCRIP FROM LimiteReglamentoEstructuraDetalle where CodEstructura = '" & Codigo_LimiteRE_Cliente & "'" & " and CodLimite in ('" & Linea_Descuento_Letras_Facturas & "','" & Linea_Financiamiento_Proveedores & "') and Estado  = '01' "
        CargarControlLista strSQL, cboLineaCliente, arrLineaCliente(), Sel_Defecto
        'Call cboLineaCliente_Click
        
    Else

        If strCodTipoInstrumentoOrden = "014" Then
            strSQL = "SELECT CodLimite CODIGO, DescripLimite DESCRIP FROM LimiteReglamentoEstructuraDetalle where CodEstructura = '" & Codigo_LimiteRE_Cliente & "'" & " and CodLimite = '" & Linea_Descuento_Letras_Facturas & "' and Estado  = '01' "
            CargarControlLista strSQL, cboLineaCliente, arrLineaCliente(), ""
            Call cboLineaCliente_Click       'Para obligar a que se seleccione el único elemento de la lista
        Else
        
            If strCodTipoInstrumentoOrden = "016" Or strCodTipoInstrumentoOrden = "021" Then    'Descuento de Flujos Dinerarios o Préstamos
                strSQL = "SELECT CodLimite CODIGO, DescripLimite DESCRIP FROM LimiteReglamentoEstructuraDetalle where CodEstructura = '" & Codigo_LimiteRE_Cliente & "'" & " and CodLimite = '" & Linea_Contrato_Flujo_Dinerario & "' and Estado  = '01' "
                CargarControlLista strSQL, cboLineaCliente, arrLineaCliente(), ""
                Call cboLineaCliente_Click   'Para obligar a que se seleccione el único elemento de la lista
            Else

                If strCodTipoInstrumentoOrden = "010" Then   'Letras por maquinarias
                    strSQL = "SELECT CodLimite CODIGO, DescripLimite DESCRIP FROM LimiteReglamentoEstructuraDetalle where CodEstructura = '" & Codigo_LimiteRE_Cliente & "'" & " and CodLimite = '" & Linea_Compra_Maquinarias & "' and Estado  = '01' "
                    CargarControlLista strSQL, cboLineaCliente, arrLineaCliente(), ""
                    Call cboLineaCliente_Click   'Para obligar a que se seleccione el único elemento de la lista
                    
                End If
            
            End If
        
        End If
    End If

    If cboLineaCliente.ListCount > 0 Then cboLineaCliente.ListIndex = 0
    'Fin ACC 31/03/2010
    
    '*** SubClase de Instrumento ***
    strSQL = "SELECT CodSubDetalleFile CODIGO,DescripSubDetalleFile DESCRIP FROM InversionSubDetalleFile WHERE " & "CodDetalleFile='" & strCodClaseInstrumento & "' AND CodFile='" & strCodTipoInstrumentoOrden & "' AND IndVigente='X' ORDER BY DescripSubDetalleFile"
        
    CargarControlLista strSQL, cboSubClaseInstrumento, arrSubClaseInstrumento(), Sel_Defecto
    
    If cboSubClaseInstrumento.ListCount > 1 Then
        cboSubClaseInstrumento.ListIndex = ObtenerItemLista(arrSubClaseInstrumento(), strCodClaseInstrumento)
    Else

        If cboSubClaseInstrumento.ListCount > 0 Then cboSubClaseInstrumento.ListIndex = 0
    End If
    
    cboSubClaseInstrumento.Enabled = True

    If strCodClaseInstrumento = "001" Then strCalcVcto = "V"   'tasa de interés
    If strCodClaseInstrumento = "002" Then strCalcVcto = "D"    'Al descuento

End Sub

Private Sub cboEmisor_Click()

    Dim adoRegistro As ADODB.Recordset
    
    strCodTitulo = Valor_Caracter: strCodGrupo = Valor_Caracter: strCodCiiu = Valor_Caracter
    strCodEmisor = Valor_Caracter: strCodAnalitica = Valor_Caracter
    
    If cboEmisor.ListIndex < 0 Then Exit Sub

    strCodEmisor = arrEmisor(cboEmisor.ListIndex)

    '*** Validar Limites ***
    If strCodTipoInstrumentoOrden = Valor_Caracter Then Exit Sub
    If Not PosicionLimites() Then Exit Sub

    'txtDescripOrden = Trim(cboTipoInstrumentoOrden.Text) & " - " & Trim(cboEmisor.Text)
    If blnCancelaPrepago = False Then
        strCodTitulo = strCodFondoOrden & strCodFile & strCodAnalitica
    End If
    
    With adoComm
        Set adoRegistro = New ADODB.Recordset
                        
        '*** Categoría del instrumento emitido por el emisor ***
        .CommandText = "SELECT CodCategoriaRiesgo,CodRiesgoFinal,CodSubRiesgoFinal FROM EmisionInstitucionPersona " & "WHERE CodEmisor='" & strCodEmisor & "' AND CodFile='" & strCodTipoInstrumentoOrden & "' AND " & "CodDetalleFile='" & strCodClaseInstrumento & "'"
            
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            strCodRiesgo = Trim(adoRegistro("CodRiesgoFinal"))
            strCodSubRiesgo = Trim(adoRegistro("CodSubRiesgoFinal"))
        Else

            If strCodEmisor <> Valor_Caracter Then
                'MsgBox "La Clasificación de Riesgo no está definida...", vbCritical, Me.Caption
                cboLineaCliente_Click
                Exit Sub
            End If
        End If

        adoRegistro.Close
        
        '*** Obtener el Riesgo ***
        .CommandText = "SELECT ValorParametro FROM AuxiliarParametro WHERE CodTipoParametro='TIPRIE' AND CodParametro='" & strCodRiesgo & "'"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            strCodCategoria = Trim(adoRegistro("ValorParametro"))
        End If

        adoRegistro.Close: Set adoRegistro = Nothing
    
    End With
    
    cboLineaCliente_Click
    
End Sub

Private Function PosicionLimites() As Boolean

    PosicionLimites = False
        
    If cboTipoInstrumentoOrden.ListIndex <= 0 Then
        MsgBox "Debe seleccionar el Tipo de Instrumento.", vbCritical, Me.Caption
        cboEmisor.ListIndex = -1: cboTitulo.ListIndex = -1

        If cboTipoInstrumentoOrden.Enabled Then cboTipoInstrumentoOrden.SetFocus
        Exit Function
    End If

    '    If strCodTipoOrden = Codigo_Orden_Compra Then ValidLimites strCodEmisor, Convertyyyymmdd(dtpFechaOrden.Value), CDbl(txtTipoCambio.Text), strCodFile, strCodFondoOrden

    '*** Si todo pasó OK ***
    PosicionLimites = True
    
End Function

Private Sub cboEstado_Click()

    strCodEstado = Valor_Caracter

    If cboEstado.ListIndex < 0 Then Exit Sub
    
    strCodEstado = Trim(arrEstado(cboEstado.ListIndex))
    
    Call Buscar
End Sub

Private Sub cboFondo_Click()

    Dim adoRegistro As ADODB.Recordset
    
    strCodFondo = Valor_Caracter

    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(arrFondo(cboFondo.ListIndex))
    
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Fecha Vigente, Moneda ***
        .CommandText = "{ call up_ACSelDatosParametro(23,'" & strCodFondo & "','" & gstrCodAdministradora & "','000') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            gdatFechaActual = CVDate(adoRegistro("FechaCuota"))
            dtpFechaOrdenDesde.Value = gdatFechaActual
            dtpFechaOrdenHasta.Value = dtpFechaOrdenDesde.Value
            strCodMoneda = Trim(adoRegistro("CodMoneda"))
                        
            frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
        End If

        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    '*** Tipo de Instrumento ***
    strSQL = "SELECT FIF.CodFile CODIGO,DescripFile DESCRIP " & "FROM FondoInversionFile FIF JOIN InversionFile IVF ON(IVF.CodFile=FIF.CodFile) " & "WHERE TipoValor='" & Codigo_Valor_RentaFija & "' AND TipoPlazo='" & Codigo_Valor_CortoPlazo & "' AND IndInstrumento='X' AND IndVigente='X' AND " & "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND FIF.CodFile = '" & CodFile_Descuento_Flujos_Dinerarios & "' ORDER BY DescripFile"
    CargarControlLista strSQL, cboTipoInstrumento, arrTipoInstrumento(), Sel_Todos
    
    If cboTipoInstrumento.ListCount > 0 Then cboTipoInstrumento.ListIndex = 1
        
End Sub

Private Sub cboFondoOrden_Click()

    Dim adoRegistro As ADODB.Recordset
    
    strCodFondoOrden = Valor_Caracter

    If cboFondoOrden.ListIndex < 0 Then Exit Sub
    
    strCodFondoOrden = Trim(arrFondoOrden(cboFondoOrden.ListIndex))

    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Fecha Vigente, Moneda, Tipo de Cambio ***
        .CommandText = "{ call up_ACSelDatosParametro(23,'" & strCodFondoOrden & "','" & gstrCodAdministradora & "','000') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            gdatFechaActual = CVDate(adoRegistro("FechaCuota"))
            gstrFechaActual = Convertyyyymmdd(gdatFechaActual)
            dtpFechaOrden.Value = gdatFechaActual
            dtpFechaLiquidacion.Value = dtpFechaOrden.Value
            dtpFechaPago.Value = dtpFechaOrden.Value
            strCodMoneda = Trim(adoRegistro("CodMoneda"))
            '            txtTipoCambio.Text = CStr(dblTipoCambio)
                        
            frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
            Me.Refresh
        End If

        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    '*** Tipo de Instrumento ***
    strSQL = "SELECT FIF.CodFile CODIGO,DescripFile DESCRIP " & "FROM FondoInversionFile FIF JOIN InversionFile IVF ON(IVF.CodFile=FIF.CodFile) " & "WHERE TipoValor='" & Codigo_Valor_RentaFija & "' AND TipoPlazo='" & Codigo_Valor_CortoPlazo & "' AND IndInstrumento='X' AND IndVigente='X' AND " & "CodFondo='" & strCodFondoOrden & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND FIF.CodFile = '" & CodFile_Descuento_Flujos_Dinerarios & "' ORDER BY DescripFile"
    CargarControlLista strSQL, cboTipoInstrumentoOrden, arrTipoInstrumentoOrden(), Sel_Defecto
    
    If cboTipoInstrumentoOrden.ListCount > 0 Then cboTipoInstrumentoOrden.ListIndex = 1
        
End Sub

Private Sub cboGestor_Click()
    
    strCodGestor = Valor_Caracter

    If cboGestor.ListIndex < 0 Then Exit Sub
    
    strCodGestor = Trim(arrGestor(cboGestor.ListIndex))

End Sub

Private Sub cboLineaCliente_Click()

    Dim adoRegistro As New ADODB.Recordset

    strLineaCliente = Valor_Caracter
    strTipoPersonaLim = Valor_Caracter
    strCodPersonaLim = Valor_Caracter
    
    If cboLineaCliente.ListIndex < 0 Then Exit Sub
    
    strLineaCliente = Trim(arrLineaCliente(cboLineaCliente.ListIndex))
    
    'Obteniendo algunos valores para Limites y generación de Anexo
    strTipoPersonaLim = Codigo_Tipo_Persona_Emisor

    If strLineaCliente = Linea_Financiamiento_Proveedores Then
        strCodPersonaLim = strCodObligado
    Else
        strCodPersonaLim = strCodEmisor
    End If

End Sub

Private Sub cboLineaClienteLista_Click()

    strLineaClienteLista = Valor_Caracter

    If cboLineaClienteLista.ListIndex < 0 Then Exit Sub
    
    strLineaClienteLista = Trim(arrLineaClienteLista(cboLineaClienteLista.ListIndex))

    Call Buscar

End Sub

Private Sub cboObligado_Click()

    strCodObligado = Valor_Caracter

    If cboObligado.ListIndex < 0 Then Exit Sub
    
    strCodObligado = Trim(arrObligado(cboObligado.ListIndex))

    cboLineaCliente_Click

End Sub

Private Sub cboSubClaseInstrumento_Click()

    Dim adoRegistro As ADODB.Recordset          'ACC 12/03/2010  Agregado

    strCodSubClaseInstrumento = Valor_Caracter

    If cboSubClaseInstrumento.ListIndex < 0 Then Exit Sub
    
    strCodSubClaseInstrumento = Trim(arrSubClaseInstrumento(cboSubClaseInstrumento.ListIndex))
    
    'ACC 16/03/2010  Agregado
    'Si se trata de letras cambiarias (con Tasa de Interés) permitir poder elegir el modo de pago de de los intereses (inicio o vencimiento)
    If strCodTipoInstrumentoOrden = "015" And strCodClaseInstrumento = "001" And strCodSubClaseInstrumento = "001" Then
        'txtIntAdicional(0).Enabled = True   'Habilitar la edición de intereses adicionales
    
        'ACC 12/03/2010  Agregado
     
        'Obteniendo los dìas adicionales
        Set adoRegistro = New ADODB.Recordset
        adoComm.CommandText = "SELECT CONVERT(int,ValorParametro) AS DiasAdicionales FROM ParametroGeneral WHERE CodParametro = '20'"
        Set adoRegistro = adoComm.Execute

        If Not (adoRegistro.EOF) Then
            intDiasAdicionales = adoRegistro("DiasAdicionales")
        End If

        If intDiasAdicionales = Null Then
            intDiasAdicionales = 0
        End If

        adoRegistro.Close: Set adoRegistro = Nothing
    
    End If
    
    Call cboTipoOrden_Click
    
End Sub

Private Sub cboOrigen_Click()

    strCodOrigen = Valor_Caracter

    If cboOrigen.ListIndex < 0 Then Exit Sub
    
    strCodOrigen = Trim(arrOrigen(cboOrigen.ListIndex))
    
End Sub

Private Sub cboTipoInstrumento_Click()

    strCodTipoInstrumento = Valor_Caracter

    If cboTipoInstrumento.ListIndex < 0 Then Exit Sub
    
    strCodTipoInstrumento = Trim(arrTipoInstrumento(cboTipoInstrumento.ListIndex))
    
    'En caso de documentos cambiarios podría usarse más de una línea dependiendo del destino
    'de uso del capital. Por tanto se mostrarán las líneas para su elección
    If strCodTipoInstrumento = "015" Then   'Documentos cambiarios
        strSQL = "SELECT CodLimite CODIGO, DescripLimite DESCRIP FROM LimiteReglamentoEstructuraDetalle where CodEstructura = '" & Codigo_LimiteRE_Cliente & "'" & " and CodLimite in ('" & Linea_Descuento_Letras_Facturas & "','" & Linea_Financiamiento_Proveedores & "') and Estado  = '01' "
        CargarControlLista strSQL, cboLineaClienteLista, arrLineaClienteLista(), Sel_Defecto
    Else

        If strCodTipoInstrumento = "014" Then
            strSQL = "SELECT CodLimite CODIGO, DescripLimite DESCRIP FROM LimiteReglamentoEstructuraDetalle where CodEstructura = '" & Codigo_LimiteRE_Cliente & "'" & " and CodLimite = '" & Linea_Descuento_Letras_Facturas & "' and Estado  = '01' "
            CargarControlLista strSQL, cboLineaClienteLista, arrLineaClienteLista(), ""
        Else

            If strCodTipoInstrumento = "016" Or strCodTipoInstrumento = "021" Then   'Descuento de Flujos Dinerarios o Préstamos
                strSQL = "SELECT CodLimite CODIGO, DescripLimite DESCRIP FROM LimiteReglamentoEstructuraDetalle where CodEstructura = '" & Codigo_LimiteRE_Cliente & "'" & " and CodLimite = '" & Linea_Contrato_Flujo_Dinerario & "' and Estado  = '01' "
                CargarControlLista strSQL, cboLineaClienteLista, arrLineaClienteLista(), ""
            Else

                If strCodTipoInstrumento = "010" Then   'Letras por maquinarias
                    strSQL = "SELECT CodLimite CODIGO, DescripLimite DESCRIP FROM LimiteReglamentoEstructuraDetalle where CodEstructura = '" & Codigo_LimiteRE_Cliente & "'" & " and CodLimite = '" & Linea_Compra_Maquinarias & "' and Estado  = '01' "
                    CargarControlLista strSQL, cboLineaClienteLista, arrLineaClienteLista(), ""
                End If
            
            End If
        
        End If
    End If

    If cboLineaClienteLista.ListCount > 0 Then cboLineaClienteLista.ListIndex = 0
    
    Call Buscar
    
End Sub

Private Sub cboTipoInstrumentoOrden_Click()
    
    Dim adoRegistro As ADODB.Recordset
    Dim strFecha    As String
    
    strCodTipoInstrumentoOrden = Valor_Caracter
    strIndPacto = Valor_Caracter: strIndNegociable = Valor_Caracter

    If cboTipoInstrumentoOrden.ListIndex < 0 Then Exit Sub
    
    strCodTipoInstrumentoOrden = Trim(arrTipoInstrumentoOrden(cboTipoInstrumentoOrden.ListIndex))

    strSQL = "SELECT IFTON.CodTipoOperacion CODIGO,DescripParametro DESCRIP " & "FROM InversionFileTipoOperacionNegociacion IFTON JOIN AuxiliarParametro AUX ON(AUX.CodParametro=IFTON.CodTipoOperacion AND AUX.CodTipoParametro = 'OPECAJ') " & "WHERE IFTON.CodFile='" & strCodTipoInstrumentoOrden & "' AND IFTON.CodTipoOperacion in ('10','26') ORDER BY CodTipoOperacion DESC"
    CargarControlLista strSQL, cboTipoOrden, arrTipoOrden(), Valor_Caracter

    If cboTipoOrden.ListCount > 0 Then
        cboTipoOrden.ListIndex = 0
    End If
    
    strCodFile = strCodTipoInstrumentoOrden

    '*** Clase de Instrumento ***
    strSQL = "SELECT CodDetalleFile CODIGO,DescripDetalleFile DESCRIP FROM InversionDetalleFile WHERE CodFile='" & strCodTipoInstrumentoOrden & "' AND IndVigente='X' ORDER BY DescripDetalleFile"
    CargarControlLista strSQL, cboClaseInstrumento, arrClaseInstrumento(), Sel_Defecto
    
    If cboClaseInstrumento.ListCount > 0 Then
        cboClaseInstrumento.ListIndex = 0
        cboClaseInstrumento.Enabled = True
    End If
    
    cboLineaCliente.Clear
            
End Sub

Private Sub cboMoneda_Click()
    
    If cboMoneda.ListIndex < 0 Then Exit Sub
    
    strCodMoneda = Trim(arrMoneda(cboMoneda.ListIndex))
    
End Sub

Private Sub cboOperacion_Click()

    strCodOperacion = Valor_Caracter

    If cboOperacion.ListIndex < 0 Then Exit Sub
    
    strCodOperacion = Trim(arrOperacion(cboOperacion.ListIndex))
    
End Sub

Private Sub cboTipoOrden_Click()

    Dim adoRegistro As ADODB.Recordset
    
    strCodTipoOrden = Valor_Caracter

    If cboTipoOrden.ListIndex < 0 Then Exit Sub

    strCodTipoOrden = Trim(arrTipoOrden(cboTipoOrden.ListIndex))
    blnCancelaPrepago = False
    
    strCodFile = strCodTipoInstrumentoOrden
    
    Me.MousePointer = vbDefault
    
End Sub

Private Sub cboTitulo_Click()

    Dim adoRegistro As ADODB.Recordset
    Dim intRegistro As Integer
    
    strCodGarantia = Valor_Caracter: txtDescripOrden = Valor_Caracter: strCodAnalitica = Valor_Caracter
    
    strCodEmisor = Valor_Caracter: strCodGrupo = Valor_Caracter

    If cboTitulo.ListIndex < 0 Then Exit Sub

    strCodGarantia = Trim(arrTitulo(cboTitulo.ListIndex))

    With adoComm
        Set adoRegistro = New ADODB.Recordset

        .CommandText = "SELECT CodAnalitica,ValorNominal,CodMoneda,CodEmisor,CodGrupo,FechaEmision,FechaVencimiento," & "TasaInteres,CodRiesgo,CodSubRiesgo,CodTipoTasa,BaseAnual,Nemotecnico " & "FROM InstrumentoInversion WHERE CodTitulo='" & strCodGarantia & "'"
        Set adoRegistro = .Execute

        If Not adoRegistro.EOF Then
            strCodAnalitica = Trim(adoRegistro("CodAnalitica"))
                        
            intRegistro = ObtenerItemLista(arrMoneda(), adoRegistro("CodMoneda"))

            If intRegistro >= 0 Then cboMoneda.ListIndex = intRegistro
                
            strCodEmisor = Trim(adoRegistro("CodEmisor")): strCodGrupo = Trim(adoRegistro("CodGrupo"))
            strCodRiesgo = Trim(adoRegistro("CodRiesgo"))
            strCodSubRiesgo = Trim(adoRegistro("CodSubRiesgo"))
            
            cboMoneda.Enabled = False
        End If

        adoRegistro.Close

        .CommandText = "SELECT FechaPago " & "FROM InstrumentoInversionCalendario WHERE CodTitulo='" & strCodGarantia & "'"
        Set adoRegistro = .Execute

        If Not adoRegistro.EOF Then
            dtpFechaPago.Value = adoRegistro("FechaPago")
            dtpFechaPago_Change
            dtpFechaPago.Enabled = False
        End If

        adoRegistro.Close
        
        '*** Obtener el Riesgo ***
        .CommandText = "SELECT ValorParametro FROM AuxiliarParametro WHERE CodTipoParametro='TIPRIE' AND CodParametro='" & strCodRiesgo & "'"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            strCodCategoria = Trim(adoRegistro("ValorParametro"))
        End If

        adoRegistro.Close
        
        '*** Validar Limites ***
        If Not PosicionLimites() Then Exit Sub

        adoRegistro.Close: Set adoRegistro = Nothing

    End With

    txtDescripOrden = Trim(cboTipoInstrumentoOrden.Text) & " - " & Left(cboTitulo.Text, 15)
        
End Sub

Private Sub cmdBuscarSolicitud_Click()

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
        
        frmBus.Caption = "Solicitudes pendientes de cancelación"
        .sSql = "{ call up_IVSelSolicitudPendienteCancelacion('" & strCodFondoOrden & "','" & gstrCodAdministradora & "','" & gstrFechaActual & "') }"
        
        .OutputColumns = "1"
        .HiddenColumns = ""
        
        .BuscarTabla
        
        Screen.MousePointer = vbNormal
        frmBus.Show 1
       
        If .iParams.Count = 0 Then Exit Sub
        
        If .iParams(1).Valor <> "" Then
            txtNumOperacionOrig.Text = .iParams(1).Valor
            txtNumOperacionOrig_KeyPress vbKeyReturn
        End If
       
    End With
    
    Set frmBus = Nothing

End Sub

Private Sub cmdEnviar_Click()

    Dim strFechaDesde As String, strFechaHasta        As String
    Dim intRegistro   As Integer, intContador         As Integer
    Dim datFecha      As Date
    
    If adoConsulta.Recordset.RecordCount = 0 Then Exit Sub
    
    strFechaDesde = Convertyyyymmdd(dtpFechaOrdenDesde.Value)
    datFecha = DateAdd("d", 1, dtpFechaOrdenHasta.Value)
    strFechaHasta = Convertyyyymmdd(datFecha)
    
    intContador = tdgConsulta.SelBookmarks.Count - 1
    
    If intContador < 0 Then
        MsgBox "No se ha seleccionado ningún registro", vbCritical, Me.Caption
        Exit Sub
    End If
        
    For intRegistro = 0 To intContador
        tdgConsulta.Row = tdgConsulta.SelBookmarks(intRegistro) - 1
               
        If strCodEstado = Estado_Orden_Ingresada Then
            adoComm.CommandText = "UPDATE InversionOrden SET EstadoOrden='" & Estado_Orden_Enviada & "'," & "UsuarioEdicion='" & gstrLogin & "',FechaEdicion='" & strFechaDesde & Space(1) & Format(Time, "hh:mm") & "' " & "WHERE NumOrden='" & Trim(tdgConsulta.Columns(0)) & "' AND CodFondo='" & strCodFondo & "' AND " & "CodAdministradora='" & gstrCodAdministradora & "' AND EstadoOrden='" & Estado_Orden_Ingresada & "'"
        ElseIf strCodEstado = Estado_Orden_Enviada Then
            adoComm.CommandText = "UPDATE InversionOrden SET EstadoOrden='" & Estado_Orden_Ingresada & "'," & "UsuarioEdicion='" & gstrLogin & "',FechaEdicion='" & strFechaDesde & Space(1) & Format(Time, "hh:mm") & "' " & "WHERE NumOrden='" & Trim(tdgConsulta.Columns(0)) & "' AND CodFondo='" & strCodFondo & "' AND " & "CodAdministradora='" & gstrCodAdministradora & "' AND EstadoOrden='" & Estado_Orden_Enviada & "'"
        End If

        adoConn.Execute adoComm.CommandText
    Next
    
    If strCodEstado = Estado_Orden_Ingresada Then
        MsgBox Mensaje_Envio_Exitoso, vbExclamation, gstrNombreEmpresa
    Else
        MsgBox Mensaje_Desenvio_Exitoso, vbExclamation, gstrNombreEmpresa
    End If

    Call Buscar
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub dtpFechaLiquidacion_Change()
        
    If Not EsDiaUtil(dtpFechaLiquidacion.Value) Then
        MsgBox "La Fecha no es un día útil...se cambiará por una fecha correcta !", vbInformation, Me.Caption
        dtpFechaLiquidacion.Value = ProximoDiaUtil(dtpFechaLiquidacion.Value)
    End If
    
    FormatoGrillaCancelacion "", 1, txtCuotasPago.Text
        
End Sub

Private Sub dtpFechaLiquidacionDesde_Click()

    If IsNull(dtpFechaLiquidacionDesde.Value) Then
        dtpFechaLiquidacionHasta.Value = Null
    Else
        dtpFechaLiquidacionDesde.Value = gdatFechaActual
        dtpFechaLiquidacionHasta.Value = dtpFechaLiquidacionDesde.Value
    End If
    
End Sub

Private Sub dtpFechaLiquidacionHasta_Click()

    If IsNull(dtpFechaLiquidacionHasta.Value) Then
        dtpFechaLiquidacionDesde.Value = Null
    Else
        dtpFechaLiquidacionDesde.Value = gdatFechaActual
        dtpFechaLiquidacionHasta.Value = dtpFechaLiquidacionDesde.Value
    End If
    
End Sub

Private Sub dtpFechaOrdenDesde_Click()

    If IsNull(dtpFechaOrdenDesde.Value) Then
        dtpFechaOrdenHasta.Value = Null
    Else
        dtpFechaOrdenDesde.Value = gdatFechaActual
        dtpFechaOrdenHasta.Value = dtpFechaOrdenDesde.Value
    End If
    
End Sub

Private Sub dtpFechaOrdenHasta_Click()

    If IsNull(dtpFechaOrdenHasta.Value) Then
        dtpFechaOrdenDesde.Value = Null
    Else
        dtpFechaOrdenDesde.Value = gdatFechaActual
        dtpFechaOrdenHasta.Value = dtpFechaOrdenDesde.Value
    End If
    
End Sub

Private Sub dtpFechaPago_Change()
    
    If Not EsDiaUtil(dtpFechaPago.Value) Then
        MsgBox "La Fecha de Pago no es un día útil...se cambiará por una fecha correcta !", vbInformation, Me.Caption
        dtpFechaPago.Value = ProximoDiaUtil(dtpFechaPago.Value)
    End If
    
End Sub

Private Sub Form_Activate()

    frmMainMdi.stbMdi.Panels(3).Text = Me.Caption
    Call CargarReportes
    
End Sub

Private Sub Form_Deactivate()

    Call OcultarReportes
    
End Sub

Private Sub Form_Load()

    indCargaPantalla = True
    blnCargadoDesdeCartera = False
    ConfGrid gDetalle, False, True, False, True
    
    Call InicializarValores
    Call CargarListas
    Call CargarReportes
    Call Buscar

    Call ValidarPermisoUsoControl(Trim(gstrLogin), Me, Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)

    CentrarForm Me
    indCargaPantalla = False
            
    Call ValidaExisteTipoCambio(Codigo_TipoCambio_SBS, gstrFechaActual)
    
End Sub

Public Sub Buscar()

    Dim strFechaOrdenDesde       As String, strFechaOrdenHasta        As String
    Dim strFechaLiquidacionDesde As String, strFechaLiquidacionHasta  As String
    Dim datFechaSiguiente        As Date

    Me.MousePointer = vbHourglass
    
    If Not IsNull(dtpFechaOrdenDesde.Value) And Not IsNull(dtpFechaOrdenHasta.Value) Then
        strFechaOrdenDesde = Convertyyyymmdd(dtpFechaOrdenDesde.Value)
        datFechaSiguiente = DateAdd("d", 1, dtpFechaOrdenHasta.Value)
        strFechaOrdenHasta = Convertyyyymmdd(datFechaSiguiente)
    End If
    
    If Not IsNull(dtpFechaLiquidacionDesde.Value) And Not IsNull(dtpFechaLiquidacionHasta.Value) Then
        strFechaLiquidacionDesde = Convertyyyymmdd(dtpFechaLiquidacionDesde.Value)
        datFechaSiguiente = DateAdd("d", 1, dtpFechaLiquidacionHasta.Value)
        strFechaLiquidacionHasta = Convertyyyymmdd(datFechaSiguiente)
    End If
    
    strSQL = "SELECT IOR.NumOrden,FechaOrden,FechaLiquidacion,CodTitulo,Nemotecnico,EstadoOrden,IOR.CodFile,CodAnalitica,TipoOrden,IOR.CodMoneda," & _
       "(RTRIM(DescripParametro) + SPACE(1) + DescripOrden) DescripOrden,CantOrden,ValorNominal,PrecioUnitarioMFL1,MontoTotalMFL1, " & _
       "CodSigno DescripMoneda, IOR.NumAnexo, NumDocumentoFisico,IOR.CodDetalleFile, IOR.CodSubDetalleFile, IOR.CodFondo, IOR.CodGirador, " & _
       "IP1.DescripPersona DesGirador, IOR.CodObligado, IP2.DescripPersona DesObligado, IOR.CodGestor, IP3.DescripPersona DesGestor, CodLimiteCli, DescripLimite DesLimiteCli, " & _
       "IOR.CodEstructura, IOR.CodPersonaLim, IOR.TipoPersonaLim " & _
       "FROM InversionOrden IOR JOIN AuxiliarParametro AUX ON(AUX.CodParametro=IOR.TipoOrden AND AUX.CodTipoParametro = 'OPECAJ') " & _
       "JOIN Moneda MON ON(MON.CodMoneda=IOR.CodMoneda) " & _
       "LEFT JOIN InstitucionPersona IP1 ON (IP1.CodPersona = IOR.CodGirador AND IP1.TipoPersona = '" & Codigo_Tipo_Persona_Emisor & "') " & _
       "LEFT JOIN InstitucionPersona IP2 ON (IP2.CodPersona = IOR.CodObligado AND IP2.TipoPersona = '" & Codigo_Tipo_Persona_Emisor & "') " & _
       "LEFT JOIN InstitucionPersona IP3 ON (IP3.CodPersona = IOR.CodGestor AND IP3.TipoPersona = '" & Codigo_Tipo_Persona_Emisor & "') " & _
       "LEFT JOIN LimiteReglamentoEstructuraDetalle LR ON (LR.CodLimite = IOR.CodLimiteCli AND LR.CodEstructura = IOR.CodEstructura ) " & _
       "WHERE IOR.CodAdministradora='" & gstrCodAdministradora & "' AND IOR.CodFondo='" & strCodFondo & "' "
        
    strSQL = strSQL & "AND IOR.TipoOrden in ('" & Codigo_Orden_Prepago & "','" & Codigo_Orden_PagoCancelacion & "') "
        
    If strCodTipoInstrumento <> Valor_Caracter Then
        strSQL = strSQL & "AND IOR.CodFile='" & strCodTipoInstrumento & "' "
    Else
        strSQL = strSQL & "AND IOR.CodFile IN " & strCodigosFile & " "
    End If

    If Not IsNull(dtpFechaOrdenDesde.Value) And Not IsNull(dtpFechaOrdenHasta.Value) Then
        strSQL = strSQL & "AND (FechaOrden >='" & strFechaOrdenDesde & "' AND FechaOrden <'" & strFechaOrdenHasta & "') "
    End If
    
    If Not IsNull(dtpFechaLiquidacionDesde.Value) And Not IsNull(dtpFechaLiquidacionHasta.Value) Then
        strSQL = strSQL & "AND (FechaLiquidacion >='" & strFechaLiquidacionDesde & "' AND FechaLiquidacion <'" & strFechaLiquidacionHasta & "') "
    End If
    
    If strCodEstado <> Valor_Caracter Then
        strSQL = strSQL & "AND EstadoOrden='" & strCodEstado & "' "
    End If
    
    If (strLineaClienteLista <> Valor_Caracter) And (cboTipoInstrumento.ListIndex > 0) Then
        strSQL = strSQL & "AND IOR.CodLimiteCli='" & strLineaClienteLista & "' AND IOR.CodEstructura ='" & Codigo_LimiteRE_Cliente & "' "
    End If
    
    strSQL = strSQL & "ORDER BY IOR.NumOrden"
    
    strEstado = Reg_Defecto

    With adoConsulta
        .ConnectionString = gstrConnectConsulta
        .RecordSource = strSQL
        .Refresh
    End With

    tdgConsulta.Refresh

    If adoConsulta.Recordset.RecordCount > 0 Then strEstado = Reg_Consulta

    Me.MousePointer = vbDefault
    
End Sub

Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Ordenes de Inversión"
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Text = "Papeleta de Inversión"
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo3").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo3").Text = "Anexo"
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo4").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo4").Text = "Anexo Cliente"
    
End Sub

Private Sub CargarListas()

    Dim intRegistro As Integer
    
     '*** Fondos ***
    strSQL = "SELECT CodFondo CODIGO,DescripFondo DESCRIP FROM Fondo WHERE CodAdministradora='" & gstrCodAdministradora & "' AND Estado='01' and CodFondo = '" & gstrCodFondoContable & "' ORDER BY DescripFondo"
    CargarControlLista strSQL, cboFondo, arrFondo(), Valor_Caracter
    CargarControlLista strSQL, cboFondoOrden, arrFondoOrden(), Valor_Caracter
        
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
            
    '*** Estados de la Orden ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='ESTORD' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboEstado, arrEstado(), Sel_Todos
    
    intRegistro = ObtenerItemLista(arrEstado(), Estado_Orden_Ingresada)

    If intRegistro >= 0 Then cboEstado.ListIndex = intRegistro
        
    '*** Tipo de Orden ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPORD' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboTipoOrden, arrTipoOrden(), Valor_Caracter
    
    '*** Emisor ***
    strSQL = "SELECT (CodPersona) CODIGO,DescripPersona DESCRIP FROM InstitucionPersona WHERE TipoPersona='" & Codigo_Tipo_Persona_Emisor & "' AND IndVigente='X' ORDER BY DescripPersona"
    CargarControlLista strSQL, cboEmisor, arrEmisor(), Sel_Defecto
    
    '*** Obligado ***
    strSQL = "SELECT (CodPersona) CODIGO,DescripPersona DESCRIP FROM InstitucionPersona WHERE TipoPersona='" & Codigo_Tipo_Persona_Emisor & "' AND IndVigente='X' ORDER BY DescripPersona"
    CargarControlLista strSQL, cboObligado, arrObligado(), Sel_Defecto

    '*** Gestor ***
    strSQL = "SELECT (CodPersona) CODIGO,DescripPersona DESCRIP FROM InstitucionPersona WHERE TipoPersona='" & Codigo_Tipo_Persona_Emisor & "' AND CodPersona = '00000001' ORDER BY DescripPersona"
    CargarControlLista strSQL, cboGestor, arrGestor(), Valor_Caracter
    
    If cboGestor.ListCount > 0 Then cboGestor.ListIndex = 0
     
    '*** Mercado de Negociación ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='MDONEG' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboOrigen, arrOrigen(), Valor_Caracter
            
    '*** Moneda ***
    strSQL = "{ call up_ACSelDatos(2) }"
    CargarControlLista strSQL, cboMoneda, arrMoneda(), Sel_Defecto
    
    '*** Tipo Liquidación Operación ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPLIQ' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboOperacion, arrOperacion(), Valor_Caracter
    
    '*** Carga Lista para el tab de formas de pago ***
    indCargaPantalla = True
    
End Sub

Public Sub Accion(nAccion As ButtonAction)
    
    Select Case nAccion
                
        Case vNew
            Call Adicionar

        Case vDelete
            Call Eliminar

        Case vSearch
            Call Buscar

        Case vReport

            'Call Imprimir
        Case vSave
            Call GrabarNew

        Case vCancel
            blnCancelaPrepago = False
            Call Cancelar

        Case vExit
            Call Salir
        
    End Select
    
End Sub

Private Sub InicializarValores()
    
    Dim adoRegistro As ADODB.Recordset
    
    '*** Valores Iniciales ***
    strEstado = Reg_Defecto
    tabRFCortoPlazo.Tab = 0

    SwCalculo = True 'indica cambio directo en la pantalla (false)

    dtpFechaOrdenDesde.Value = gdatFechaActual
    dtpFechaOrdenHasta.Value = dtpFechaOrdenDesde.Value
    dtpFechaPago.Value = dtpFechaOrdenDesde.Value
    dtpFechaLiquidacionDesde.Value = Null
    dtpFechaLiquidacionHasta.Value = dtpFechaLiquidacionDesde.Value
    
    strTipoPersonaLim = Valor_Caracter
    strCodPersonaLim = Valor_Caracter
    strCodComisionista = Valor_Caracter
    intSecuencialComisionista = 0
    strCodAnaliticaOrig = Valor_Caracter
    strCodTitulo = Valor_Caracter
    strResponsablePagoCancel = Valor_Caracter
    strViaCobranza = Valor_Caracter
   
    Set adoRegistro = New ADODB.Recordset

    With adoComm
        .CommandText = "SELECT CodFile FROM InversionFile " & "WHERE TipoValor='" & Codigo_Valor_RentaFija & "' AND TipoPlazo='" & Codigo_Valor_CortoPlazo & "' AND IndInstrumento='X' AND IndVigente='X' " & "ORDER BY DescripFile"
        Set adoRegistro = .Execute
                
        strCodigosFile = Valor_Caracter

        Do While Not adoRegistro.EOF

            If strCodigosFile <> Valor_Caracter Then strCodigosFile = strCodigosFile & ",'"
            
            strCodigosFile = strCodigosFile & Trim(adoRegistro("CodFile")) & "'"
        
            adoRegistro.MoveNext
        Loop

        adoRegistro.Close: Set adoRegistro = Nothing
                
        strCodigosFile = "('" & strCodigosFile & ",'009')"
    End With
    
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgConsulta.Columns(0).Width = tdgConsulta.Width * 0.01 * 8
    tdgConsulta.Columns(1).Width = tdgConsulta.Width * 0.01 * 10
    tdgConsulta.Columns(3).Width = tdgConsulta.Width * 0.01 * 20
    tdgConsulta.Columns(4).Width = tdgConsulta.Width * 0.01 * 4
    tdgConsulta.Columns(5).Width = tdgConsulta.Width * 0.01 * 5
    tdgConsulta.Columns(7).Width = tdgConsulta.Width * 0.01 * 6
    tdgConsulta.Columns(8).Width = tdgConsulta.Width * 0.01 * 9
    tdgConsulta.Columns(9).Width = tdgConsulta.Width * 0.01 * 8
    tdgConsulta.Columns(10).Width = tdgConsulta.Width * 0.01 * 7
    tdgConsulta.Columns(16).Width = tdgConsulta.Width * 0.01 * 8
    tdgConsulta.Columns(18).Width = tdgConsulta.Width * 0.01 * 8
    tdgConsulta.Columns(20).Width = tdgConsulta.Width * 0.01 * 8
    tdgConsulta.Columns(22).Width = tdgConsulta.Width * 0.01 * 15
    
    Set cmdOpcion.FormularioActivo = Me
  '  Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    
    'Leer si las comisiones van a ser definidas en la operación o ya viene establecida
    Set adoRegistro = New ADODB.Recordset
    adoComm.CommandText = "SELECT ValorParametro AS PersonalizaComi FROM ParametroGeneral WHERE CodParametro = '32'"
    Set adoRegistro = adoComm.Execute

    If Not (adoRegistro.EOF) Then
        strPersonalizaComision = Trim(adoRegistro("PersonalizaComi"))
    End If
    
    dblPorcDescuento = 100
        
    adoRegistro.Close: Set adoRegistro = Nothing
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmOrdenRentaFijaCortoPlazo = Nothing
    Call OcultarReportes
    frmMainMdi.stbMdi.Panels(3).Text = "Acción"
    
End Sub

Private Sub tabRFCortoPlazo_Click(PreviousTab As Integer)
            
    Dim dblMontoInicio As Double
    Dim dblMontoFin    As Double
    
    Select Case tabRFCortoPlazo.Tab

        Case 1, 2, 3

            If PreviousTab = 0 And blnCargadoDesdeCartera = False And strEstado = Reg_Consulta Then tabRFCortoPlazo.Tab = 0
            If strEstado = Reg_Defecto Then tabRFCortoPlazo.Tab = 0
            
            dblMontoInicio = txtMontoRecibido.Value
            dblMontoFin = txtMontoRecibido.Value

    End Select
            
End Sub

Private Sub tdgConsulta_FormatText(ByVal ColIndex As Integer, _
                                   Value As Variant, _
                                   Bookmark As Variant)

    If ColIndex = 4 Then
        Call DarFormatoValor(Value, Decimales_Monto)
    End If
    
    If ColIndex = 6 Then
        Call DarFormatoValor(Value, Decimales_Precio)
    End If
    
    If ColIndex = 7 Then
        Call DarFormatoValor(Value, Decimales_Monto)
    End If
    
End Sub

Private Sub txtCuotasPago_KeyPress(KeyAscii As Integer)
    Dim strMsgError As String

    On Error GoTo err

    If KeyAscii = 13 Then

        FormatoGrillaCancelacion strMsgError, 2, txtCuotasPago.Text

        If strMsgError <> "" Then GoTo err
    End If

    Exit Sub
err:

    If strMsgError = "" Then strMsgError = err.Description
    MsgBox strMsgError, vbInformation, App.Title
End Sub

Private Sub txtMontoRecibido_Change()
    txtSaldoDeuda.Text = Val(txtDeudaFecha.Value) - Val(txtMontoRecibido.Value)
End Sub

Private Sub txtMontoRecibido_KeyPress(KeyAscii As Integer)
    Dim strMsgError As String

    On Error GoTo err

    If KeyAscii = 13 Then CalculaPrelacion strMsgError
    If strMsgError <> "" Then GoTo err

    Exit Sub
err:

    If strMsgError = "" Then strMsgError = err.Description
    MsgBox strMsgError, vbInformation, App.Title
End Sub

Private Sub txtNumOperacionOrig_KeyPress(KeyAscii As Integer)
    Dim strMsgError As String

    On Error GoTo err

    If KeyAscii = vbKeyReturn Then
        txtNumOperacionOrig.Text = Format(txtNumOperacionOrig.Text, "0000000000")
        Call CargarSolicitud(strCodFondoOrden, gstrCodAdministradora, txtNumOperacionOrig.Text, 0)
        txtCuotasPago.Text = 1

        FormatoGrillaCancelacion strMsgError, 1, txtCuotasPago.Text

        If strMsgError <> "" Then GoTo err
    End If

    Exit Sub
err:

    If strMsgError = "" Then strMsgError = err.Description
    MsgBox strMsgError, vbInformation, App.Title
End Sub

Public Function CalculoInteresDescuento(numPorcenTasa As Double, _
                                        strCodTipoTasa As String, _
                                        strCodPeriodoTasa As String, _
                                        strCodBaseCalculo As String, _
                                        numMontoBaseCalculo As Double, _
                                        datFechaInicial As Date, _
                                        datFechaFinal As Date) As Double

    Dim intNumPeriodoAnualTasa As Integer
    Dim intDiasProvision       As Integer
    Dim intDiasBaseAnual       As Integer
    Dim numPorcenTasaAnual     As Double
    Dim numMontoCalculoInteres As Double
    Dim adoConsulta            As ADODB.Recordset
        
    With adoComm
        Set adoConsulta = New ADODB.Recordset
    
        '*** Obtener el número de días del periodo de tasa ***
        .CommandText = "SELECT ValorParametro FROM AuxiliarParametro WHERE CodTipoParametro='TIPFRE' AND CodParametro='" & strCodPeriodoTasa & "'"
        Set adoConsulta = .Execute
    
        If Not adoConsulta.EOF Then
            intNumPeriodoAnualTasa = CInt(360 / adoConsulta("ValorParametro"))     '*** Numero del periodos por año de la tasa ***
        End If

        adoConsulta.Close: Set adoConsulta = Nothing
    End With
   
    Select Case strCodBaseCalculo

        Case Codigo_Base_30_360:
            intDiasBaseAnual = 360
            intDiasProvision = DateDiff("d", datFechaInicial, datFechaFinal)       'Dias360(datFechaInicial, datFechaFinal, True)  ACC 21/04/2010

        Case Codigo_Base_Actual_365:
            intDiasBaseAnual = 365
            intDiasProvision = DateDiff("d", datFechaInicial, datFechaFinal) + 1

        Case Codigo_Base_Actual_360:
            intDiasBaseAnual = 360
            intDiasProvision = DateDiff("d", datFechaInicial, datFechaFinal) + 1

        Case Codigo_Base_30_365:
            intDiasBaseAnual = 365
            intDiasProvision = Dias360(datFechaInicial, datFechaFinal, True)
    End Select

    Select Case strCodTipoTasa

        Case Codigo_Tipo_Tasa_Efectiva:
            numPorcenTasaAnual = (1 + (numPorcenTasa / 100)) ^ (intNumPeriodoAnualTasa) - 1
            numMontoCalculoInteres = Round(numMontoBaseCalculo * ((((1 + numPorcenTasaAnual)) ^ (intDiasProvision / intDiasBaseAnual)) - 1), 2) 'adoRegistro("MontoDevengo") + curMontoRenta

        Case Codigo_Tipo_Tasa_Nominal:
            numPorcenTasaAnual = (numPorcenTasa / 100) * intNumPeriodoAnualTasa
            numMontoCalculoInteres = Round(numMontoBaseCalculo * ((numPorcenTasaAnual * (intDiasProvision / intDiasBaseAnual))), 2)

        Case Codigo_Tipo_Tasa_Flat:
            numPorcenTasaAnual = numPorcenTasa / 100
            numMontoCalculoInteres = Round(numMontoBaseCalculo * (numPorcenTasaAnual), 2)
    End Select

    CalculoInteresDescuento = numMontoCalculoInteres

End Function

Public Sub CargarSolicitud(strpCodFondoOrden As String, _
                           strpCodAdministradora As String, _
                           strpNumOperacionOrig As String, _
                           intpIndCartera As Integer)
    
    If intpIndCartera = 1 Then Form_Load
   
    'Obteniendo los datos de la operación de compra original
    Dim strSQL               As String
    Dim intRegistro          As Integer
    Dim strCodOperacionOrden As String
    Dim adoOperacionOrig     As ADODB.Recordset
    
    Set adoOperacionOrig = New ADODB.Recordset

    With adoComm
    
        .CommandText = "SELECT  CodFondo           ,CodAdministradora   ,NumSolicitud   ,FechaSolicitud    ,CodTitulo" & _
                        ",EstadoSolicitud   ,CodFile            ,CodAnalitica   ,CodDetalleFile    ,CodSubDetalleFile" & _
                        ",TipoSolicitud     ,DescripSolicitud   ,CodEmisor      ,CodComisionista   ,NumSecuencialComisionistaCondicion, " & _
                        "FechaConfirmacion  ,FechaVencimiento" & _
                        ",FechaLiquidacion  ,FechaEmision       ,CodMoneda      ,ValorTipoCambio   ,MontoSolicitud" & _
                        ",MontoAprobado     ,TipoTasa           ,BaseAnual      ,TasaInteres       ,Observacion " & _
                        ",MontoConsumido " & _
                        "FROM InversionSolicitud " & "WHERE CodFondo = '" & strpCodFondoOrden & "' AND CodAdministradora = '" & _
                                                    strpCodAdministradora & "' AND NumSolicitud='" & strpNumOperacionOrig & "'"
        Set adoOperacionOrig = .Execute

        If Not adoOperacionOrig.EOF Then

            'txtNum_Solicitud.Text = strNumSolicitud

            intRegistro = ObtenerItemLista(arrFondoOrden(), adoOperacionOrig.Fields("CodFondo"))

            If intRegistro >= 0 Then cboFondoOrden.ListIndex = intRegistro
            
            intRegistro = ObtenerItemLista(arrTipoInstrumentoOrden(), adoOperacionOrig.Fields("CodFile"))

            If intRegistro >= 0 Then cboTipoInstrumentoOrden.ListIndex = intRegistro
                                        
            intRegistro = ObtenerItemLista(arrClaseInstrumento(), adoOperacionOrig.Fields("CodDetalleFile"))

            If intRegistro >= 0 Then cboClaseInstrumento.ListIndex = intRegistro
               
            intRegistro = ObtenerItemLista(arrEmisor(), adoOperacionOrig.Fields("CodEmisor"))

            If intRegistro >= 0 Then cboEmisor.ListIndex = intRegistro
            If intRegistro >= 0 Then cboObligado.ListIndex = intRegistro
                
            intRegistro = ObtenerItemLista(arrSubClaseInstrumento(), adoOperacionOrig.Fields("CodSubDetalleFile"))

            If intRegistro >= 0 Then cboSubClaseInstrumento.ListIndex = intRegistro
                
            strCodAnaliticaOrig = adoOperacionOrig.Fields("CodAnalitica")
            strCodTituloOrigen = adoOperacionOrig.Fields("CodTitulo")
            strCodComisionista = adoOperacionOrig.Fields("CodComisionista")
            intSecuencialComisionista = adoOperacionOrig.Fields("NumSecuencialComisionistaCondicion")
            
            txtTasa.Text = adoOperacionOrig.Fields("TasaInteres")

            intRegistro = ObtenerItemLista(arrMoneda(), adoOperacionOrig.Fields("CodMoneda"))

            If intRegistro >= 0 Then cboMoneda.ListIndex = intRegistro
                
            txtDescripOrden.Text = adoOperacionOrig.Fields("DescripSolicitud")

            txtObservacion.Text = adoOperacionOrig.Fields("Observacion")
                    
        End If

        adoOperacionOrig.Close: Set adoOperacionOrig = Nothing

    End With
            
End Sub

Public Function ObtenerTotalDeuda(strpCodFondoOrden As String, _
                                  strpCodAdministradora As String, _
                                  strpNumOperacionOrig As String, _
                                  gstrpLogin As String) As Double

    Dim adoConsulta       As ADODB.Recordset
    Dim dblMontoDeuda     As Double
    Dim strFechaPagoCuota As String

    ObtenerTotalDeuda = 0

    strFechaPagoCuota = Convertyyyymmdd(dtpFechaOrden.Value)

    Set adoConsulta = New ADODB.Recordset
    
    With adoComm

        .CommandText = "{ call up_ACCalcularDeudaTotal ('" & strpCodFondoOrden & "','" & strpCodAdministradora & "','" & strpNumOperacionOrig & "','" & strFechaPagoCuota & "','" & strFechaPagoCuota & "','" & gstrpLogin & "' ) }"

        Set adoConsulta = .Execute

        If Not adoConsulta.EOF Then
            If adoConsulta("Deuda") <> "" Then
                dblMontoDeuda = CDbl(adoConsulta("Deuda"))
            Else
                dblMontoDeuda = 0
            End If
        End If
    
        adoConsulta.Close: Set adoConsulta = Nothing

    End With
    
    ObtenerTotalDeuda = dblMontoDeuda

End Function

Public Function ObtenerInteresesAdicionales(strpCodFondoOrden As String, _
                                            strpCodAdministradora As String, _
                                            strpNumOperacionOrig As String, _
                                            gstrpLogin As String) As Double

    Dim adoConsulta       As ADODB.Recordset
    Dim dblInteresesAdic  As Double
    Dim strFechaPagoCuota As String

    ObtenerInteresesAdicionales = 0

    strFechaPagoCuota = Convertyyyymmdd(dtpFechaOrden.Value)

    Set adoConsulta = New ADODB.Recordset
    
    With adoComm

        .CommandText = "{ call up_ACCalcularInteresesAdicionales ('" & strpCodFondoOrden & "','" & strpCodAdministradora & "','" & strpNumOperacionOrig & "','" & strFechaPagoCuota & "','" & strFechaPagoCuota & "','" & gstrpLogin & "' ) }"

        Set adoConsulta = .Execute

        If Not adoConsulta.EOF Then
            If adoConsulta("DeudaInteresAdicional") <> "" Then
                dblInteresesAdic = CDbl(adoConsulta("DeudaInteresAdicional"))
            Else
                dblInteresesAdic = 0
            End If
        End If
    
        adoConsulta.Close: Set adoConsulta = Nothing

    End With
    
    ObtenerInteresesAdicionales = dblInteresesAdic

End Function

Public Sub HabilitaCombos(ByVal pBloquea As Boolean)

    cboFondoOrden.Enabled = pBloquea
    cboTipoInstrumentoOrden.Enabled = pBloquea
    cboClaseInstrumento.Enabled = pBloquea
    cboSubClaseInstrumento.Enabled = pBloquea
    cboTipoOrden.Enabled = pBloquea
    cboTitulo.Enabled = pBloquea
    cboEmisor.Enabled = pBloquea

    If (strCodTipoOrden <> Codigo_Orden_Compra) And (strCodTipoOrden <> Codigo_Orden_Renovacion) Then
        cboObligado.Enabled = pBloquea
    End If

    cboGestor.Enabled = pBloquea
    cboOperacion.Enabled = pBloquea
    cboOrigen.Enabled = pBloquea
    cboLineaCliente.Enabled = pBloquea
    'txtNumAnexoOrig.Enabled = pBloquea

End Sub

Public Sub mostrarForm(ByVal strNumSolicitud As String)

    Load Me
    
    'indCargadoDesdeBandeja = True
    Adicionar
    
    txtNumOperacionOrig.Text = strNumSolicitud
    txtNumOperacionOrig_KeyPress 13
    
    '''    txtValorNominal.Text = txtMontoSolicitud.Value - txtMontoConsumido.Value
        
    Me.Show
End Sub

Private Sub FormatoGrillaCancelacion(ByRef strMsgError As String, _
                                     intTipoSel As Integer, _
                                     intCuotasPago As Integer) 'JCB
    Dim rsgrilla     As New ADODB.Recordset
    Dim rst          As New ADODB.Recordset

    Dim AcumulaCupon As Double

    On Error GoTo err
    '********FORMATO GRILLA***********
    rsgrilla.Fields.Append "NumSecuencial", adInteger, , adFldRowID
    rsgrilla.Fields.Append "NumCupon", adVarChar, 3, adFldIsNullable
    rsgrilla.Fields.Append "FechaEmision", adVarChar, 20, adFldIsNullable
    rsgrilla.Fields.Append "FechaVencimiento", adVarChar, 20, adFldIsNullable
    
    rsgrilla.Fields.Append "TotalCupon", adDouble, , adFldIsNullable
    
    rsgrilla.Fields.Append "Principal", adDouble, , adFldIsNullable
    rsgrilla.Fields.Append "Intereses", adDouble, , adFldIsNullable
    rsgrilla.Fields.Append "IGVIntereses", adDouble, , adFldIsNullable
    
    rsgrilla.Fields.Append "InteresAdicional", adDouble, , adFldIsNullable
    rsgrilla.Fields.Append "IGVInteresAdicional", adDouble, , adFldIsNullable
    
    rsgrilla.Fields.Append "PagoPrincipal", adDouble, , adFldIsNullable
    rsgrilla.Fields.Append "PagoIntereses", adDouble, , adFldIsNullable
    rsgrilla.Fields.Append "PagoIGVIntereses", adDouble, , adFldIsNullable
    
    rsgrilla.Fields.Append "PagoInteresAdicional", adDouble, , adFldIsNullable
    rsgrilla.Fields.Append "PagoIGVInteresAdicional", adDouble, , adFldIsNullable
    
    rsgrilla.Open
    
    'Set rst = DataProcedimiento("up_IVSelDatosSolicitudCuponera", strMsgError, strCodFondoOrden, gstrCodAdministradora, txtNumOperacionOrig.Text, dtpFechaLiquidacion.Value, intTipoSel, intCuotasPago)
'    Dim cm As New ADODB.Command
    Dim strSQL As String
'    cm.ActiveConnection = gstrConnectConsulta
'    cm.CommandType = adCmdText
    strSQL = "{call up_IVSelDatosSolicitudCuponera('" & strCodFondoOrden & "','" & gstrCodAdministradora & "','" & txtNumOperacionOrig.Text & "','" & Convertyyyymmdd(dtpFechaLiquidacion.Value) & "'," & intTipoSel & "," & intCuotasPago & ") }"
'    cm.Parameters.Append cm.CreateParameter("CodFondo", adVarChar, adParamInput, 3, strCodFondoOrden)
'    cm.Parameters.Append cm.CreateParameter("CodAdministradora", adVarChar, adParamInput, 3, gstrCodAdministradora)
'    cm.Parameters.Append cm.CreateParameter("NumSolicitud", adVarChar, adParamInput, 10, txtNumOperacionOrig.Text)
'    cm.Parameters.Append cm.CreateParameter("FecCancelacion", adDate, adParamInput, 20, dtpFechaLiquidacion.Value)
'    cm.Parameters.Append cm.CreateParameter("TipoSel", adInteger, adParamInput, 100, intTipoSel)
'    cm.Parameters.Append cm.CreateParameter("intCuotasPago", adInteger, adParamInput, 100, intCuotasPago)

    rst.Open strSQL, adoConn ', adOpenDynamic, adLockOptimistic
    
    If strMsgError <> "" Then GoTo err
    'rst.Open cm.CommandText, gstrConnectConsulta, adOpenDynamic, adLockOptimistic
    Do While Not rst.EOF

        rsgrilla.AddNew

        rsgrilla.Fields("NumSecuencial") = "" & rst.Fields("NumSecuencial")
        rsgrilla.Fields("NumCupon") = "" & rst.Fields("NumCuota")

        rsgrilla.Fields("FechaEmision") = "" & rst.Fields("FechaEmision")
        rsgrilla.Fields("FechaVencimiento") = "" & rst.Fields("FechaVencimiento")

        rsgrilla.Fields("Principal") = Round(CDbl("" & rst.Fields("SaldoPrincipal")), 2)

        rsgrilla.Fields("Intereses") = Round(CDbl("" & rst.Fields("SaldoIntereses")), 2)
        rsgrilla.Fields("IGVIntereses") = Round(CDbl("" & rst.Fields("SaldoIntereses") * gdblTasaIgv), 2)
        rsgrilla.Fields("InteresAdicional") = Round(CDbl("" & rst.Fields("InteresAdicional")), 2)
        rsgrilla.Fields("IGVInteresAdicional") = Round(CDbl("" & rst.Fields("InteresAdicional") * gdblTasaIgv), 2)

        rsgrilla.Fields("TotalCupon") = Round(CDbl("" & rst.Fields("TotalCupon")), 2)

        rsgrilla.Fields("PagoPrincipal") = 0
        rsgrilla.Fields("PagoIntereses") = 0
        rsgrilla.Fields("PagoIGVIntereses") = 0

        rsgrilla.Fields("PagoInteresAdicional") = 0
        rsgrilla.Fields("PagoIGVInteresAdicional") = 0

        rst.MoveNext

    Loop
    
    'Set gDetalle.DataSource = Nothing
    mostrarDatosGridSQL gDetalle, rsgrilla, strMsgError, "NumCupon"

    If strMsgError <> "" Then GoTo err
    
    txtDeudaFecha.Text = gDetalle.Columns.ColumnByFieldName("Principal").SummaryFooterValue + gDetalle.Columns.ColumnByFieldName("Intereses").SummaryFooterValue + gDetalle.Columns.ColumnByFieldName("IGVIntereses").SummaryFooterValue + gDetalle.Columns.ColumnByFieldName("InteresAdicional").SummaryFooterValue + gDetalle.Columns.ColumnByFieldName("IGVInteresAdicional").SummaryFooterValue
    txtMontoRecibido.Text = txtDeudaFecha.Value
    
    txtMontoRecibido_KeyPress 13
    Exit Sub
err:

    If strMsgError = "" Then strMsgError = err.Description
End Sub

Private Sub CalculaPrelacion(ByRef strMsgError As String)
    Dim dblMontoRestado As Double
    Dim intPrelacion    As Integer
    Dim strColDato      As String
    Dim strColPago      As String
    
    ' Variables para prelacion
    Dim dblInteresAdicPagado As Double
    Dim dblIGVInteresAdicPagado As Double
    Dim dblInteresPagado As Double
    Dim dblIGVInteresPagado As Double
    Dim dblPrincipalPagado As Double
    
    On Error GoTo err

    gDetalle.Dataset.First

    Do While Not gDetalle.Dataset.EOF
        gDetalle.Dataset.Edit
        gDetalle.Columns.ColumnByFieldName("PagoPrincipal").Value = 0
        gDetalle.Columns.ColumnByFieldName("PagoIntereses").Value = 0
        gDetalle.Columns.ColumnByFieldName("PagoIGVIntereses").Value = 0
        gDetalle.Columns.ColumnByFieldName("PagoInteresAdicional").Value = 0
        gDetalle.Columns.ColumnByFieldName("PagoIGVInteresAdicional").Value = 0
        gDetalle.Dataset.Post
    
        gDetalle.Dataset.Next
    Loop
                
    dblMontoRestado = txtMontoRecibido.Value

    intPrelacion = 0

    gDetalle.Dataset.First

    Do While Not gDetalle.Dataset.EOF And dblMontoRestado <> 0
'
'        Select Case intPrelacion
'
'            Case 0 'Intereses Adicionales
'                strColDato = "InteresAdicional"
'                strColPago = "PagoInteresAdicional"
'
'            Case 1 'IGV Intereses Adicionales
'                strColDato = "IGVInteresAdicional"
'                strColPago = "PagoIGVInteresAdicional"
'
'            Case 2 'Intereses
'                strColDato = "Intereses"
'                strColPago = "PagoIntereses"
'
'            Case 3 'IGV Intereses
'                strColDato = "IGVIntereses"
'                strColPago = "PagoIGVIntereses"
'
'            Case 4 'Principal
'                strColDato = "Principal"
'                strColPago = "PagoPrincipal"
'
'        End Select
        
        dblInteresAdicPagado = 0
        dblIGVInteresAdicPagado = 0
        dblInteresPagado = 0
        dblIGVInteresPagado = 0
        dblPrincipalPagado = 0

        If gDetalle.Columns.ColumnByFieldName("InteresAdicional").Value > 0 Then
            If dblMontoRestado >= gDetalle.Columns.ColumnByFieldName("InteresAdicional").Value + gDetalle.Columns.ColumnByFieldName("IGVInteresAdicional").Value Then
                dblInteresAdicPagado = gDetalle.Columns.ColumnByFieldName("InteresAdicional").Value
                dblIGVInteresAdicPagado = gDetalle.Columns.ColumnByFieldName("IGVInteresAdicional").Value
            Else
                dblInteresAdicPagado = dblMontoRestado / (1 + gdblTasaIgv)
                dblIGVInteresAdicPagado = dblMontoRestado - dblInteresAdicPagado
            End If
            
            dblMontoRestado = dblMontoRestado - dblInteresAdicPagado - dblIGVInteresAdicPagado
        
        End If
        
        If gDetalle.Columns.ColumnByFieldName("Intereses").Value > 0 Then
            If dblMontoRestado >= gDetalle.Columns.ColumnByFieldName("Intereses").Value + gDetalle.Columns.ColumnByFieldName("IGVIntereses").Value Then
                dblInteresPagado = gDetalle.Columns.ColumnByFieldName("Intereses").Value
                dblIGVInteresPagado = gDetalle.Columns.ColumnByFieldName("IGVIntereses").Value
            Else
                dblInteresPagado = dblMontoRestado / (1 + gdblTasaIgv)
                dblIGVInteresPagado = dblMontoRestado - dblInteresPagado
            End If
            
            dblMontoRestado = dblMontoRestado - dblInteresPagado - dblIGVInteresPagado
        
        End If
        
        If gDetalle.Columns.ColumnByFieldName("Principal").Value > 0 Then
            
            If dblMontoRestado >= CDec(gDetalle.Columns.ColumnByFieldName("Principal").Value) Then
                dblPrincipalPagado = CDec(gDetalle.Columns.ColumnByFieldName("Principal").Value)
            Else
                dblPrincipalPagado = dblMontoRestado
            End If
            dblMontoRestado = dblMontoRestado - dblPrincipalPagado
        End If
        
        gDetalle.Dataset.Edit
        gDetalle.Columns.ColumnByFieldName("PagoPrincipal").Value = dblPrincipalPagado
        gDetalle.Columns.ColumnByFieldName("PagoIntereses").Value = dblInteresPagado
        gDetalle.Columns.ColumnByFieldName("PagoIGVIntereses").Value = dblIGVInteresPagado
        gDetalle.Columns.ColumnByFieldName("PagoInteresAdicional").Value = dblInteresAdicPagado
        gDetalle.Columns.ColumnByFieldName("PagoIGVInteresAdicional").Value = dblIGVInteresAdicPagado
        
        gDetalle.Dataset.Post
        
        gDetalle.Dataset.Next
'        If gDetalle.Columns.ColumnByFieldName(strColDato).Value > 0 Then
'            gDetalle.Dataset.Edit
'
'            If dblMontoRestado >= CDec(gDetalle.Columns.ColumnByFieldName(strColDato).Value) Then
'                gDetalle.Columns.ColumnByFieldName(strColPago).Value = CDec(gDetalle.Columns.ColumnByFieldName(strColDato).Value)
'                dblMontoRestado = dblMontoRestado - CDec(gDetalle.Columns.ColumnByFieldName(strColDato).Value)
'            Else
'                gDetalle.Columns.ColumnByFieldName(strColPago).Value = dblMontoRestado
'                dblMontoRestado = 0
'            End If
'
'            gDetalle.Dataset.Post
'        End If
'
'        intPrelacion = intPrelacion + 1
'
'        If intPrelacion > 4 Then
'            intPrelacion = 0
'            gDetalle.Dataset.Next
'        End If

    Loop

    gDetalle.Dataset.First
    Exit Sub
err:

    If strMsgError = "" Then strMsgError = err.Description
End Sub

