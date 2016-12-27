VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{830D5A06-9B70-4F8C-98B6-7A19C4E7760B}#1.0#0"; "TAMControls.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmCancelacionPrestamos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancelacion de Préstamos - Operaciones de Financiamiento"
   ClientHeight    =   8400
   ClientLeft      =   7080
   ClientTop       =   3615
   ClientWidth     =   14415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   14415
   ShowInTaskbar   =   0   'False
   Begin TAMControls2.ucBotonEdicion2 cmdOpcion 
      Height          =   735
      Left            =   720
      TabIndex        =   53
      Top             =   7600
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
      TabIndex        =   52
      Top             =   7600
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
      Picture         =   "frmCancelacionPrestamos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   7600
      Width           =   1200
   End
   Begin MSAdodcLib.Adodc adoConsulta 
      Height          =   330
      Left            =   9390
      Top             =   7800
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
      Height          =   7515
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   14355
      _ExtentX        =   25321
      _ExtentY        =   13256
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
      TabPicture(0)   =   "frmCancelacionPrestamos.frx":0582
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraCriterio"
      Tab(0).Control(1)=   "tdgConsulta"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Datos Orden Inversión"
      TabPicture(1)   =   "frmCancelacionPrestamos.frx":059E
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
         TabIndex        =   35
         Top             =   3780
         Width           =   14085
         Begin DXDBGRIDLibCtl.dxDBGrid gDetalle 
            Height          =   2535
            Left            =   120
            OleObjectBlob   =   "frmCancelacionPrestamos.frx":05BA
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   270
            Width           =   13815
         End
         Begin TAMControls.TAMTextBox txtDeudaFecha 
            Height          =   315
            Left            =   11910
            TabIndex        =   39
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
            Container       =   "frmCancelacionPrestamos.frx":634C
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
            TabIndex        =   41
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
            Container       =   "frmCancelacionPrestamos.frx":6368
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
            TabIndex        =   43
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
            Container       =   "frmCancelacionPrestamos.frx":6384
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
            TabIndex        =   44
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
            TabIndex        =   42
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
            TabIndex        =   40
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
         Height          =   1875
         Left            =   120
         TabIndex        =   26
         Top             =   390
         Width           =   14085
         Begin VB.ComboBox cboFondoOrden 
            Height          =   315
            Left            =   1830
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   360
            Width           =   4725
         End
         Begin TAMControls.TAMTextBox txtInstrumento 
            Height          =   315
            Left            =   1830
            TabIndex        =   55
            Top             =   720
            Width           =   4695
            _ExtentX        =   8281
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
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Locked          =   -1  'True
            Container       =   "frmCancelacionPrestamos.frx":63A0
            Apariencia      =   1
            Borde           =   1
            MaximoValor     =   0
         End
         Begin VB.CommandButton cmdBuscarPrestamo 
            Caption         =   "..."
            Height          =   300
            Left            =   4260
            TabIndex        =   45
            Top             =   1440
            Width           =   315
         End
         Begin VB.TextBox txtNumOperacionOrig 
            Height          =   315
            Left            =   1830
            TabIndex        =   28
            Top             =   1440
            Width           =   2385
         End
         Begin TAMControls.TAMTextBox txtClaseInstrumento 
            Height          =   315
            Left            =   1830
            TabIndex        =   56
            Top             =   1080
            Width           =   4695
            _ExtentX        =   8281
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
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Locked          =   -1  'True
            Container       =   "frmCancelacionPrestamos.frx":63BC
            Apariencia      =   1
            Borde           =   1
            MaximoValor     =   0
         End
         Begin TAMControls.TAMTextBox txtAcreedor 
            Height          =   315
            Left            =   9300
            TabIndex        =   57
            Top             =   360
            Width           =   4455
            _ExtentX        =   7858
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
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Locked          =   -1  'True
            Container       =   "frmCancelacionPrestamos.frx":63D8
            Apariencia      =   1
            Borde           =   1
            MaximoValor     =   0
         End
         Begin TAMControls.TAMTextBox txtGestor 
            Height          =   315
            Left            =   9300
            TabIndex        =   58
            Top             =   720
            Width           =   4455
            _ExtentX        =   7858
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
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Locked          =   -1  'True
            Container       =   "frmCancelacionPrestamos.frx":63F4
            Apariencia      =   1
            Borde           =   1
            MaximoValor     =   0
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
            TabIndex        =   34
            Top             =   1520
            Width           =   1170
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
            TabIndex        =   33
            Top             =   800
            Width           =   570
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
            TabIndex        =   32
            Top             =   1160
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
            TabIndex        =   31
            Top             =   420
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
            TabIndex        =   30
            Top             =   800
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
            TabIndex        =   29
            Top             =   420
            Width           =   570
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
         Height          =   1575
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
            Picture         =   "frmCancelacionPrestamos.frx":6410
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "Enviar a BackOffice"
            Top             =   360
            Width           =   1200
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
            TabIndex        =   18
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
            Format          =   175505409
            CurrentDate     =   38785
         End
         Begin MSComCtl2.DTPicker dtpFechaOrdenHasta 
            Height          =   315
            Left            =   9600
            TabIndex        =   19
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
            Format          =   175505409
            CurrentDate     =   38785
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
            TabIndex        =   25
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
            TabIndex        =   24
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
            Left            =   8880
            TabIndex        =   23
            Top             =   800
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
            TabIndex        =   22
            Top             =   420
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
            TabIndex        =   21
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
            TabIndex        =   20
            Top             =   420
            Width           =   1110
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
         Top             =   2340
         Width           =   14085
         Begin VB.TextBox txtTasaMoratoria 
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
            Left            =   5340
            Locked          =   -1  'True
            MaxLength       =   45
            TabIndex        =   59
            Text            =   "0.0000"
            Top             =   630
            Width           =   1470
         End
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   5340
            Locked          =   -1  'True
            MaxLength       =   45
            TabIndex        =   48
            Text            =   "0.0000"
            Top             =   240
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
            TabIndex        =   47
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
            Format          =   175505409
            CurrentDate     =   38776
         End
         Begin MSComCtl2.DTPicker dtpFechaLiquidacion 
            Height          =   315
            Left            =   7770
            TabIndex        =   6
            Top             =   960
            Visible         =   0   'False
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
            Format          =   175505409
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
            Format          =   175505409
            CurrentDate     =   38776
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tasa Mora(%)"
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
            Left            =   4110
            TabIndex        =   60
            Top             =   690
            Width           =   1170
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
            Left            =   4110
            TabIndex        =   49
            Top             =   300
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
            TabIndex        =   46
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
            Left            =   6630
            TabIndex        =   9
            Top             =   1005
            Visible         =   0   'False
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
         Bindings        =   "frmCancelacionPrestamos.frx":696B
         Height          =   5235
         Left            =   -74880
         OleObjectBlob   =   "frmCancelacionPrestamos.frx":6985
         TabIndex        =   36
         Top             =   2100
         Width           =   14100
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   35
         Left            =   -67920
         TabIndex        =   37
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
      TabIndex        =   54
      Top             =   9210
      Visible         =   0   'False
      Width           =   1860
   End
End
Attribute VB_Name = "frmCancelacionPrestamos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrFondo()                  As String, arrFondoOrden()              As String
Dim arrTipoInstrumento()        As String
Dim arrEstado()                 As String
Dim arrTipoOrden()              As String
Dim arrMoneda()                 As String
Dim arrOrigen()                 As String
Dim arrLineaClienteLista()      As String
Dim strCodFondo                 As String
Dim strCodFondoOrden            As String
Dim strCodTipoInstrumento       As String
Dim strCodTipoInstrumentoOrden  As String
Dim strCodEstado                As String
Dim strCodTipoOrden             As String
Dim strCodMoneda                As String
Dim strCodGestor                As String
Dim strCodBaseAnual             As String
Dim strCodClaseInstrumento      As String
Dim strEstado                   As String
Dim strSQL                      As String
Dim strLineaClienteLista        As String
Dim strCodFile                  As String
Dim strCodAnalitica             As String
Dim strCodDetalleFile           As String
Dim strCodAcreedor              As String
Dim strEstadoOrden              As String
Dim strCodigosFile              As String
Dim rsg                         As New ADODB.Recordset
Dim rsgVcto                     As New ADODB.Recordset
Dim intCantDiasPlazo            As Integer
Dim strNemotecnico              As String
Dim dblTasaInteres              As Double
Dim dblTasaMoratoria            As Double
Dim strTipoTasa                 As String
Dim strPeriodoTasa              As String
Dim strPeriodoCapitalizacion    As String
Dim strTipoAmortizacion         As String
Dim strPeriodoCuota             As String
Dim intCantUnidadesPeriodo      As String
Dim indFinPeriodo               As String
Dim indIGV                      As String
Dim datFechaPrimerCorte         As Date
Dim strFechaOrden               As String
Dim strFechaLiquidacion         As String
Dim strFechaEmision             As String
Dim strFechaVencimiento         As String
Dim strFechaPago                As String
Dim indFechaAPartir             As String
Dim datFechaAPartir             As Date
Dim intDiasPagoMinimoInteres    As Integer
Dim strCodDesplazamientoCorte   As String
Dim strCodDesplazamientoPago    As String
Dim intCuotasPeriodoGracia      As Integer
Dim intCantTramos               As Integer
Dim strTipoTramo                As String
Dim dblValorNominalInicial      As Double

Dim blnCargadoDesdeCartera    As Boolean
Dim blnCargarCabeceraAnexo    As Boolean

Public Sub Adicionar()
    Dim strMsgError As String

    On Error GoTo err

    If cboTipoInstrumento.ListCount > 0 Then
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
    Dim intRegistro     As Integer
    Dim rsgrilla As ADODB.Recordset
  
    Select Case strModo

        Case Reg_Adicion
        
            If blnCargarCabeceraAnexo = False Then  'si no he precargado datos

                intRegistro = ObtenerItemLista(arrFondoOrden(), strCodFondo)

                If intRegistro >= 0 Then cboFondoOrden.ListIndex = intRegistro
            
                txtClaseInstrumento.Text = Valor_Caracter
                txtAcreedor.Text = Valor_Caracter
                txtGestor.Text = Valor_Caracter
                txtInstrumento.Text = Valor_Caracter
                
                txtTasa.Text = Valor_Caracter
                txtTasaMoratoria.Text = Valor_Caracter
                                                
                Set rsgrilla = New ADODB.Recordset
                
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
                rsgrilla.Fields.Append "DeudaTotal", adDouble, , adFldIsNullable
            
                rsgrilla.Fields.Append "PagoPrincipal", adDouble, , adFldIsNullable
                rsgrilla.Fields.Append "PagoIntereses", adDouble, , adFldIsNullable
                rsgrilla.Fields.Append "PagoIGVIntereses", adDouble, , adFldIsNullable
            
                rsgrilla.Fields.Append "PagoInteresAdicional", adDouble, , adFldIsNullable
                rsgrilla.Fields.Append "PagoIGVInteresAdicional", adDouble, , adFldIsNullable

                rsgrilla.Open
                
                mostrarDatosGridSQL gDetalle, rsgrilla, "", "NumCupon"
                   
            End If
           
            intRegistro = ObtenerItemLista(arrMoneda(), strCodMoneda)

            If intRegistro >= 0 Then cboMoneda.ListIndex = intRegistro
            
            txtNumOperacionOrig.Text = ""
            dtpFechaOrden.Value = gdatFechaActual
            dtpFechaLiquidacion.Value = dtpFechaOrden.Value
            
            txtDescripOrden.Text = Valor_Caracter
            txtObservacion.Text = Valor_Caracter
            txtAcreedor.Text = Valor_Caracter
            txtTasa.Text = "0"
            txtTasaMoratoria.Text = "0"
            txtGestor.Text = Valor_Caracter
            dblValorNominalInicial = 0
            txtInstrumento.Text = Valor_Caracter
            txtClaseInstrumento.Text = Valor_Caracter
            txtSaldoDeuda.Text = "0"
            txtMontoRecibido.Text = "0"
            txtDeudaFecha.Text = "0"
            txtNumOperacionOrig.Text = Valor_Caracter
            
            
    End Select
    
    'Obteniendo el parámetro que indica si se puede cambiar el TC en la operación
    Set adoRecord = New ADODB.Recordset
    
    adoComm.CommandText = "SELECT upper(ValorParametro) AS CambiarTCOpe FROM ParametroGeneral WHERE CodParametro = '21'"
    Set adoRecord = adoComm.Execute
    
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
                adoComm.CommandText = "UPDATE FinanciamientoOrden SET EstadoOrden='" & Estado_Orden_Anulada & "' WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'  AND NumOrden='" & Trim(tdgConsulta.Columns(0)) & "'"
                    
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
    Dim strMensaje             As String ', strIndTitulo             As String
    Dim intAccion               As Integer
    Dim lngNumError            As Long
 
    'Dim i                      As Integer
    'Dim xmlDocFPIni            As DOMDocument60 'JCB
    Dim xmlDocCancelacion      As DOMDocument60 'JCB
    Dim strMsgError            As String 'JCB
    'Dim intDiasAdicionalesVcto As Integer
    'Dim strIndDevolucion       As String
    
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

            dblSumaPrincipalCuota = Round(CDec(gDetalle.Columns.ColumnByFieldName("PagoPrincipal").SummaryFooterValue), 2)
            dblSumaInteresCuota = Round(CDec(gDetalle.Columns.ColumnByFieldName("PagoIntereses").SummaryFooterValue), 2)
            dblSumaInteresAdicionalCuota = Round(CDec(gDetalle.Columns.ColumnByFieldName("PagoInteresAdicional").SummaryFooterValue), 2)
            dblSumaIGVInteresCuota = Round(CDec(gDetalle.Columns.ColumnByFieldName("PagoIGVIntereses").SummaryFooterValue), 2)
            dblSumaIGVInteresAdicionalCuota = Round(CDec(gDetalle.Columns.ColumnByFieldName("PagoIGVInteresAdicional").SummaryFooterValue), 2)
            
            dblSumaTotalCuota = Round(dblSumaPrincipalCuota + dblSumaInteresCuota + dblSumaInteresAdicionalCuota + dblSumaIGVInteresCuota + dblSumaIGVInteresAdicionalCuota, 2)
            
            If strMsgError <> "" Then GoTo CtrlError
        
            strEstadoOrden = Estado_Orden_Ingresada

            strMensaje = "_____________________________________________________" & Chr(vbKeyReturn) & Chr(vbKeyReturn) & Space(8) & "<<<<<     " & Trim(UCase(cboFondoOrden.Text)) & "     >>>>>" & Chr(vbKeyReturn) & "_____________________________________________________" & Chr(vbKeyReturn) & Chr(vbKeyReturn) & "Para proceder al Registro de la Orden Confirme lo siquiente : " & Chr(vbKeyReturn) & Chr(vbKeyReturn) & "Fecha de Operación        " & Space(3) & ">" & Space(2) & CStr(dtpFechaOrden.Value) & Chr(vbKeyReturn) & "Fecha de Liquidación      " & Space(3) & ">" & Space(2) & CStr(dtpFechaLiquidacion.Value) & Chr(vbKeyReturn) & Chr(vbKeyReturn) & "Fecha de Pago             " & Space(3) & ">" & Space(2) & CStr(dtpFechaPago.Value) & Chr(vbKeyReturn) & Chr(vbKeyReturn) & "Monto del Pago             " & Space(3) & ">" & Space(2) & CStr(txtMontoRecibido.Text) & Chr(vbKeyReturn) & Chr(vbKeyReturn) & Chr(vbKeyReturn) & Chr(vbKeyReturn) & "¿ Seguro de continuar ?"

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
                
                .CommandText = "select FechaVencimiento from FinanciamientoOrden where CodFondo = '" & strCodFondoOrden & "' and CodAdministradora = '" & gstrCodAdministradora & "' and CodFile = '" & strCodFile & "' and CodAnalitica = '" & strCodAnalitica & "'"
                
                Set adoTemporal = .Execute
            
                If Not adoTemporal.EOF Then
                    strFechaVencimiento = Convertyyyymmdd(adoTemporal("FechaVencimiento"))
                End If

                adoTemporal.Close
                
                Set adoTemporal = New ADODB.Recordset
                           
                If CDec(txtMontoRecibido.Text) < CDec(txtDeudaFecha.Text) Then
                    strCodTipoOrden = Codigo_Orden_Prepago
                Else
                    strCodTipoOrden = Codigo_Orden_PagoCancelacion
                End If

           
                Dim dblAmortizacion As Double
                 dblAmortizacion = CDec(txtMontoRecibido.Text) - dblSumaInteresCuota - dblSumaInteresAdicionalCuota - dblSumaIGVInteresCuota - dblSumaIGVInteresAdicionalCuota
              
                .CommandText = "{ call up_FIAdicFinanciamientoOrden('" & strCodFondoOrden & "','" & gstrCodAdministradora & "','','" & strFechaOrden & "','" & strEstadoOrden & "','" & _
                                strCodTipoOrden & "','" & gstrPeriodoActual & "','" & gstrMesActual & "','','" & strCodFile & "','" & _
                                strCodAnalitica & "','" & strCodDetalleFile & "','" & Trim(txtDescripOrden.Text) & "','" & strNemotecnico & "','" & _
                                strCodAcreedor & "','" & strCodGestor & "','" & strFechaEmision & "','" & strFechaPago & "','" & strFechaLiquidacion & "','" & _
                                strFechaVencimiento & "'," & intCantDiasPlazo & ",'','','','','','" & strCodMoneda & "','" & strCodMoneda & "'," & dblAmortizacion & ",1,1," & _
                                dblValorNominalInicial & ",100," & dblSumaPrincipalCuota & "," & CDec(dblTasaInteres) & "," & CDec(dblTasaMoratoria) & ",'" & _
                                strTipoTasa & "','" & strPeriodoTasa & "','" & strPeriodoCapitalizacion & "','" & strCodBaseAnual & "','" & strTipoAmortizacion & "','" & _
                                strPeriodoCuota & "'," & intCantUnidadesPeriodo & ",'" & indFinPeriodo & "','" & indIGV & "','" & Convertyyyymmdd(datFechaPrimerCorte) & "','" & _
                                indFechaAPartir & "','" & Convertyyyymmdd(datFechaAPartir) & "'," & intDiasPagoMinimoInteres & ",'" & strCodDesplazamientoCorte & "','" & _
                                strCodDesplazamientoPago & "'," & intCuotasPeriodoGracia & "," & intCantTramos & ",'" & strTipoTramo & "'," & _
                                dblSumaInteresCuota & "," & gdblTasaIgv * 100 & "," & dblSumaIGVInteresCuota & ",0,0,0,0," & gdblTasaIgv * 100 & "," & _
                                "0," & dblSumaTotalCuota & "," & CDec(txtMontoRecibido.Value) & ",0," & dblSumaInteresAdicionalCuota & "," & _
                                dblSumaIGVInteresAdicionalCuota & ",0,0," & dblSumaTotalCuota & ",'" & Trim(txtObservacion.Text) & "','" & gstrLogin & "','" & gstrFechaActual & "') }"
                
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
        
    'Dim adoRegistro   As ADODB.Recordset
    'Dim strFechaDesde As String, strFechaHasta        As String
    
    TodoOK = False
          
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
    'Dim strFechaDesde        As String', strFechaHasta        As String
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
    strSQL = "SELECT CodFile CODIGO,DescripFile DESCRIP " & "FROM InversionFile " & "WHERE IndVigente='X' AND CodFile = '" & CodFile_Financiamiento_Prestamos & "' ORDER BY CODIGO"
    CargarControlLista strSQL, cboTipoInstrumento, arrTipoInstrumento(), Valor_Caracter
    If cboTipoInstrumento.ListCount > 0 Then cboTipoInstrumento.ListIndex = 0
        
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
        
End Sub

Private Sub cboTipoInstrumento_Click()

    strCodTipoInstrumento = Valor_Caracter

    If cboTipoInstrumento.ListIndex < 0 Then Exit Sub
    
    strCodTipoInstrumento = Trim(arrTipoInstrumento(cboTipoInstrumento.ListIndex))
    
    Call Buscar
    
End Sub

Private Sub cboMoneda_Click()
    
    If cboMoneda.ListIndex < 0 Then Exit Sub
    
    strCodMoneda = Trim(arrMoneda(cboMoneda.ListIndex))
    
End Sub

Private Sub cmdBuscarPrestamo_Click()

    'Dim sSql As String
   
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
        
        frmBus.Caption = "Préstamos pendientes de cancelación"
        .sSql = "{ call up_FISelPrestamoPendienteCancelacion('" & strCodFondoOrden & "','" & gstrCodAdministradora & "') }"
        
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

    Dim strFechaDesde As String
    Dim intRegistro   As Integer, intContador         As Integer
    Dim datFecha      As Date
    
    If adoConsulta.Recordset.RecordCount = 0 Then Exit Sub
    
    strFechaDesde = Convertyyyymmdd(dtpFechaOrdenDesde.Value)
    datFecha = DateAdd("d", 1, dtpFechaOrdenHasta.Value)
    
    intContador = tdgConsulta.SelBookmarks.Count - 1
    
    If intContador < 0 Then
        MsgBox "No se ha seleccionado ningún registro", vbCritical, Me.Caption
        Exit Sub
    End If
        
    For intRegistro = 0 To intContador
        tdgConsulta.Row = tdgConsulta.SelBookmarks(intRegistro) - 1
               
        If strCodEstado = Estado_Orden_Ingresada Then
            adoComm.CommandText = "UPDATE FinanciamientoOrden SET EstadoOrden='" & Estado_Orden_Enviada & "'," & "UsuarioEdicion='" & gstrLogin & "',FechaEdicion='" & strFechaDesde & Space(1) & Format(Time, "hh:mm") & "' " & "WHERE NumOrden='" & Trim(tdgConsulta.Columns(0)) & "' AND CodFondo='" & strCodFondo & "' AND " & "CodAdministradora='" & gstrCodAdministradora & "' AND EstadoOrden='" & Estado_Orden_Ingresada & "'"
        ElseIf strCodEstado = Estado_Orden_Enviada Then
            adoComm.CommandText = "UPDATE FinanciamientoOrden SET EstadoOrden='" & Estado_Orden_Ingresada & "'," & "UsuarioEdicion='" & gstrLogin & "',FechaEdicion='" & strFechaDesde & Space(1) & Format(Time, "hh:mm") & "' " & "WHERE NumOrden='" & Trim(tdgConsulta.Columns(0)) & "' AND CodFondo='" & strCodFondo & "' AND " & "CodAdministradora='" & gstrCodAdministradora & "' AND EstadoOrden='" & Estado_Orden_Enviada & "'"
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
    
    FormatoGrillaCancelacion "", 1
        
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

    blnCargadoDesdeCartera = False
    ConfGrid gDetalle, False, True, False, True
    
    Call InicializarValores
    Call CargarListas
    Call CargarReportes
    Call Buscar

    Call ValidarPermisoUsoControl(Trim(gstrLogin), Me, Trim(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)

    CentrarForm Me
            
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
    
    strSQL = "SELECT FO.NumOrden,FechaOrden,FechaLiquidacion,'' CodTitulo,Nemotecnico,EstadoOrden,FO.CodFile,CodAnalitica,TipoOrden,FO.CodMoneda," & _
       "(RTRIM(DescripParametro) + SPACE(1) + DescripOrden) DescripOrden,CantOrden,ValorNominal,1 PrecioUnitarioMFL1,MontoLiquidado MontoTotalMFL1, " & _
       "CodSigno DescripMoneda,FO.CodDetalleFile, FO.CodFondo, FO.CodAcreedor, " & _
       "IP1.DescripPersona DesAcreedor,  FO.CodGestor, IP3.DescripPersona DesGestor " & _
       "FROM FinanciamientoOrden FO JOIN AuxiliarParametro AUX ON(AUX.CodParametro=FO.TipoOrden AND AUX.CodTipoParametro = 'OPECAJ') " & _
       "JOIN Moneda MON ON(MON.CodMoneda=FO.CodMoneda) " & _
       "LEFT JOIN InstitucionPersona IP1 ON (IP1.CodPersona = FO.CodAcreedor AND IP1.TipoPersona = '" & Codigo_Tipo_Persona_Emisor & "') " & _
       "LEFT JOIN InstitucionPersona IP3 ON (IP3.CodPersona = FO.CodGestor AND IP3.TipoPersona = '" & Codigo_Tipo_Persona_Emisor & "') " & _
       "WHERE FO.CodAdministradora='" & gstrCodAdministradora & "' AND FO.CodFondo='" & strCodFondo & "' "
        
    strSQL = strSQL & "AND FO.TipoOrden in ('" & Codigo_Orden_Prepago & "','" & Codigo_Orden_PagoCancelacion & "') "
        
    If strCodTipoInstrumento <> Valor_Caracter Then
        strSQL = strSQL & "AND FO.CodFile='" & strCodTipoInstrumento & "' "
    Else
        strSQL = strSQL & "AND FO.CodFile IN " & strCodigosFile & " "
    End If

    If Not IsNull(dtpFechaOrdenDesde.Value) And Not IsNull(dtpFechaOrdenHasta.Value) Then
        strSQL = strSQL & "AND (FechaOrden >='" & strFechaOrdenDesde & "' AND FechaOrden <'" & strFechaOrdenHasta & "') "
    End If
    
    If strCodEstado <> Valor_Caracter Then
        strSQL = strSQL & "AND EstadoOrden='" & strCodEstado & "' "
    End If
    
    strSQL = strSQL & "ORDER BY FO.NumOrden"
    
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
        
    '*** Moneda ***
    strSQL = "{ call up_ACSelDatos(2) }"
    CargarControlLista strSQL, cboMoneda, arrMoneda(), Sel_Defecto
    
    '*** Carga Lista para el tab de formas de pago ***
    
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

    dtpFechaOrdenDesde.Value = gdatFechaActual
    dtpFechaOrdenHasta.Value = dtpFechaOrdenDesde.Value
    dtpFechaPago.Value = dtpFechaOrdenDesde.Value
    
    strCodAnalitica = Valor_Caracter
   
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
    tdgConsulta.Columns(2).Width = tdgConsulta.Width * 0.01 * 30
    tdgConsulta.Columns(3).Width = tdgConsulta.Width * 0.01 * 4
    tdgConsulta.Columns(4).Width = tdgConsulta.Width * 0.01 * 6
    tdgConsulta.Columns(5).Width = tdgConsulta.Width * 0.01 * 6
    tdgConsulta.Columns(7).Width = tdgConsulta.Width * 0.01 * 6
    tdgConsulta.Columns(12).Width = tdgConsulta.Width * 0.01 * 30
'    tdgConsulta.Columns(9).Width = tdgConsulta.Width * 0.01 * 8
'    tdgConsulta.Columns(10).Width = tdgConsulta.Width * 0.01 * 7
'    tdgConsulta.Columns(16).Width = tdgConsulta.Width * 0.01 * 15
    
    Set cmdOpcion.FormularioActivo = Me
  '  Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    
    'Leer si las comisiones van a ser definidas en la operación o ya viene establecida
    Set adoRegistro = New ADODB.Recordset
    adoComm.CommandText = "SELECT ValorParametro AS PersonalizaComi FROM ParametroGeneral WHERE CodParametro = '32'"
    Set adoRegistro = adoComm.Execute
        
    adoRegistro.Close: Set adoRegistro = Nothing
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmOrdenRentaFijaCortoPlazo = Nothing
    Call OcultarReportes
    frmMainMdi.stbMdi.Panels(3).Text = "Acción"
    
End Sub

Private Sub tabRFCortoPlazo_Click(PreviousTab As Integer)
            
    Dim dblMontoInicio As Double
    'Dim dblMontoFin    As Double
    
    Select Case tabRFCortoPlazo.Tab

        Case 1, 2, 3

            If PreviousTab = 0 And blnCargadoDesdeCartera = False And strEstado = Reg_Consulta Then tabRFCortoPlazo.Tab = 0
            If strEstado = Reg_Defecto Then tabRFCortoPlazo.Tab = 0
            
            dblMontoInicio = txtMontoRecibido.Value

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

        FormatoGrillaCancelacion strMsgError, 2

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

        FormatoGrillaCancelacion strMsgError, 1

        If strMsgError <> "" Then GoTo err
    End If

    Exit Sub
err:

    If strMsgError = "" Then strMsgError = err.Description
    MsgBox strMsgError, vbInformation, App.Title
End Sub

Public Sub CargarSolicitud(strpCodFondoOrden As String, _
                           strpCodAdministradora As String, _
                           strpNumOperacionOrig As String, _
                           intpIndCartera As Integer)
    
    If intpIndCartera = 1 Then Form_Load
   
    'Obteniendo los datos de la operación de compra original

    Dim intRegistro          As Integer
    Dim adoOperacionOrig     As ADODB.Recordset
    
    Set adoOperacionOrig = New ADODB.Recordset

    With adoComm
    
        .CommandText = "SELECT CodFondo, CodAdministradora, NumOperacion, FechaOperacion, " & _
                        "EstadoOperacion,FO.CodFile, IFL.DescripFile,CodAnalitica,FO.CodDetalleFile, IDF.DescripDetalleFile , " & _
                        "TipoOperacion,DescripOperacion,CodAcreedor, IP1.DescripPersona as DescripAcreedor, " & _
                        "CodGestor, IP2.DescripPersona as DescripGestor,FechaConfirmacion, " & _
                        "FechaVencimiento, FechaLiquidacion, FechaEmision, CodMoneda, " & _
                        "ValorTipoCambio, ValorNominal, ValorNominalDscto, TipoTasa, " & _
                        "BaseAnual, TasaInteres, TasaMoratoria, Observacion, Nemotecnico, PeriodoTasa, PeriodoCapitalizacion, TipoAmortizacion, " & _
                        "PeriodoCuota, CantUnidadesPeriodo, IndFinPeriodo, IndIGV, FechaPrimerCorte, IndFechaAPartir, FechaAPartir, DiasPagoMinimoInteres, " & _
                        "CodDesplazamientoCorte, CodDesplazamientoPago, CuotasPeriodoGracia, CantTramos, TipoTramo, CantDiasPlazo " & _
                        "FROM FinanciamientoOperacion FO " & _
                        "join InversionFile IFL on (FO.CodFile = IFL.CodFile) " & _
                        "join InversionDetalleFile IDF on (FO.CodFile = IDF.CodFile and FO.CodDetalleFile = IDF.CodDetalleFile ) " & _
                        "join InstitucionPersona IP1 on (FO.CodAcreedor = IP1.CodPersona and IP1.TipoPersona = '02') " & _
                        "join InstitucionPersona IP2 on (FO.CodGestor = IP2.CodPersona and IP2.TipoPersona = '02') " & _
                        "WHERE CodFondo = '" & strpCodFondoOrden & _
                        "' AND CodAdministradora = '" & strpCodAdministradora & "' AND NumOperacion ='" & strpNumOperacionOrig & "'"
        Set adoOperacionOrig = .Execute

        If Not adoOperacionOrig.EOF Then

            intRegistro = ObtenerItemLista(arrFondoOrden(), adoOperacionOrig.Fields("CodFondo"))

            If intRegistro >= 0 Then cboFondoOrden.ListIndex = intRegistro
            
            txtInstrumento.Text = Trim$(adoOperacionOrig.Fields("DescripFile"))
            strCodFile = Trim$(adoOperacionOrig.Fields("CodFile"))
           
            txtClaseInstrumento.Text = Trim$(adoOperacionOrig.Fields("DescripDetalleFile"))
            strCodDetalleFile = Trim$(adoOperacionOrig.Fields("CodDetalleFile"))

            txtAcreedor.Text = Trim$(adoOperacionOrig.Fields("DescripAcreedor"))
            strCodAcreedor = Trim$(adoOperacionOrig.Fields("CodAcreedor"))
            
            txtGestor.Text = Trim$(adoOperacionOrig.Fields("DescripGestor"))
            strCodGestor = Trim$(adoOperacionOrig.Fields("CodGestor"))
                
            strCodAnalitica = Trim$(adoOperacionOrig.Fields("CodAnalitica"))
            dblValorNominalInicial = adoOperacionOrig.Fields("ValorNominal")
            
            txtTasa.Text = adoOperacionOrig.Fields("TasaInteres")
            dblTasaInteres = adoOperacionOrig.Fields("TasaInteres")
            txtTasaMoratoria.Text = adoOperacionOrig.Fields("TasaMoratoria")
            dblTasaMoratoria = adoOperacionOrig.Fields("TasaMoratoria")
            
            strCodBaseAnual = Trim$(adoOperacionOrig.Fields("BaseAnual"))
            strTipoTasa = Trim$(adoOperacionOrig.Fields("TipoTasa"))
            strPeriodoTasa = Trim$(adoOperacionOrig.Fields("PeriodoTasa"))
            strPeriodoCapitalizacion = Trim$(adoOperacionOrig.Fields("PeriodoCapitalizacion"))
            strTipoAmortizacion = Trim$(adoOperacionOrig.Fields("TipoAmortizacion"))
            strPeriodoCuota = Trim$(adoOperacionOrig.Fields("PeriodoCuota"))
            intCantUnidadesPeriodo = adoOperacionOrig.Fields("CantUnidadesPeriodo")
            indFinPeriodo = Trim$(adoOperacionOrig.Fields("IndFinPeriodo"))
            indIGV = Trim$(adoOperacionOrig.Fields("IndIGV"))
            datFechaPrimerCorte = adoOperacionOrig.Fields("FechaPrimerCorte")
            indFechaAPartir = Trim$(adoOperacionOrig.Fields("IndFechaAPartir"))
            datFechaAPartir = adoOperacionOrig.Fields("FechaAPartir")
            intDiasPagoMinimoInteres = adoOperacionOrig.Fields("DiasPagoMinimoInteres")
            strCodDesplazamientoCorte = Trim$(adoOperacionOrig.Fields("CodDesplazamientoCorte"))
            strCodDesplazamientoPago = Trim$(adoOperacionOrig.Fields("CodDesplazamientoPago"))
            intCuotasPeriodoGracia = adoOperacionOrig.Fields("CuotasPeriodoGracia")
            intCantTramos = adoOperacionOrig.Fields("CantTramos")
            strTipoTramo = Trim$(adoOperacionOrig.Fields("TipoTramo"))
            intCantDiasPlazo = adoOperacionOrig.Fields("CantDiasPlazo")
            strFechaEmision = Convertyyyymmdd(adoOperacionOrig.Fields("FechaEmision"))

            intRegistro = ObtenerItemLista(arrMoneda(), adoOperacionOrig.Fields("CodMoneda"))
            If intRegistro >= 0 Then cboMoneda.ListIndex = intRegistro
                
            txtDescripOrden.Text = Trim$(adoOperacionOrig.Fields("DescripOperacion"))
            strNemotecnico = Trim$(adoOperacionOrig.Fields("Nemotecnico"))
            txtObservacion.Text = Trim$(adoOperacionOrig.Fields("Observacion"))
                    
        
        End If

        adoOperacionOrig.Close: Set adoOperacionOrig = Nothing

    End With
            
End Sub

Public Sub HabilitaCombos(ByVal pBloquea As Boolean)

    cboFondoOrden.Enabled = pBloquea

End Sub

Public Sub mostrarForm(ByVal strNumSolicitud As String)

    Load Me
    
    Adicionar
    
    txtNumOperacionOrig.Text = strNumSolicitud
    txtNumOperacionOrig_KeyPress 13
        
    Me.Show
End Sub

Private Sub FormatoGrillaCancelacion(ByRef strMsgError As String, intTipoSel As Integer)
    Dim rsgrilla     As New ADODB.Recordset
    Dim rst          As New ADODB.Recordset

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
    rsgrilla.Fields.Append "DeudaTotal", adDouble, , adFldIsNullable
    
    rsgrilla.Fields.Append "PagoPrincipal", adDouble, , adFldIsNullable
    rsgrilla.Fields.Append "PagoIntereses", adDouble, , adFldIsNullable
    rsgrilla.Fields.Append "PagoIGVIntereses", adDouble, , adFldIsNullable
    
    rsgrilla.Fields.Append "PagoInteresAdicional", adDouble, , adFldIsNullable
    rsgrilla.Fields.Append "PagoIGVInteresAdicional", adDouble, , adFldIsNullable
    
    rsgrilla.Open
    
    Dim strSQL As String
    strSQL = "{call up_FIObtenerSaldosOperacionPrestamo('" & strCodFondoOrden & "','" & gstrCodAdministradora & "','" & txtNumOperacionOrig.Text & "','" & Convertyyyymmdd(dtpFechaLiquidacion.Value) & "'," & intTipoSel & ") }"
    rst.Open strSQL, adoConn
    
    If strMsgError <> "" Then GoTo err
    
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
        rsgrilla.Fields("DeudaTotal") = rsgrilla.Fields("Principal") + rsgrilla.Fields("Intereses") + rsgrilla.Fields("IGVIntereses") + rsgrilla.Fields("InteresAdicional") + rsgrilla.Fields("IGVInteresAdicional")
        
        rsgrilla.Fields("TotalCupon") = Round(CDbl("" & rst.Fields("TotalCupon")), 2)

        rsgrilla.Fields("PagoPrincipal") = 0
        rsgrilla.Fields("PagoIntereses") = 0
        rsgrilla.Fields("PagoIGVIntereses") = 0

        rsgrilla.Fields("PagoInteresAdicional") = 0
        rsgrilla.Fields("PagoIGVInteresAdicional") = 0
        
        'rsGrilla.Fields("PagoTotal") = 0

        rst.MoveNext

    Loop
    
    'Set gDetalle.DataSource = Nothing
    mostrarDatosGridSQL gDetalle, rsgrilla, strMsgError, "NumCupon"

    If strMsgError <> "" Then GoTo err
    
    txtDeudaFecha.Text = gDetalle.Columns.ColumnByFieldName("DeudaTotal").SummaryFooterValue ' gDetalle.Columns.ColumnByFieldName("Principal").SummaryFooterValue + gDetalle.Columns.ColumnByFieldName("Intereses").SummaryFooterValue + gDetalle.Columns.ColumnByFieldName("IGVIntereses").SummaryFooterValue + gDetalle.Columns.ColumnByFieldName("InteresAdicional").SummaryFooterValue + gDetalle.Columns.ColumnByFieldName("IGVInteresAdicional").SummaryFooterValue
    txtMontoRecibido.Text = txtDeudaFecha.Value
    
    txtMontoRecibido_KeyPress 13
    Exit Sub
err:

    If strMsgError = "" Then strMsgError = err.Description
End Sub

Private Sub CalculaPrelacion(ByRef strMsgError As String)
    Dim dblMontoRestado As Double
 
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

    gDetalle.Dataset.First

    Do While Not gDetalle.Dataset.EOF And dblMontoRestado <> 0
        
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


    Loop

    gDetalle.Dataset.First
    Exit Sub
err:

    If strMsgError = "" Then strMsgError = err.Description
End Sub

