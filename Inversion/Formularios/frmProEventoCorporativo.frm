VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{830D5A06-9B70-4F8C-98B6-7A19C4E7760B}#1.0#0"; "TAMControls.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmProEventoCorporativo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Confirmación Eventos Corporativos"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   11595
   Begin TAMControls2.ucBotonEdicion2 cmdOpcion 
      Height          =   735
      Left            =   1200
      TabIndex        =   81
      Top             =   8160
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   1296
      Buttons         =   2
      Caption0        =   "Con&firmar"
      Tag0            =   "3"
      Visible0        =   0   'False
      ToolTipText0    =   "Confirmar"
      Caption1        =   "&Eliminar"
      Tag1            =   "4"
      Visible1        =   0   'False
      ToolTipText1    =   "Eliminar"
      UserControlWidth=   2700
   End
   Begin MSAdodcLib.Adodc adoConsulta 
      Height          =   330
      Left            =   4740
      Top             =   7890
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
   Begin TabDlg.SSTab tabEvento 
      Height          =   7995
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   14102
      _Version        =   393216
      Style           =   1
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
      TabPicture(0)   =   "frmProEventoCorporativo.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraCriterios"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "tdgConsulta"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Confirmación"
      TabPicture(1)   =   "frmProEventoCorporativo.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdAccion"
      Tab(1).Control(1)=   "fraDatos"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Forma de Pago"
      TabPicture(2)   =   "frmProEventoCorporativo.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).Control(1)=   "frmInicio"
      Tab(2).ControlCount=   2
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   -68040
         TabIndex        =   83
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
      Begin VB.Frame Frame1 
         Caption         =   "Detalle de la Forma de Pago"
         Height          =   4035
         Left            =   -74790
         TabIndex        =   37
         Top             =   570
         Width           =   12435
         Begin VB.CommandButton cmdAccionFP 
            Caption         =   "&Adicionar"
            Height          =   375
            Index           =   0
            Left            =   10500
            TabIndex        =   50
            Top             =   2790
            Width           =   1275
         End
         Begin VB.CommandButton cmdAccionFP 
            Caption         =   "Ca&ncelar"
            Height          =   375
            Index           =   4
            Left            =   10500
            TabIndex        =   49
            Top             =   3300
            Width           =   1275
         End
         Begin VB.ComboBox cboBancoOrigen 
            Height          =   315
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   1920
            Width           =   4545
         End
         Begin VB.ComboBox cboCtaBancariaOrig 
            Height          =   315
            Left            =   7890
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   1920
            Width           =   4065
         End
         Begin VB.TextBox txtAnombreOrig 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1980
            MaxLength       =   45
            TabIndex        =   46
            Top             =   2340
            Width           =   4545
         End
         Begin VB.CommandButton cmdAccionFP 
            Caption         =   "&Actualizar"
            Height          =   375
            Index           =   3
            Left            =   10500
            TabIndex        =   45
            Top             =   2790
            Width           =   1275
         End
         Begin VB.TextBox txtTipoCambioFP2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.0000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   9510
            MaxLength       =   45
            TabIndex        =   44
            Top             =   480
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.ComboBox cboModoFPago 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   480
            Width           =   2475
         End
         Begin VB.TextBox txtObservaciones 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   645
            Left            =   1980
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   42
            Top             =   3210
            Width           =   5805
         End
         Begin VB.TextBox txtFax 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   7290
            MaxLength       =   15
            TabIndex        =   41
            Top             =   2760
            Width           =   1800
         End
         Begin VB.ComboBox cboFormaPago 
            Height          =   315
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   900
            Width           =   2475
         End
         Begin VB.ComboBox cboMonedaFP 
            Height          =   315
            Left            =   6060
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   900
            Visible         =   0   'False
            Width           =   2115
         End
         Begin VB.TextBox txtContacto 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1980
            MaxLength       =   50
            TabIndex        =   38
            Top             =   2760
            Width           =   4545
         End
         Begin TAMControls.TAMTextBox txtMontoFPago 
            Height          =   315
            Left            =   10050
            TabIndex        =   51
            Top             =   1290
            Width           =   1755
            _ExtentX        =   3096
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
            Locked          =   -1  'True
            Container       =   "frmProEventoCorporativo.frx":0054
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
         Begin TAMControls.TAMTextBox txtMontoMonedaPago 
            Height          =   315
            Left            =   1980
            TabIndex        =   52
            Top             =   1320
            Width           =   1875
            _ExtentX        =   3307
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
            Container       =   "frmProEventoCorporativo.frx":0070
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
         Begin TAMControls.TAMTextBox txtTipoCambioFP 
            Height          =   315
            Left            =   6060
            TabIndex        =   53
            Top             =   1320
            Width           =   1545
            _ExtentX        =   2725
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
            Container       =   "frmProEventoCorporativo.frx":008C
            Text            =   "0.0000"
            Decimales       =   4
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   12648447
            EnterTab        =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   4
         End
         Begin VB.Label lblMoneda2Arbitraje 
            AutoSize        =   -1  'True
            Caption         =   "PEN"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   11880
            TabIndex        =   68
            Top             =   1350
            Width           =   420
         End
         Begin VB.Label lblMoneda1Arbitraje 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "PEN"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   4080
            TabIndex        =   67
            Top             =   1350
            Width           =   330
         End
         Begin VB.Label lblMonedasArbitraje 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "(PEN/USD)"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   7650
            TabIndex        =   66
            Top             =   1350
            Width           =   840
         End
         Begin VB.Label lblMontoMonedaOpe 
            AutoSize        =   -1  'True
            Caption         =   "Monto Moneda"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   8790
            TabIndex        =   65
            Top             =   1350
            Width           =   1080
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto Moneda Pago"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   96
            Left            =   300
            TabIndex        =   64
            Top             =   1350
            Width           =   1500
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   95
            Left            =   4890
            TabIndex        =   63
            Top             =   960
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Banco"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   89
            Left            =   300
            TabIndex        =   62
            Top             =   1950
            Width           =   465
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Cta.Bancaria"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   90
            Left            =   6720
            TabIndex        =   61
            Top             =   1950
            Width           =   915
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "A nombre de"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   91
            Left            =   300
            TabIndex        =   60
            Top             =   2370
            Width           =   900
         End
         Begin VB.Label lblTC 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Cambio"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   4860
            TabIndex        =   59
            Top             =   1350
            Width           =   1110
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Ejecutado al "
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   87
            Left            =   300
            TabIndex        =   58
            Top             =   540
            Width           =   930
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   94
            Left            =   300
            TabIndex        =   57
            Top             =   3240
            Width           =   1065
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fax"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   92
            Left            =   6810
            TabIndex        =   56
            Top             =   2790
            Width           =   375
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Forma de Pago"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   98
            Left            =   300
            TabIndex        =   55
            Top             =   960
            Width           =   1080
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Contacto"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   93
            Left            =   300
            TabIndex        =   54
            Top             =   2775
            Width           =   645
         End
         Begin VB.Line Line4 
            X1              =   240
            X2              =   12150
            Y1              =   1770
            Y2              =   1770
         End
      End
      Begin VB.Frame frmInicio 
         Caption         =   "Al Inicio"
         Height          =   2925
         Left            =   -74820
         TabIndex        =   28
         Top             =   4590
         Width           =   12465
         Begin VB.CommandButton cmdAccionFP 
            Caption         =   "&Eliminar"
            Height          =   375
            Index           =   1
            Left            =   9000
            TabIndex        =   30
            Top             =   2190
            Width           =   1275
         End
         Begin VB.CommandButton cmdAccionFP 
            Caption         =   "&Modificar"
            Height          =   375
            Index           =   2
            Left            =   10440
            TabIndex        =   29
            Top             =   2190
            Width           =   1215
         End
         Begin DXDBGRIDLibCtl.dxDBGrid gFPagoIni 
            Height          =   1365
            Left            =   300
            OleObjectBlob   =   "frmProEventoCorporativo.frx":00A8
            TabIndex        =   31
            Top             =   510
            Width           =   11775
         End
         Begin TAMControls.TAMTextBox txtSumaInicio 
            Height          =   315
            Left            =   5850
            TabIndex        =   32
            Top             =   2430
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   556
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            FontBold        =   -1  'True
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Locked          =   -1  'True
            Container       =   "frmProEventoCorporativo.frx":73A7
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
         Begin TAMControls.TAMTextBox txtMontoTotalOpera 
            Height          =   315
            Left            =   5850
            TabIndex        =   33
            Top             =   2040
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   556
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            FontBold        =   -1  'True
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Locked          =   -1  'True
            Container       =   "frmProEventoCorporativo.frx":73C3
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
         Begin VB.Label lblDescripInicio 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000006&
            Height          =   225
            Left            =   270
            TabIndex        =   36
            Top             =   1830
            Width           =   5565
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto Total Formas de Pago"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   77
            Left            =   3120
            TabIndex        =   35
            Top             =   2460
            Width           =   2460
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Total Operación"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   78
            Left            =   4170
            TabIndex        =   34
            Top             =   2100
            Width           =   1380
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmProEventoCorporativo.frx":73DF
         Height          =   4425
         Left            =   390
         OleObjectBlob   =   "frmProEventoCorporativo.frx":73F9
         TabIndex        =   16
         Top             =   2670
         Width           =   10395
      End
      Begin VB.Frame fraDatos 
         Caption         =   "Datos"
         Height          =   6345
         Left            =   -74610
         TabIndex        =   6
         Top             =   570
         Width           =   10545
         Begin VB.ComboBox cboAgente 
            Height          =   315
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   73
            ToolTipText     =   "Agente"
            Top             =   1770
            Width           =   7635
         End
         Begin MSComCtl2.DTPicker dtpFechaEntrega 
            Height          =   315
            Left            =   2160
            TabIndex        =   19
            Top             =   2700
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Format          =   175570945
            CurrentDate     =   38790
         End
         Begin TAMControls.TAMTextBox txtValorComision 
            Height          =   315
            Left            =   7620
            TabIndex        =   69
            Top             =   5370
            Width           =   2235
            _ExtentX        =   3942
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
            Container       =   "frmProEventoCorporativo.frx":D289
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
         Begin TAMControls.TAMTextBox txtValor 
            Height          =   315
            Left            =   7620
            TabIndex        =   70
            Top             =   4440
            Width           =   2235
            _ExtentX        =   3942
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
            Container       =   "frmProEventoCorporativo.frx":D2A5
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
         Begin TAMControls.TAMTextBox txtTasaComision 
            Height          =   315
            Left            =   5190
            TabIndex        =   71
            Top             =   5370
            Width           =   2235
            _ExtentX        =   3942
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
            Container       =   "frmProEventoCorporativo.frx":D2C1
            Text            =   "0.000000"
            Decimales       =   6
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            AceptaNegativos =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   6
            MaximoValor     =   999999999
         End
         Begin MSComCtl2.DTPicker dtpFechaVencimiento 
            Height          =   315
            Left            =   2160
            TabIndex        =   79
            Top             =   3150
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Format          =   175570945
            CurrentDate     =   38790
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Vencimiento"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   17
            Left            =   360
            TabIndex        =   80
            Top             =   3210
            Width           =   870
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Titulo Referencia"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   20
            Left            =   360
            TabIndex        =   78
            Top             =   1350
            Width           =   1215
         End
         Begin VB.Label lblTituloReferencia 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   2160
            TabIndex        =   77
            Top             =   1290
            Width           =   7605
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Titulo"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   18
            Left            =   360
            TabIndex        =   76
            Top             =   930
            Width           =   390
         End
         Begin VB.Label lblTitulo 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   2160
            TabIndex        =   75
            Top             =   870
            Width           =   7605
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Acuerdo"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   16
            Left            =   360
            TabIndex        =   74
            Top             =   480
            Width           =   1185
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Agente"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   15
            Left            =   360
            TabIndex        =   72
            Top             =   1830
            Width           =   510
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   14
            Left            =   810
            TabIndex        =   27
            Top             =   2190
            Width           =   45
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   12
            Left            =   840
            TabIndex        =   26
            Top             =   2220
            Width           =   45
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   11
            Left            =   930
            TabIndex        =   25
            Top             =   2520
            Width           =   45
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   10
            Left            =   840
            TabIndex        =   24
            Top             =   3690
            Width           =   45
         End
         Begin VB.Label lblDescripMoneda 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Nuevos Soles (S/.)"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   8370
            TabIndex        =   23
            Top             =   3720
            Width           =   1335
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000015&
            X1              =   330
            X2              =   10110
            Y1              =   5040
            Y2              =   5040
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   7620
            TabIndex        =   22
            Top             =   5790
            Width           =   2235
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Total"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   13
            Left            =   6330
            TabIndex        =   21
            Top             =   5850
            Width           =   360
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comision"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   9
            Left            =   4170
            TabIndex        =   20
            Top             =   5400
            Width           =   630
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Entrega"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   8
            Left            =   360
            TabIndex        =   18
            Top             =   2760
            Width           =   555
         End
         Begin VB.Label lblDescrip 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "TipoAcuerdo"
            ForeColor       =   &H00800000&
            Height          =   315
            Index           =   7
            Left            =   2145
            TabIndex        =   17
            Top             =   420
            Width           =   7605
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Derechos"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   4
            Left            =   6300
            TabIndex        =   15
            Top             =   4020
            Width           =   690
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Real"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   5
            Left            =   6300
            TabIndex        =   14
            Top             =   4485
            Width           =   330
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   6
            Left            =   6300
            TabIndex        =   13
            Top             =   2775
            Width           =   585
         End
         Begin VB.Label lblMoneda 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nuevos Soles"
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   7620
            TabIndex        =   12
            Top             =   2700
            Width           =   2175
         End
         Begin VB.Label lblValor 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.000000"
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   7620
            TabIndex        =   11
            Top             =   4020
            Width           =   2205
         End
         Begin VB.Line lineDatos 
            BorderColor     =   &H80000015&
            X1              =   390
            X2              =   10110
            Y1              =   3570
            Y2              =   3570
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Corte"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   10
            Top             =   2310
            Width           =   375
         End
         Begin VB.Label lblFechaOperacion 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "01/01/2002"
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   2160
            TabIndex        =   9
            Top             =   2250
            Width           =   1575
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Acciones Base"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   6300
            TabIndex        =   8
            Top             =   2310
            Width           =   1065
         End
         Begin VB.Label lblCantAcciones 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   7620
            TabIndex        =   7
            Top             =   2280
            Width           =   2175
         End
      End
      Begin VB.Frame fraCriterios 
         Caption         =   "Criterios de Búsqueda"
         Height          =   1575
         Left            =   360
         TabIndex        =   1
         Top             =   840
         Width           =   10395
         Begin VB.ComboBox cboFondo 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   450
            Width           =   5865
         End
         Begin VB.ComboBox cboEvento 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   930
            Width           =   3735
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fondo"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   720
            TabIndex        =   5
            Top             =   450
            Width           =   450
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Evento"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   720
            TabIndex        =   4
            Top             =   930
            Width           =   510
         End
      End
   End
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   8040
      TabIndex        =   82
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
End
Attribute VB_Name = "frmProEventoCorporativo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Ficha de Confirmación de Eventos Corporativos"
Option Explicit

Dim arrFondo()              As String, arrEvento()          As String
Dim arrAgente()             As String

'Variables agregadas por las formas de pago ACC 12/05/2009
Dim arrBcoOrigen()          As String, arrBcoDestino()              As String
Dim arrCtaBancariaOrig()    As String, arrCtaBancariaDest()         As String
Dim arrFormaPago()          As String, arrMonedaFP()                As String
Dim arrModoFPago()          As String

Dim strCodFondo             As String, strCodEvento         As String
Dim strEstado               As String, strSQL               As String
Dim strModoFPago            As String, strClaseFPagoIni     As String
Dim strCodModoFPago         As String, strCodMonedaFP       As String

Dim strCodFile              As String, strCodAnalitica      As String
Dim strCodMoneda            As String, strCodDetalleFile    As String
Dim strCodTitulo            As String, strCodAgente         As String
Dim dblValorNominal         As Double, strCodEmisor         As String
Dim lngNumAcuerdo           As Long

Dim strCodAnaliticaReferencia As String
Dim strCodFileReferencia As String
Dim strCodTituloReferencia As String

' Variables adicionadas ACC 01/03/2010
Dim indInicializaGrilla     As Boolean
Dim indCargaPantalla        As Boolean
Dim indLlamadoCboFP         As Boolean
Dim rsg                     As New ADODB.Recordset
Dim rsgVcto                 As New ADODB.Recordset
Dim indInserta              As Boolean
Dim indActualizaFP          As Boolean
Dim dblSumaFPCnt            As Double
Dim dblSumaFPVto            As Double
Dim strDefectTipoCuenta     As String
Dim strDefectCodBanco       As String
Dim strDefectNumCuenta      As String
Dim strIndCuentaCorriente, strIndCuentaAhorros As String
Dim dblBkpMontoFPago        As Double
Dim strCodMonedaParEvaluacion As String
Dim strCodMonedaParPorDefecto As String


Public Sub Accion(nAccion As ButtonAction)
    
    Select Case nAccion
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

Private Sub RegistrarKardex()
    
    '*** Registro del Kardex para Acciones Liberadas ***
    Dim adoTemporal                 As ADODB.Recordset
    Dim strFechaRegistro            As String, strFechaMas1Dia              As String
    Dim strCodTipoOrden             As String, strIndPorcenPrecio           As String
    Dim strCodGarantia              As String, strIndInversion              As String
    Dim strIndTitulo                As String, strIndGenerado               As String
    Dim strIndCuponCero             As String, strDescripOrden              As String
    Dim strCodMonedaPago            As String, strCodSubDetalleFile         As String
    Dim strFechaOrden               As String, strFechaGrabar               As String
    Dim strFechaLiquidacion         As String, strFechaEmision              As String
    Dim strFechaVencimiento         As String, strFechaPago                 As String
    'Dim strCodEmisor                As String, strCodCiiu                   As String
    Dim strCodCiiu                   As String
    
    Dim strCodGrupo                 As String, strCodSector                 As String
    Dim strCodTipoTasa              As String, strCodBaseAnual              As String
    Dim strCodAgente                As String, strCodOperacion              As String
    Dim strCodNegociacion           As String, strCodOrigen                 As String
    Dim strCodReportado             As String, strCodGirador                As String
    Dim strCodAceptante             As String, strIndCustodia               As String
    Dim strIndKardex                As String, strTipoTasa                  As String
    Dim strBaseAnual                As String, strRiesgo                    As String
    Dim strSubRiesgo                As String, strCodNemonico               As String
    Dim strIndAmortizacion          As String, strFlgTvac                   As String
    Dim strIndUltimoMovimiento      As String, strTipoMovimientoKardex      As String
    Dim strNumOrden                 As String, strCodTipoOperacion          As String
    Dim strObservacion              As String
    Dim dblPrecioUnitario           As Double, dblTipoCambio                As Double
    Dim dblTipoCambioMonedaPago     As Double, dblTasaInteres               As Double
    Dim dblTirBruta                 As Double, dblTirNeta                   As Double
    Dim dblMontoVencimiento         As Double, dblKarValProm                As Double
    Dim dblKarIadProm               As Double, dblKarSldAmort               As Double
    Dim dblTirPromAnt               As Double, dblValorAmortizacion         As Double
    Dim dblTirNetaKardex            As Double, dblTirOperacionKardex        As Double
    Dim dblTirPromedioKardex        As Double, dblValorPromedioKardex       As Double
    Dim dblInteresCorridoPromedio   As Double
    Dim curValComi                  As Currency, curCantOrden               As Currency
    Dim curValorMovimiento          As Currency, curValorNominal            As Currency
    Dim curCantMovimiento           As Currency, curVacCorrido              As Currency
    Dim curKarSldInic               As Currency, curKarSldFina              As Currency
    Dim curKarValSald               As Currency, curKarIadSald              As Currency
    Dim curSaldoInicialKardex       As Currency, curSaldoFinalKardex        As Currency
    Dim curValorSaldoKardex         As Currency, curSaldoInteresCorrido     As Currency
    Dim curSaldoAmortizacion        As Currency
    Dim intDiasPlazo                As Integer
        
    With adoComm
        Set adoTemporal = New ADODB.Recordset

        strFechaRegistro = Convertyyyymmdd(CDate(lblFechaOperacion.Caption))
        strFechaMas1Dia = Convertyyyymmdd(DateAdd("d", 1, CDate(lblFechaOperacion.Caption)))
        strFechaVencimiento = Convertyyyymmdd(dtpFechaVencimiento.Value)

        '*** Obtener Secuenciales ***
        strNumOperacion = Valor_Caracter 'ObtenerSecuencialInversionOperacion(strCodFondo, Valor_NumOperacion)
        strNumKardex = Valor_Caracter 'ObtenerSecuencialInversionOperacion(strCodFondo, Valor_NumKardex)
        strNumOrden = Valor_Caracter

        strCodTipoOrden = Codigo_Orden_Compra
        strCodTipoOperacion = Codigo_Operacion_Contado

        '*** El Precio es % ? ***
        strIndPorcenPrecio = Valor_Indicador

        strIndTitulo = Valor_Caracter
        strCodGarantia = Valor_Caracter
        strIndInversion = Valor_Caracter
        strIndGenerado = Valor_Caracter
        strIndCuponCero = Valor_Caracter
        dblPrecioUnitario = 0

        '*** Valores Comunes ***
        curCantOrden = CCur(txtValor.Text)
        curValorMovimiento = 0
        curValorNominal = 1
        curCantMovimiento = curCantOrden * curValorNominal
        curVacCorrido = 0
        curCtaCostoSAB = 0
        curCtaCostoBVL = 0
        curCtaCostoCavali = 0
        curCtaCostoConasev = 0
        curCtaCostoFondoLiquidacion = 0
        curCtaCostoFondoGarantia = 0
        curCtaGastoBancario = 0
        curCtaComisionEspecial = 0
        curCtaImpuesto = 0
        strDescripOrden = Trim(tdgConsulta.Columns("DescripEntrega").Value)
        strCodMonedaPago = strCodMoneda
        dblTipoCambio = 0 'gdblTipoCambio
        dblTipoCambioMonedaPago = dblTipoCambio
        strCodSubDetalleFile = Valor_Caracter
        strFechaOrden = strFechaRegistro
        strFechaGrabar = strFechaOrden & Space(1) & Format(Time, "hh:mm")
        strFechaLiquidacion = strFechaRegistro
        strFechaEmision = strFechaRegistro
        
        'strFechaVencimiento = strFechaRegistro
        
        strFechaPago = strFechaRegistro
        strCodEmisor = Valor_Caracter
        strCodCiiu = Valor_Caracter
        strCodGrupo = Valor_Caracter
        strCodSector = Valor_Caracter
        strCodTipoTasa = Valor_Caracter
        strCodBaseAnual = Valor_Caracter
        dblTasaInteres = 0
        dblTirBruta = 0
        dblTirNeta = 0
        dblMontoVencimiento = 0
        intDiasPlazo = 0
        strCodAgente = Valor_Caracter
        strCodOperacion = Valor_Caracter
        strCodNegociacion = Valor_Caracter
        strCodOrigen = Valor_Caracter
        strCodReportado = Valor_Caracter
        strCodGirador = Valor_Caracter
        strCodAceptante = Valor_Caracter
        strIndCustodia = Valor_Caracter
        strIndKardex = Valor_Indicador
        strTipoTasa = Valor_Caracter
        strBaseAnual = Valor_Caracter
        strRiesgo = Valor_Caracter
        strSubRiesgo = Valor_Caracter
        strCodNemonico = Valor_Caracter
        strObservacion = Trim(tdgConsulta.Columns("DescripEntrega").Value)

        strIndAmortizacion = Valor_Caracter
        strFlgTvac = Valor_Caracter
            
            
        .CommandText = "SELECT dbo.uf_ACObtenerTipoCambioMoneda1('" & gstrCodClaseTipoCambioOperacionFondo & "','" & Codigo_Valor_TipoCambioCompra & "','" & gstrFechaActual & "','" & strCodMoneda & "','" & Codigo_Moneda_Local & "',5) AS 'ValorTipoCambio'"
        Set adoTemporal = .Execute
        
        If Not adoTemporal.EOF Then
            dblTipoCambio = adoTemporal("ValorTipoCambio")
        Else
            dblTipoCambio = 1
        End If
        adoTemporal.Close
            
        '*** Obtener Inventario Actual del Kardex ***
        '*** NO tomar en cuenta los Mov. Anulados (IndAnulado<>'X') ***
        .CommandText = "SELECT SaldoInicial,SaldoFinal,MontoSaldo,SaldoInteresCorrido,PromedioInteresCorrido,ValorPromedio,SaldoAmortizacion,TirPromedio FROM InversionKardex " & _
            "WHERE CodAnalitica='" & strCodAnalitica & "' AND CodFile='" & strCodFile & "' AND " & _
            "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
            "NumKardex = dbo.uf_IVObtenerUltimoMovimientoKardexValor(CodFondo,CodAdministradora,CodTitulo,'" & strFechaRegistro & "') AND " & _
            "SaldoFinal > 0 " & _
            "ORDER BY NumKardex"
        Set adoTemporal = .Execute

        If adoTemporal.EOF Then
            curKarSldInic = 0: curKarSldFina = 0: curKarValSald = 0: dblKarValProm = 0
        Else
            curKarSldInic = CCur(adoTemporal("SaldoFinal"))
            curKarSldFina = CCur(adoTemporal("SaldoFinal"))
            curKarValSald = CCur(adoTemporal("MontoSaldo"))
            curKarIadSald = CCur(adoTemporal("SaldoInteresCorrido"))
            dblKarIadProm = CDbl(adoTemporal("PromedioInteresCorrido"))
            dblKarValProm = CDbl(adoTemporal("ValorPromedio"))
            dblKarSldAmort = CDbl(adoTemporal("SaldoAmortizacion"))
            dblTirPromAnt = CDbl(adoTemporal("TirPromedio"))
        End If
        adoTemporal.Close: Set adoTemporal = Nothing

        curValComi = 0
        dblPrecioUnitario = 0
        strIndUltimoMovimiento = "X"
        
        If curCantOrden > 0 Then
            strTipoMovimientoKardex = "E"
        Else
            strTipoMovimientoKardex = "S"
        End If
                
        If curKarSldInic = 0 Then '*** Primera Compra ***
            curSaldoInicialKardex = 0
            curSaldoFinalKardex = curCantOrden
            curValorSaldoKardex = curValorMovimiento
            curSaldoInteresCorrido = 0
            curSaldoAmortizacion = curSaldoFinalKardex
        Else
            curSaldoInicialKardex = curKarSldFina
            curSaldoFinalKardex = curKarSldFina + curCantOrden
            curSaldoAmortizacion = curSaldoFinalKardex
        End If
                
        If curKarSldFina = 0 Then
            curSaldoInteresCorrido = 0
            curValorSaldoKardex = curValorMovimiento
        Else
            curSaldoInteresCorrido = curKarIadSald
            curValorSaldoKardex = curKarValSald
        End If

        dblTirNetaKardex = 0
        dblTirOperacionKardex = 0
        dblTirPromedioKardex = 0

        If curSaldoFinalKardex <> 0 Then
            If strIndPorcenPrecio = Valor_Indicador Then
                dblValorPromedioKardex = curValorSaldoKardex / (curSaldoFinalKardex * dblValorNominal)
            Else
                dblValorPromedioKardex = curValorSaldoKardex / (curSaldoFinalKardex * dblValorNominal)
            End If
            dblInteresCorridoPromedio = 0
        Else
            dblValorPromedioKardex = 0
            dblInteresCorridoPromedio = 0
        End If

        '*** Transacción ***
        gblnRollBack = False
        
        .CommandText = "{ call up_IVProcKardexEvento('" & strCodFondo & "','" & gstrCodAdministradora & "','" & strFechaLiquidacion & "','" & strFechaOrden & "','" & strFechaGrabar & "','" & _
            gstrPeriodoActual & "','" & gstrMesActual & "','" & strCodFile & "','" & strCodAnalitica & "','" & strIndTitulo & "','" & strCodTitulo & "','" & strCodDetalleFile & "','" & _
            strCodMoneda & "','" & strCodMonedaPago & "','" & strCodMonedaPago & "','" & strNumOrden & "','" & strNumOperacion & "','" & strNumKardex & "','" & strNumAsiento & "','" & strCodEmisor & "','" & strCodAgente & "','" & _
            strTipoMovimientoKardex & "'," & CDec(curCantOrden) & "," & CDec(dblPrecioUnitario) & "," & CDec(dblPrecioUnitario) & "," & CDec(dblPrecioUnitario) & "," & CDec(curValorMovimiento) & "," & CDec(curValComi) & "," & CDec(curSaldoInicialKardex) & "," & CDec(curSaldoFinalKardex) & "," & _
            CDec(curValorSaldoKardex) & ",'" & strDescripOrden & "'," & CDec(dblValorPromedioKardex) & ",'" & strIndUltimoMovimiento & "'," & CDec(dblInteresCorridoPromedio) & "," & CDec(curSaldoInteresCorrido) & "," & dblTirOperacionKardex & "," & _
            dblTirOperacionKardex & "," & dblTirPromedioKardex & "," & dblTirPromedioKardex & "," & CDec(curVacCorrido) & "," & CDec(dblTirNetaKardex) & "," & CDec(curSaldoAmortizacion) & ",'01','" & strCodSubDetalleFile & "','" & strCodTipoOperacion & "','" & strCodNegociacion & "','" & _
            strCodOrigen & "','" & strCodGarantia & "','" & strFechaVencimiento & "','" & strFechaEmision & "'," & CDec(curValorNominal) & ",0,0,0,0,0,0,0,0,0,0,0,0,0,0," & _
            "0,0,0,0,0,0,0,0,0,0,0,0,0," & CDec(dblMontoVencimiento) & "," & intDiasPlazo & ",'" & strCodGrupo & "','" & strCodOperacion & "','" & _
            strCodReportado & "','" & strCodGirador & "','" & strCodAceptante & "','" & strIndCustodia & "','" & strIndKardex & "','" & strTipoTasa & "','" & strBaseAnual & "'," & CDec(dblTasaInteres) & "," & _
            CDec(dblTirBruta) & "," & CDec(dblTirNeta) & ",'" & strRiesgo & "','" & strSubRiesgo & "','" & strObservacion & "'," & CDec(dblTipoCambio) & "," & CDec(dblTipoCambioMonedaPago) & ",'" & gstrLogin & "') }"
        
        adoConn.Execute .CommandText

    
    
    End With
    
End Sub

Private Sub RegistrarMovimiento()

    '*** Registrar asiento contable - Orden Cobro/Pago - Operación ***
    Call GenerarAsientoDividendo(strCodFile, strCodAnalitica, strCodFondo, gstrCodAdministradora, strCodDetalleFile, Codigo_Dinamica_Dividendos, CCur(txtValor.Text), CCur(txtValorComision.Text), strCodMoneda, Trim(tdgConsulta.Columns("DescripEntrega").Value), frmMainMdi.Tag, Trim(tdgConsulta.Columns("CodTitulo").Value), dtpFechaEntrega.Value, Codigo_Tipo_Persona_Agente, strCodAgente, strCodEmisor)
    
End Sub


Public Sub Salir()

    Unload Me
    
End Sub
Public Sub Cancelar()

    cmdOpcion.Visible = True
    With tabEvento
        .TabEnabled(0) = True
        .Tab = 0
    End With
    Call Buscar
    
End Sub
Public Sub Grabar()

    Dim strFechaInicio  As String, strFechaFin  As String
    Dim strFechaEntrega As String
    Dim intRegistro     As Integer
    
    If strEstado = Reg_Consulta Then Exit Sub
    
    If strEstado = Reg_Edicion Then
        If TodoOK() Then
            strFechaEntrega = Convertyyyymmdd(dtpFechaEntrega.Value)
            
            Select Case tdgConsulta.Columns("TipoAcuerdo").Value
                Case Codigo_Evento_Liberacion, Codigo_Evento_Preferente, Codigo_Evento_Nominal '*** Movimiento en Kardex **
                    
                    If lblFechaOperacion.Caption = gdatFechaActual Then
                        Call RegistrarKardex
                   
                        '*** Actualizar Orden Evento ***
                        adoComm.CommandText = "UPDATE EventoCorporativoOrden SET " & _
                            "EstadoEvento='" & Estado_Entrega_Procesado & "',CantLiberadasReal=" & CDec(txtValor.Text) & ",FechaEntrega='" & strFechaEntrega & "', MontoComision=" & CCur(txtValorComision.Text) & ", PorcenComision=" & CCur(txtTasaComision.Text)
                    Else
                        '*** Actualizar Orden Evento ***
                        adoComm.CommandText = "UPDATE EventoCorporativoOrden SET " & _
                            "CantLiberadasReal=" & CDec(txtValor.Text) & ",FechaEntrega='" & strFechaEntrega & "', MontoComision=" & CCur(txtValorComision.Text) & ", PorcenComision=" & CCur(txtTasaComision.Text)
                    End If
                
                Case Codigo_Evento_Dividendo '*** Movimiento Contable ***
                    '*** Verificar dinámica ***
                    If lblFechaOperacion.Caption = gdatFechaActual Then
                        If Not ExisteDinamica(strCodFile, strCodDetalleFile, gstrCodAdministradora, Codigo_Dinamica_Dividendos, strCodMoneda) Then Exit Sub
                        
                        Call RegistrarMovimiento
                        
                        '*** Actualizar Orden Evento ***
                        adoComm.CommandText = "UPDATE EventoCorporativoOrden SET " & _
                            "EstadoEvento='" & Estado_Entrega_Procesado & "',MontoDividendosReal=" & CDec(txtValor.Text) & ",FechaEntrega='" & strFechaEntrega & "', MontoComision=" & CCur(txtValorComision.Text) & ", PorcenComision=" & CCur(txtTasaComision.Text)
                    
                    Else
                        '*** Actualizar Orden Evento ***
                        adoComm.CommandText = "UPDATE EventoCorporativoOrden SET " & _
                            "MontoDividendosReal=" & CDec(txtValor.Text) & ",FechaEntrega='" & strFechaEntrega & "', MontoComision=" & CCur(txtValorComision.Text) & ", PorcenComision=" & CCur(txtTasaComision.Text)
                    End If
                    
'                Case Codigo_Evento_Nominal
'
'                    If lblFechaOperacion.Caption = gdatFechaActual Then
'                        Call RegistrarKardex
'
'                        '*** Actualizar Orden Evento ***
'                        adoComm.CommandText = "UPDATE EventoCorporativoOrden SET " & _
'                            "EstadoEvento='" & Estado_Entrega_Procesado & "',ValorNominalReal=" & CDec(txtValor.Text) & ",FechaEntrega='" & strFechaEntrega & "', MontoComision=" & CCur(txtValorComision.Text) & ", PorcenComision=" & CCur(txtTasaComision.Text)
'                    Else
'                        adoComm.CommandText = "UPDATE EventoCorporativoOrden SET " & _
'                            "ValorNominalReal=" & CDec(txtValor.Text) & ",FechaEntrega='" & strFechaEntrega & "', MontoComision=" & CCur(txtValorComision.Text) & ", PorcenComision=" & CCur(txtTasaComision.Text)
'                    End If
                    
            
            End Select
            
            adoComm.CommandText = adoComm.CommandText & " WHERE CodFondo = '" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodTitulo='" & Trim(tdgConsulta.Columns("CodTitulo")) & "' AND NumEntrega=" & CInt(tdgConsulta.Columns("NumEntrega"))
            adoConn.Execute adoComm.CommandText

            Me.MousePointer = vbDefault
        
            MsgBox Mensaje_Edicion_Exitosa, vbExclamation
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
            cmdOpcion.Visible = True
            With tabEvento
                .TabEnabled(0) = True
                .Tab = 0
            End With
            
            Call Buscar
        End If
    End If

End Sub

Private Function TodoOK() As Boolean
        
    TodoOK = False
                          
    If CDate(lblFechaOperacion.Caption) > dtpFechaEntrega.Value Then
        MsgBox "La Fecha de Entrega no puede ser menor a la Fecha de Corte.", vbCritical, Me.Caption
        If dtpFechaEntrega.Enabled Then dtpFechaEntrega.SetFocus
        Exit Function
    End If
    
    If CDec(txtValor.Text) = 0 Then
        MsgBox "Debe indicar el valor real del evento.", vbCritical, Me.Caption
        If txtValor.Enabled Then txtValor.SetFocus
        Exit Function
    End If

    '*** Si todo paso OK ***
    TodoOK = True
  
End Function

Public Sub Imprimir()

End Sub
Public Sub Buscar()
    
    Dim adoBuscar       As ADODB.Recordset
            
    strSQL = "SELECT NumEntrega,NumAcuerdo,II.Nemotecnico,III.Nemotecnico as NemotecnicoReferencia,FechaMovimiento," & _
        "EC.CodTitulo,EC.CodTituloReferencia,CantAccionesBase,DescripEntrega,EstadoEvento,TipoAcuerdo," & _
        "(CASE WHEN TipoAcuerdo IN ('" & Codigo_Evento_Liberacion & "','" & Codigo_Evento_Preferente & "','" & Codigo_Evento_Nominal & "') THEN  CantLiberadas " & _
        " WHEN TipoAcuerdo='" & Codigo_Evento_Dividendo & "' THEN MontoDividendos " & _
        " ELSE 0 END) ValorEvento " & _
        "FROM EventoCorporativoOrden EC " & _
        "JOIN InstrumentoInversion II ON(II.CodTitulo=EC.CodTitulo) " & _
        "JOIN InstrumentoInversion III ON(III.CodTitulo=EC.CodTituloReferencia) " & _
        "WHERE EC.CodFondo='" & strCodFondo & "' AND EC.CodAdministradora='" & gstrCodAdministradora & "' AND " & _
        "TipoAcuerdo='" & strCodEvento & "' AND EstadoEvento='" & Estado_Entrega_Generado & "' ORDER BY NumEntrega"
                                
                                
    strSQL = "SELECT NumEntrega,NumAcuerdo,II.Nemotecnico,FechaMovimiento," & _
        "EC.CodTitulo,CantAccionesBase,DescripEntrega,EstadoEvento,TipoAcuerdo," & _
        "(CASE WHEN TipoAcuerdo IN ('" & Codigo_Evento_Liberacion & "','" & Codigo_Evento_Preferente & "','" & Codigo_Evento_Nominal & "') THEN  CantLiberadas " & _
        " WHEN TipoAcuerdo='" & Codigo_Evento_Dividendo & "' THEN MontoDividendos " & _
        " ELSE 0 END) ValorEvento " & _
        "FROM EventoCorporativoOrden EC " & _
        "JOIN InstrumentoInversion II ON(II.CodTitulo=EC.CodTitulo) " & _
        "WHERE EC.CodFondo='" & strCodFondo & "' AND EC.CodAdministradora='" & gstrCodAdministradora & "' AND " & _
        "TipoAcuerdo='" & strCodEvento & "' AND EstadoEvento='" & Estado_Entrega_Generado & "' ORDER BY NumEntrega"
                                
                                
        '"WHEN TipoAcuerdo='" & Codigo_Evento_Nominal & "' THEN EC.ValorNominal ELSE 0 END) ValorEvento " & _

'    strEstado = Reg_Defecto
'
'    Set adoBuscar = New ADODB.Recordset
'
'        adoComm.CommandText = strSql
'        Set adoBuscar = adoComm.Execute
'        tdgConsulta.DataSource = adoBuscar
'
'    If adoBuscar.RecordCount > 0 Then strEstado = Reg_Consulta


        strEstado = Reg_Defecto
        
        Set adoBuscar = New ADODB.Recordset
        
         With adoBuscar
            .ActiveConnection = gstrConnectConsulta
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockBatchOptimistic
            .Open strSQL
         End With

        tdgConsulta.DataSource = adoBuscar
        
        If adoBuscar.RecordCount > 0 Then strEstado = Reg_Consulta



        
End Sub
Public Sub Eliminar()

    If strEstado = Reg_Consulta Or strEstado = Reg_Edicion Then
        If tdgConsulta.Columns("EstadoEvento").Value <> Estado_Entrega_Generado Then Exit Sub
        If MsgBox("Se procederá a eliminar el acuerdo número" & Space(1) & CStr(tdgConsulta.Columns("NumEntrega").Value) & _
            Space(1) & "(" & Trim(tdgConsulta.Columns("Nemotecnico").Value) & ")" & vbNewLine & vbNewLine & vbNewLine & _
            "¿ Seguro de continuar ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
    
            '*** Anular Acuerdo ***
            adoComm.CommandText = "UPDATE EventoCorporativoOrden SET EstadoEvento='" & Estado_Entrega_Anulado & "' WHERE CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodTitulo='" & tdgConsulta.Columns("CodTitulo") & "' AND NumEntrega='" & tdgConsulta.Columns("NumEntrega") & "'"
            adoConn.Execute adoComm.CommandText
                                    
            tabEvento.TabEnabled(0) = True
            tabEvento.Tab = 0
            Call Buscar
            
            Exit Sub
        End If
    End If
    
End Sub
Public Sub Modificar()

    If strEstado = Reg_Consulta Then
        strEstado = Reg_Edicion
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabEvento
            .TabEnabled(0) = False
            .Tab = 1
        End With
    End If
    
End Sub

Private Sub LlenarFormulario(strModo As String)

    Dim adoRecord   As ADODB.Recordset
    
    Select Case strModo
        Case Reg_Edicion
            Set adoRecord = New ADODB.Recordset
            
            
            adoComm.CommandText = "SELECT EC.FechaVencimiento,PorcenComision, MontoComision, CantAccionesBase,II.CodMoneda, " & _
                "EC.CodFile,II.CodDetalleFile,EC.CodAnalitica,EC.CodTituloReferencia,EC.CodFileReferencia,EC.CodAnaliticaReferencia," & _
                "II.DescripTitulo, III.DescripTitulo AS DescripTituloReferencia, FechaCorte,FechaEntrega,II.ValorNominal,EC.TipoAcuerdo " & _
                "FROM EventoCorporativoOrden EC " & _
                "JOIN InstrumentoInversion II ON(II.CodTitulo=EC.CodTitulo) " & _
                "JOIN InstrumentoInversion III ON(III.CodTitulo=EC.CodTituloReferencia) " & _
                "WHERE NumEntrega=" & CInt(tdgConsulta.Columns("NumEntrega").Value) & " AND EC.CodTitulo='" & _
                Trim(tdgConsulta.Columns("CodTitulo").Value) & "' AND EC.CodAdministradora='" & gstrCodAdministradora & "'"
            Set adoRecord = adoComm.Execute
            
            If Not adoRecord.EOF Then
            
                cboAgente.ListIndex = -1
                
                fraDatos.Caption = Trim(tdgConsulta.Columns("Nemotecnico").Value)
                lblDescrip(7).Caption = Trim(cboEvento.Text)
                lblFechaOperacion.Caption = CStr(adoRecord("FechaCorte"))
                lblCantAcciones.Caption = CStr(adoRecord("CantAccionesBase"))
                lblValor.Caption = CStr(tdgConsulta.Columns("ValorEvento").Value)
                strCodMoneda = adoRecord("CodMoneda")
                lblMoneda.Caption = ObtenerDescripcionMoneda(strCodMoneda)
                txtValor.Text = "0"
                strCodDetalleFile = adoRecord("CodDetalleFile")
                strCodAnalitica = adoRecord("CodAnalitica")
                strCodFile = adoRecord("CodFile")
                strCodTitulo = Trim(tdgConsulta.Columns("CodTitulo").Value)
                
                strCodAnaliticaReferencia = adoRecord("CodAnaliticaReferencia")
                strCodFileReferencia = adoRecord("CodFileReferencia")
                strCodTituloReferencia = adoRecord("CodTituloReferencia")
                
                dtpFechaEntrega.Value = adoRecord("FechaEntrega")
                
                dtpFechaVencimiento.Value = adoRecord("FechaVencimiento")
                
                dblValorNominal = CDec(adoRecord("ValorNominal"))
                txtTasaComision.Text = adoRecord("PorcenComision")
                txtValorComision.Text = adoRecord("MontoComision")
                
                lblTitulo.Caption = adoRecord("DescripTitulo")
                lblTituloReferencia.Caption = adoRecord("DescripTituloReferencia")
                                
                dtpFechaVencimiento.Visible = False
                lblDescrip(17).Visible = False
                                
                If adoRecord("TipoAcuerdo") = Codigo_Evento_Dividendo Then
                    lblDescripMoneda.Caption = ObtenerDescripcionMoneda(strCodMoneda) & " " & ObtenerSignoMoneda(strCodMoneda)
                    Line1.Visible = True
                    lblDescrip(9).Visible = True
                    txtTasaComision.Visible = True
                    txtValorComision.Visible = True
                    lblDescrip(13).Visible = True
                    lblTotal.Visible = True
                Else
                    lblDescripMoneda.Caption = "Cantidad de Acciones"
                    Line1.Visible = False
                    lblDescrip(9).Visible = False
                    txtTasaComision.Visible = False
                    txtValorComision.Visible = False
                    lblDescrip(13).Visible = False
                    lblTotal.Visible = False
                    If adoRecord("TipoAcuerdo") = Codigo_Evento_Preferente Then
                        dtpFechaVencimiento.Visible = True
                        lblDescrip(17).Visible = True
                    End If
                End If
                
            End If
            adoRecord.Close: Set adoRecord = Nothing
    End Select
    
End Sub

Private Sub cboAgente_Click()

    strCodAgente = Valor_Caracter
    If cboAgente.ListIndex < 0 Then Exit Sub
    
    strCodAgente = Trim(arrAgente(cboAgente.ListIndex))

End Sub

Private Sub cboEvento_Click()

    strCodEvento = Valor_Caracter
    If cboEvento.ListIndex < 0 Then Exit Sub
    
    strCodEvento = Trim(arrEvento(cboEvento.ListIndex))
    
    Call Buscar
    
End Sub

Private Sub cboFondo_Click()

    Dim adoRegistro         As ADODB.Recordset
    
    strCodFondo = Valor_Caracter
    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(arrFondo(cboFondo.ListIndex))
    
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Fecha Vigente, Moneda, Tipo de Cambio ***
        .CommandText = "{ call up_ACSelDatosParametro(23,'" & strCodFondo & "','" & gstrCodAdministradora & "') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            gdatFechaActual = CVDate(adoRegistro("FechaCuota"))
            gstrFechaActual = Convertyyyymmdd(gdatFechaActual)
            
            gstrPeriodoActual = Format(Year(gdatFechaActual), "0000")
            gstrMesActual = Format(Month(gdatFechaActual), "00")
'            gdblTipoCambio = ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, gdatFechaActual, Codigo_Moneda_Local, adoRegistro("CodMoneda"))
'
'            If gdblTipoCambio = 0 Then gdblTipoCambio = ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, DateAdd("d", -1, gdatFechaActual), Codigo_Moneda_Local, adoRegistro("CodMoneda"))
            
            'ACTUALIZA PARAMETROS GLOBALES POR FONDO
            If Not CargarParametrosGlobales(strCodFondo) Then Exit Sub
            
                        
            frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
            Me.Refresh
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    Call Buscar
    
End Sub

'Public Sub ConfGrid(g As dxDBGrid, indMod As Boolean, Optional mostrarFooter As Boolean, Optional mostrarGroupPanel As Boolean, Optional mostrarBandas As Boolean)
'
''ACC 21/10/2099
'
'     With g.Options
'        '***
'        If indMod Then
'            .Set (egoEditing)
'            .Set (egoCanDelete)
'            .Set (egoCanInsert)
'            '.Set (egoCanAppend)
'        End If
'        If mostrarBandas Then .Set (egoShowBands)
'        .Set (egoTabs)
'        .Set (egoTabThrough)
'        .Set (egoImmediateEditor)
'        .Set (egoShowIndicator)
'        .Set (egoCanNavigation)
'        .Set (egoHorzThrough)
'        .Set (egoVertThrough)
'        '.Set (egoAutoWidth)
'        .Set (egoEnterShowEditor)
'        .Set (egoEnterThrough)
'        .Set (egoShowButtonAlways)
'        .Set (egoColumnSizing)
'        .Set (egoColumnMoving)
'        .Set (egoTabThrough)
'        .Set (egoConfirmDelete)
'        .Set (egoCancelOnExit)
'        .Set (egoLoadAllRecords)
'        .Set (egoShowHourGlass)
'        .Set (egoUseBookmarks)
'        .Set (egoUseLocate)
'        .Set (egoAutoCalcPreviewLines)
'        .Set (egoBandSizing)
'        .Set (egoBandMoving)
'        .Set (egoDragScroll)
'        .Set (egoAutoSort)
'        .Set (egoExpandOnDblClick)
'        .Set (egoNameCaseInsensitive)
'        If mostrarFooter Then .Set (egoShowFooter)
'        .Set (egoShowGrid)
'        .Set (egoShowButtons)
'        .Set (egoShowHeader)
'        .Set (egoShowPreviewGrid)
'        .Set (egoShowBorder)
'        '.Set (egoShowRowFooter)
'        '''''.Set (egoDynamicLoad)
'
'        If mostrarGroupPanel Then .Set (egoShowGroupPanel)
'        .Set (egoEnableNodeDragging)
'        .Set (egoDragCollapse)
'        .Set (egoDragExpand)
'        .Set (egoDragScroll)
'        .Set (egoEnableNodeDragging)
'    End With
'End Sub
'
'
'Private Sub cboFormaPago_Click()
'
'Dim strModoFP           As String
'Dim dblMontoRestanteFP  As Double
'Dim strSQL              As String
'Dim adoRegistro         As ADODB.Recordset
'
'If cboFormaPago.ListIndex > 0 Then
'
'        'Limpiar los campos previamente
'        cboBancoOrigen.ListIndex = -1  '0
'        cboCtaBancariaOrig.Clear
'        txtAnombreOrig.Text = ""
'        txtContacto.Text = ""
'        txtFax.Text = ""
'        txtObservaciones.Text = ""
'
'        If indActualizaFP = False Then      'Es un ingreso. No estoy modificando
'            If strModoFPago = "I" Then      'Forma de pago al inicio o contado
'                dblMontoRestanteFP = FP_CalculaMontoFP(gFPagoIni, CDbl(lblTotal.Caption))
'            End If
'
'            txtMontoMonedaPago.Text = dblMontoRestanteFP
'            'Obtener datos de la operación
'            If Trim(strCodMoneda) = "" Then
'                strCodMoneda = Trim(lblMoneda.Tag)
'            End If
'
'            If cboMonedaFP.ListCount > 0 Then
'                cboMonedaFP.ListIndex = ObtenerItemLista(arrMonedaFP(), strCodMoneda)
'                txtTipoCambioFP.Text = 1#
'                txtMontoFPago.Text = txtMontoMonedaPago.Value
'            Else
'                txtMontoFPago.Text = 0#
'            End If
'
'        End If
'
'        If dblMontoRestanteFP > 0 Or indActualizaFP = True Then   'Aún se pueden agregar formas de pago
'
'            'Obtener la cuenta bancaria y el banco por defecto si la forma de pago corresponde
'            '==================================================================================
'            If cboMonedaFP.ListIndex > 0 Then   'Si ya está seleccionada la moneda
'
'                'Obtener los indicadores de tipo de pago de las formas de pago para determinar si afectan a cuentas bancarias
'                strSQL = "SELECT IndBCR, IndExterior, IndTransferSimple, IndAfectaCtasBanco, IndCuentaCorriente, IndCuentaAhorros FROM FormaPago WHERE CodFormaPago = '" & Trim(arrFormaPago(cboFormaPago.ListIndex)) & "'"
'                Set adoRegistro = New ADODB.Recordset
'                With adoComm
'                    .CommandText = strSQL
'                    Set adoRegistro = .Execute
'                    If Not adoRegistro.EOF Then
'
'                        strDefectTipoCuenta = ""
'                        strDefectCodBanco = ""
'                        strDefectNumCuenta = ""
'                        strIndCuentaCorriente = ""
'                        strIndCuentaAhorros = ""
'
'                        If Trim(adoRegistro("IndCuentaCorriente")) = "X" Or Trim(adoRegistro("IndCuentaAhorros")) = "X" Or Trim(adoRegistro("IndAfectaCtasBanco")) = "X" Or Trim(adoRegistro("IndBCR")) = "X" Or Trim(adoRegistro("IndExterior")) = "X" Or Trim(adoRegistro("IndTransferSimple")) = "X" Then
'
'                            cboBancoOrigen.Enabled = True
'                            cboCtaBancariaOrig.Enabled = True
'
'                            strIndCuentaCorriente = Trim(adoRegistro("IndCuentaCorriente"))
'                            strIndCuentaAhorros = Trim(adoRegistro("IndCuentaAhorros"))
'
'                            If Trim(adoRegistro("IndCuentaCorriente")) = "X" Or Trim(adoRegistro("IndCuentaAhorros")) = "X" Then
'                                Call FP_ObtenerCuentaDefecto(strCodMonedaFP, adoRegistro, strDefectTipoCuenta, strDefectCodBanco, strDefectNumCuenta)
'                            End If
'                            Call FP_CargarBancos
'
'                        Else
'
'                            cboBancoOrigen.ListIndex = -1
'                            cboBancoOrigen.Enabled = False
'                            cboCtaBancariaOrig.ListIndex = -1
'                            cboCtaBancariaOrig.Enabled = False
'
'                        End If
'
'                    End If
'                    adoRegistro.Close: Set adoRegistro = Nothing
'
'                End With
'
'            End If
'
'        Else
'
'            MsgBox "No pueden agregar más formas de pago.", vbCritical, Me.Caption
'            cboFormaPago.ListIndex = 0
'            Exit Sub
'
'        End If
'
'End If
'
'End Sub
'
'Private Sub FP_CargarBancos()
'
'Dim strSQL              As String
'Dim adoRegistro         As ADODB.Recordset
'Dim intRegistro         As Integer
'
'    If Trim(strDefectTipoCuenta) <> "" Then
'        'Cargar en el combo los bancos con que se tienen cuenta corriente o ahorros (primero se evalúa este tipo de cuentas
'        cboBancoOrigen.Enabled = True
'        cboCtaBancariaOrig.Enabled = True
'
'        'Obtener lista de bancos con cuenta
'        strSQL = "SELECT DISTINCT(IP.CodPersona + IP.TipoPersona) CODIGO,IP.DescripPersona DESCRIP FROM InstitucionPersona IP JOIN BancoCuenta BC ON (BC.CodBanco = IP.CodPersona)  WHERE BC.CodFondo = '" & strCodFondo & "' AND BC.CodAdministradora = '" & gstrCodAdministradora & "' AND IP.TipoPersona='" & Codigo_Tipo_Persona_Emisor & "' AND IP.IndVigente='X' AND IP.IndBanco='X' AND BC.TipoCuenta = '" & strDefectTipoCuenta & "' ORDER BY IP.DescripPersona"
'        CargarControlLista strSQL, cboBancoOrigen, arrBcoOrigen(), Sel_Defecto
'
'    Else  'cargar todos los bancos si la forma de pago afecta a bancos pero no es cta corriente ni ahorros
'
'        strSQL = "SELECT (CodPersona + TipoPersona) CODIGO,DescripPersona DESCRIP FROM InstitucionPersona WHERE TipoPersona='" & Codigo_Tipo_Persona_Emisor & "' AND IndVigente='X' AND IndBanco='X' ORDER BY DescripPersona"
'        CargarControlLista strSQL, cboBancoOrigen, arrBcoOrigen(), Sel_Defecto
'        If cboBancoOrigen.ListCount > 0 Then cboBancoOrigen.ListIndex = 0
'
'    End If
'
'    'Seleccionar el Banco por defecto si lo hay
'    intRegistro = ObtenerItemLista(arrBcoOrigen(), strDefectCodBanco + Codigo_Tipo_Persona_Emisor)
'    If intRegistro >= 0 Then cboBancoOrigen.ListIndex = intRegistro
'
'
'End Sub
'
'
'Private Sub FP_AdicionarFormaPago(indNuevaFP As Boolean, strModo As String)
''======================================================
''Parámetro
''indNuevaFP = true Insertar
''indNuevaFP = false Modificar
''======================================================
'Dim strMsgError As String
'Dim strModoPago As String
'
'On Error GoTo err
'
'If strModo = "I" Then
'
'    If indNuevaFP = True Then
'        rsg.AddNew
'    End If
'
'    '************ INICIO ****************
'
'    'Asignando a la grilla los valores ingresados en los combos y textbox
'    If indNuevaFP = True Then
'        rsg.Fields("Item") = rsg.RecordCount
'    End If
'    rsg.Fields("CodModoFPago") = strCodModoFPago
'    rsg.Fields("ModoFormaPago") = strModoFPago
'    rsg.Fields("ClaseFormaPago") = strClaseFPagoIni
'
'    rsg.Fields("CodFormaPago") = arrFormaPago(cboFormaPago.ListIndex)
'    rsg.Fields("DesFormaPago") = cboFormaPago.Text
'    rsg.Fields("CodMonedaPago") = arrMonedaFP(cboMonedaFP.ListIndex)
'    rsg.Fields("DesMonedaPago") = cboMonedaFP.Text
'    rsg.Fields("MontoMonedaPago") = CDbl(txtMontoMonedaPago.Text)
'    rsg.Fields("ValorTipoCambio") = CDbl(txtTipoCambioFP.Text)
'    rsg.Fields("MontoMonedaOpe") = CDbl(txtMontoFPago.Text)
'    If cboBancoOrigen.ListIndex > 0 Then
'        rsg.Fields("CodBanco") = arrBcoOrigen(cboBancoOrigen.ListIndex)
'    Else
'        rsg.Fields("CodBanco") = ""
'    End If
'    rsg.Fields("DesBanco") = cboBancoOrigen.Text
'    If cboCtaBancariaOrig.ListIndex > 0 Then
'        rsg.Fields("CodCta") = arrCtaBancariaOrig(cboCtaBancariaOrig.ListIndex)
'        rsg.Fields("NumCuenta") = traerCampo("BancoCuenta", "NumCuenta", "CodFondo", strCodFondo, " AND CodAdministradora = '" & gstrCodAdministradora & "' AND CodFile = '" & Mid(arrCtaBancariaOrig(cboCtaBancariaOrig.ListIndex), 7, 3) & "' AND CodAnalitica = '" & Mid(arrCtaBancariaOrig(cboCtaBancariaOrig.ListIndex), 10, 8) & "'")
'    Else
'        rsg.Fields("CodCta") = ""  '0
'        rsg.Fields("NumCuenta") = ""
'    End If
'    rsg.Fields("DescripCuenta") = cboCtaBancariaOrig.Text
'    rsg.Fields("AnombreDe") = Mid(txtAnombreOrig.Text, 1, 50)
'    rsg.Fields("Contacto") = txtContacto.Text
'    rsg.Fields("Fax") = txtFax.Text
'    rsg.Fields("Observaciones") = Mid(txtObservaciones.Text, 1, 100)
'    rsg.Fields("CodMonedaParDefecto") = strCodMonedaParPorDefecto
'
'    Set gFPagoIni.DataSource = Nothing
'    mostrarDatosGridSQL gFPagoIni, rsg, strMsgError
'
'    dblSumaFPCnt = dblSumaFPCnt + CDbl(txtMontoFPago.Text)    'CDbl(gFPagoIni.Columns.ColumnByFieldName("MontoMonedaOpe").Value)
'    txtSumaInicio.Text = dblSumaFPCnt
'
'End If
'
'indActualizaFP = False
'
'If strMsgError <> "" Then GoTo err
'Exit Sub
'
'err:
'    If strMsgError = "" Then strMsgError = err.Description
'    MsgBox strMsgError, vbInformation, App.Title
'
'End Sub
'
'Private Function FP_CalculaMontoEquiv(dblMontoPago As Double, dblTC As Double, strMonedaOpe As String, strMonedaFP As String) As Double
'
'FP_CalculaMontoEquiv = 0#
'
'    If strMonedaOpe = strMonedaFP Then
'        FP_CalculaMontoEquiv = dblMontoPago
'    Else
'        If strMonedaOpe = Codigo_Moneda_Local Then
'                FP_CalculaMontoEquiv = dblMontoPago * dblTC
'        Else
'            If strMonedaOpe = Codigo_Moneda_Dolar_Americano Then
'                If strMonedaOpe <> Codigo_Moneda_Euro Then
'                    If dblTC <> 0# Then
'                        FP_CalculaMontoEquiv = dblMontoPago / dblTC
'                    End If
'                Else
'                    FP_CalculaMontoEquiv = dblMontoPago * dblTC
'                End If
'            Else 'Operación ingresada en otra moneda del exterior
'                If strMonedaFP = Codigo_Moneda_Dolar_Americano Then
'                    If dblTC <> 0# Then
'                        FP_CalculaMontoEquiv = dblMontoPago / dblTC
'                    End If
'                Else
'                    MsgBox "No se ha definido más cálculos de arbitrajes. Consulte esta operación.", vbCritical
'                End If
'            End If
'
'        End If
'
'    End If
'
'End Function
'Private Function FP_CalculaMontoFP(gFPago As dxDBGrid, dblMontoTotal As Double) As Double
''=====================================
''Parámetros
''strModoFP      Modalidad de la forma de pago: inicio, vencimiento, etc. Se trabajará con 2 dígitos
''dblMontoTotal  Importe total al que no debe exceder la suma de formas de pago según modalidad
''=====================================
'
'Dim dblMontoSuma As Double
'Dim lngRegActual As Long
'Dim i As Integer
'
'    FP_CalculaMontoFP = 0
'    dblMontoSuma = 0
'
'    'Recorrer la grilla para determinar el monto faltante
'
'    If gFPago.Count >= 1 Then
'
'        lngRegActual = CLng(gFPago.Columns.ColumnByFieldName("Item").Value)
'        gFPago.Dataset.First
'
'        For i = 1 To (gFPago.Count)
'            gFPago.Dataset.RecNo = i
'            gFPago.Dataset.RecNo = i
'            If gFPago.Columns.ColumnByFieldName("CodFormaPago").Value <> "" Then
'                dblMontoSuma = dblMontoSuma + CDbl(gFPago.Columns.ColumnByFieldName("MontoMonedaOpe").Value)
'            End If
'        Next i
'
'        gFPago.Dataset.RecNo = lngRegActual
'
'    End If
'
'    FP_CalculaMontoFP = Round(dblMontoTotal - dblMontoSuma, 2)
'
'End Function
'
'Private Sub FP_EliminarFormaPago(gFPago As dxDBGrid, strModo As String)
'
'Dim strMsgError As String
'Dim i As Integer
'
'On Error GoTo err
'
'    If gFPago.Columns.ColumnByFieldName("Item").Value = 0 Or gFPago.Count = 0 Then
'        strMsgError = "No existen formas de pago para eliminar."
'        GoTo err
'    End If
'
'    If strModo = "I" Then
'        dblSumaFPCnt = dblSumaFPCnt - CDbl(gFPago.Columns.ColumnByFieldName("MontoMonedaOpe").Value)
'        txtSumaInicio.Text = dblSumaFPCnt
'    End If
'
'    gFPago.Dataset.Delete
'    If gFPago.Count > 0 Then
'        gFPago.Dataset.First
'    End If
'
'    'Actualizar el número de óndice de la grilla (Item)
'    Do While Not gFPago.Dataset.EOF
'
'        If gFPago.Columns.ColumnByFieldName("Item").Value > 0 Then
'            i = i + 1
'            gFPago.Dataset.Edit
'            gFPago.Columns.ColumnByFieldName("Item").Value = i
'            gFPago.Dataset.Post
'        End If
'        gFPago.Dataset.Next
'
'    Loop
'
'    If gFPago.Dataset.State = dsEdit Or gFPago.Dataset.State = dsInsert Then
'        gFPago.Dataset.Post
'    End If
'
'
'Exit Sub
'
'err:
'If strMsgError = "" Then strMsgError = err.Description
'MsgBox strMsgError, vbInformation, App.Title
'
'End Sub
'
'Private Sub FP_InicializaGrillaFPago()
'Dim strMsgError As String
'
'On Error GoTo err
'
'        '***********************************************
'        'Configurando la grillas de forma de pago INICIO
'        '***********************************************
'        ConfGrid gFPagoIni, True, False, False, False
'
'        Set rsg = Nothing
'        Set gFPagoIni.DataSource = Nothing
'
'        rsg.Fields.Append "Item", adInteger, , adFldRowID                           'Una columna que contiene el #Reg
'        rsg.Fields.Append "CodModoFPago", adVarChar, 8, adFldRowID                  'CodTipoParametro[char](6)+ CodParametro [char](2)
'        rsg.Fields.Append "ModoFormaPago", adVarChar, 10, adFldIsNullable           'I= al Inicio de la operación) / V= al vencimiento de la operación
'        rsg.Fields.Append "ClaseFormaPago", adVarChar, 1, adFldIsNullable           'R= Formas de pago recibidos  E=Formas de pago entregados
'        '-------------------------------
'        'Forma de pago
'        rsg.Fields.Append "CodFormaPago", adVarChar, 2, adFldRowID                  'CodTipoParametro[char](6)+ CodFormaPago [char](2)
'        rsg.Fields.Append "DesFormaPago", adVarChar, 60, adFldIsNullable            'Visible
'        '-------------------------------
'        'Moneda de la forma de pago
'        rsg.Fields.Append "CodMonedaPago", adVarChar, 2, adFldRowID
'        ''''''''''''''''long de descri MONEDA
'        rsg.Fields.Append "DesMonedaPago", adVarChar, 60, adFldIsNullable
'        rsg.Fields.Append "MontoMonedaPago", adDouble, , adFldUpdatable                  'Visible
'        rsg.Fields.Append "ValorTipoCambio", adDouble, , adFldUpdatable
'        rsg.Fields.Append "MontoMonedaOpe", adDouble, , adFldUpdatable                  'Visible
'        '-------------------------------
'        'Banco origen
'        rsg.Fields.Append "CodBanco", adVarChar, 10, adFldRowID               'CodBancoOrigen[char](8)+TipoPerBcoOrigen[char] (2)
'        rsg.Fields.Append "DesBanco", adVarChar, 80, adFldRowID
'        'Cuenta origen
'        rsg.Fields.Append "CodCta", adVarChar, 17, adFldRowID                 'CodFileCtaOrigen[char](3)+CodAnalCtaOrigen[char](8)
'        rsg.Fields.Append "NumCuenta", adVarChar, 15, adFldUnknownUpdatable
'        rsg.Fields.Append "DescripCuenta", adVarChar, 85, adFldUnknownUpdatable   'Visible
'        'A nombre origen de
'        rsg.Fields.Append "AnombreDe", adVarChar, 50, adFldUpdatable            'Visible
'        '--------------------------------
'        rsg.Fields.Append "Contacto", adVarChar, 80, adFldUpdatable                 'Visible
'        rsg.Fields.Append "Fax", adVarChar, 15, adFldUpdatable                      'Visible
'        rsg.Fields.Append "Observaciones", adVarChar, 100, adFldUpdatable           'Visible
'        rsg.Fields.Append "CodMonedaParDefecto", adVarChar, 4, adFldUpdatable       'NO Visible
'
'        If rsg.State <> 1 Then  'Si no está abierto el recordset
'            rsg.Open
'        End If
'
'        indInicializaGrilla = True
'
'Exit Sub
'err:
'If strMsgError = "" Then
'    strMsgError = err.Description
'    MsgBox strMsgError, vbCritical, Me.Caption
'End If
'End Sub
'
'Private Sub FP_LimpiaCampos()
'
'    indActualizaFP = False
'    cmdAccionFP(3).Visible = False
'    txtMontoMonedaPago.Text = 0
'    txtMontoFPago.Text = 0
'    cboModoFPago.ListIndex = 1     '0
'    cboFormaPago.ListIndex = 0
'    cboBancoOrigen.ListIndex = -1  '0
'    cboCtaBancariaOrig.Clear
'    txtAnombreOrig.Text = ""
'    txtContacto.Text = ""
'    txtFax.Text = ""
'    txtObservaciones.Text = ""
'    lblMonedasArbitraje.Visible = False
'    lblTC.Visible = False
'    lblMontoMonedaOpe.Visible = False
'    lblMoneda2Arbitraje.Visible = False
'    txtTipoCambioFP.Visible = False
'    txtMontoFPago.Visible = False
'
'End Sub
'
'Private Sub FP_ObtenerCuentaDefecto(strMoneda As String, adoParRegistro As ADODB.Recordset, ByRef strDfTipoCuenta As String, ByRef strDfCodBanco As String, ByRef strDfNumCuenta As String)
'
'Dim adoRegistro As ADODB.Recordset
'
'    strDfTipoCuenta = ""
'    strDfCodBanco = ""
'    strDfNumCuenta = ""
'
'   'Jalar la cuenta y el banco por defecto
'    strSQL = "SELECT TipoCuenta, CodBanco, NumCuenta FROM BancoCuenta JOIN InstitucionPersona IP ON (CodPersona = CodBanco AND TipoPersona = '02') WHERE CodFondo = '" & strCodFondo & "' AND CodAdministradora = '" & gstrCodAdministradora & "' AND IndCuentaDefecto = 'X' AND CodMoneda = '" & strMoneda & "' AND IP.CodPais = '001'"
'
'    If Trim(adoParRegistro("IndCuentaCorriente")) = "X" Then
'        strSQL = strSQL & " AND TipoCuenta = '" & Codigo_Tipo_Cuenta_Corriente & "'"
'    Else
'        If Trim(adoParRegistro("IndCuentaAhorros")) = "X" Then
'            strSQL = strSQL & " AND TipoCuenta = '" & Codigo_Tipo_Cuenta_Ahorro & "'"
'        End If
'    End If
'
'    Set adoRegistro = New ADODB.Recordset
'    With adoComm
'    .CommandText = strSQL
'    Set adoRegistro = .Execute
'    If Not adoRegistro.EOF Then
'        strDfTipoCuenta = Trim(adoRegistro("TipoCuenta"))
'        strDfCodBanco = Trim(adoRegistro("CodBanco"))
'        strDfNumCuenta = Trim(adoRegistro("NumCuenta"))
'    End If
'
'    adoRegistro.Close: Set adoRegistro = Nothing
'
'    End With
'
'End Sub
'
'Private Function FP_ReqFormasPagoOK() As Boolean
'' Campos requeridos para poder habilitar el tab de las formas de pago
'
'FP_ReqFormasPagoOK = False
'
' '- Revisando si están ingresados los Datos de la confirmación -------------------------
'
'    If CDbl(lblTotal.Caption) = 0 Then
'        MsgBox "Debe indicar el monto para la confirmación del dividendo.", vbCritical, Me.Caption
'        tabEvento.Tab = 1
'        txtValor.SetFocus
'        Exit Function
'    End If
'
''*** Si todo paso OK ***
'FP_ReqFormasPagoOK = True
'
'tabEvento.TabEnabled(2) = True
'tabEvento.Tab = 2
'
'End Function
'
'Private Function FP_ValidarDatosFPago(blnIndNew As Boolean, strModoFPago As String) As Boolean
'
'Dim dblMontoRestanteFP As Double
'
'FP_ValidarDatosFPago = False
'
'    'Validar que se haya seleccionado un modo de forma de pago (inicio, vencimiento)
'    If cboModoFPago.ListIndex <= 0 Then
'        MsgBox "No se ha seleccionado la modalidad de la forma de pago.", vbCritical, Me.Caption
'        cboModoFPago.SetFocus
'        Exit Function
'    End If
'
'    'Validar que se haya seleccionado una forma de pago
'    If cboFormaPago.ListIndex <= 0 Then
'        MsgBox "No se ha seleccionado una forma de pago.", vbCritical, Me.Caption
'        cboFormaPago.SetFocus
'        Exit Function
'    End If
'
'     If txtMontoMonedaPago.Text = "" Or txtMontoMonedaPago.Text = 0 Then
'        MsgBox "Debe ingresar el monto de la forma de pago.", vbCritical, Me.Caption
'        txtMontoMonedaPago.SetFocus
'        Exit Function
'     End If
'
'    If (cboMonedaFP.ListIndex <> strCodMoneda) And (txtTipoCambioFP.Value = 0#) Then
'        MsgBox "Debe ingresar el tipo de cambio si la forma de pago está en otra moneda.", vbCritical, Me.Caption
'        txtTipoCambioFP.SetFocus
'        Exit Function
'    End If
'
'    'Validar que si es cuenta corriente o de ahorros indique la cuenta
'    If cboFormaPago.ListIndex = Forma_Pago_CuentaCorriente Or cboFormaPago.ListIndex = Forma_Pago_CuentaAhorros Then
'        If cboCtaBancariaOrig.ListIndex < 1 Then
'            MsgBox "Debe indicar el número de cuenta.", vbCritical, Me.Caption
'            cboCtaBancariaOrig.SetFocus
'            Exit Function
'        End If
'    End If
'
'    'Validar que el importe no exceda a lo permitido por la operación
'    If cboFormaPago.ListIndex > 0 Then
'
'        If strModoFPago = "I" Then    'Forma de pago al inicio o contado
'           If blnIndNew = True Then   'si es un nuevo ingreso
'              dblMontoRestanteFP = FP_CalculaMontoFP(gFPagoIni, CDbl(lblTotal.Caption))
'            Else 'si es una modificación
'                    dblMontoRestanteFP = FP_CalculaMontoFP(gFPagoIni, CDbl(lblTotal.Caption)) + CDbl(gFPagoIni.Columns.ColumnByFieldName("MontoMonedaOpe").Value)
'            End If
'        End If
'
'        If CDbl(txtMontoFPago.Text) > dblMontoRestanteFP Then
'            MsgBox "El monto de la forma de pago supera a lo permitido por la operación.", vbCritical, Me.Caption
'            txtMontoMonedaPago.SetFocus
'            Exit Function
'        End If
'    End If
'
'
'FP_ValidarDatosFPago = True
'
'End Function
'
'
'Private Sub cboModoFPago_Click()
'
'Dim adoRegistro As ADODB.Recordset
'Dim strCodFPagoDefecto As String
'Dim intRegistro As Integer
'
'    strModoFPago = Valor_Caracter
'    strCodModoFPago = Valor_Caracter
'
'    If cboModoFPago.ListIndex < 0 Then Exit Sub
'
'    strCodModoFPago = Trim(arrModoFPago(cboModoFPago.ListIndex))
'
'    If cboModoFPago.ListIndex > 0 Then
'
'        If Trim(Mid(arrModoFPago(cboModoFPago.ListIndex), 7, 2)) = "01" Then    'Forma de pago al inicio o contado
'           If lblTotal.Caption <> "" Then
'                If CDbl(lblTotal.Caption) <= dblSumaFPCnt Then
'                     MsgBox "Ya no es posible agregar más formas de pago al inicio de la operación.", vbCritical, Me.Caption
'                     cboModoFPago.ListIndex = 0
'                     Exit Sub
'                End If
'           End If
'        End If
'
'        'Obtener el ValorParametro del modo de la forma de pago
'        strModoFPago = Trim(traerCampo("AuxiliarParametro", "ValorParametro", "CodTipoParametro", "EJEPAG", "AND CodParametro  = '" & Trim(Mid(arrModoFPago(cboModoFPago.ListIndex), 7, 2)) & "'"))
'
'        '--------
'
'        'Clase de forma de pago: recibimos o entregamos. Para dividendos es Recibidos
'        strClaseFPagoIni = Valor_Caracter
'
'        '*** clase de forma de pago al inicio ***
'        With adoComm
'
'            .CommandText = "SELECT DISTINCT ClaseFormaPago FROM InversionFileTipoOperacionNegociacionFormaPago WHERE CodFile = '" & strCodFile & "' AND CodTipoOperacion  = '" & Codigo_Caja_Dividendos & "' AND ModoFormaPago = 'I'"
'            Set adoRegistro = .Execute
'
'            If Not adoRegistro.EOF Then
'                strClaseFPagoIni = Trim(adoRegistro("ClaseFormaPago"))
'            End If
'            adoRegistro.Close: Set adoRegistro = Nothing
'
'        End With
'
'        '--------
'
'        'Cargar las formas de pago según el Tipo de Operación
'        If strModoFPago = "I" Then
'            strSQL = "SELECT FP.DescripFormaPago DESCRIP, IP.CodFormaPago CODIGO from InversionFileTipoOperacionNegociacionFormaPago IP JOIN FormaPago FP ON (FP.CodFormaPago = IP.CodFormaPago) WHERE CodFile = '" & strCodFile & "' AND CodTipoOperacion = '" & Codigo_Caja_Dividendos & "' AND ModoFormaPago = '" & strModoFPago & "' AND ClaseFormaPago = '" & strClaseFPagoIni & "'"
'        End If
'        CargarControlLista strSQL, cboFormaPago, arrFormaPago(), Sel_Defecto
'
'        'Seleccionar la forma de pago por defecto
'        strSQL = strSQL & " AND IndDefecto = 'X'"
'
'        Set adoRegistro = New ADODB.Recordset
'        With adoComm
'            .CommandText = strSQL
'            Set adoRegistro = .Execute
'            If Not adoRegistro.EOF Then
'                strCodFPagoDefecto = Trim(adoRegistro("CODIGO"))
'            End If
'            adoRegistro.Close: Set adoRegistro = Nothing
'        End With
'
'        intRegistro = ObtenerItemLista(arrFormaPago(), strCodFPagoDefecto)
'        If intRegistro >= 0 Then cboFormaPago.ListIndex = intRegistro
'        'FIN: Seleccionar la forma de pago por defecto
'    End If
'
'End Sub

Private Sub cmdAccion_GotFocus()

End Sub

Private Sub Form_Activate()

    frmMainMdi.stbMdi.Panels(3).Text = Me.Caption
    
End Sub

Private Sub Form_Deactivate()

    Call OcultarReportes
    
End Sub

Private Sub Form_Load()

    Call InicializarValores
    Call CargarListas
    Call CargarReportes
    Call Buscar
    'Call DarFormato
        
    Call ValidarPermisoUsoControl(Trim(gstrLoginUS), Me, Trim(App.Title) + Separador_Codigo_Objeto + _
    gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)
    
    CentrarForm Me
        
    Call ValidaExisteTipoCambio(Codigo_TipoCambio_SBS, gstrFechaActual)
    
End Sub

Private Sub DarFormato()

'    Dim intCont As Integer
'
'    For intCont = 0 To (lblDescrip.Count - 1)
'        Call FormatoEtiqueta(lblDescrip(intCont), vbLeftJustify)
'    Next
    
    Dim c As Object
    
    For Each c In Me.Controls
        If TypeOf c Is Label Then
            Call FormatoEtiqueta(c, vbRightJustify)
        End If
    Next
            
End Sub
Private Sub CargarReportes()

'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
'    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Vista Activa"
    
End Sub
Private Sub CargarListas()
    
    '*** Fondos ***
    strSQL = "{ call up_ACSelDatosParametro(74,'" & gstrCodAdministradora & "','" & gstrCodFondoContable & "') }"
    CargarControlLista strSQL, cboFondo, arrFondo(), Valor_Caracter
        
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
    
    '*** Tipo de Evento ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPEVE' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboEvento, arrEvento(), Sel_Defecto
        
    If cboEvento.ListCount > 0 Then cboEvento.ListIndex = 0
    
    '*** Moneda de la Forma de pago ***
    strSQL = "{ call up_ACSelDatos(2) }"
    CargarControlLista strSQL, cboMonedaFP, arrMonedaFP(), Sel_Defecto
    
    '*** Agentes de Bolsa ***
    strSQL = "SELECT CodPersona CODIGO,DescripPersona DESCRIP FROM InstitucionPersona WHERE TipoPersona='" & Codigo_Tipo_Persona_Agente & "' ORDER BY DescripPersona"
    CargarControlLista strSQL, cboAgente, arrAgente(), Sel_Defecto
   
    '*** Carga modo de ejecución de la forma de pago. Para Dividendos es el medio de pago al inicio ***
    strSQL = "SELECT (CodTipoParametro + CodParametro) CODIGO, DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='EJEPAG' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboModoFPago, arrModoFPago(), Sel_Defecto
    If cboModoFPago.ListCount > 0 Then cboModoFPago.ListIndex = 0 '1
    
        
End Sub
Private Sub InicializarValores()
    
    strEstado = Reg_Defecto
    tabEvento.Tab = 0

    lblFechaOperacion = CStr(gdatFechaActual)
    
    tabEvento.TabVisible(2) = False
    
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgConsulta.Columns(0).Width = tdgConsulta.Width * 0.01 * 8
    tdgConsulta.Columns(1).Width = tdgConsulta.Width * 0.01 * 20
    tdgConsulta.Columns(2).Width = tdgConsulta.Width * 0.01 * 15
    tdgConsulta.Columns(3).Width = tdgConsulta.Width * 0.01 * 26
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
                
End Sub
Private Sub Form_Unload(Cancel As Integer)

    Set frmPrecioTir = Nothing
    Call OcultarReportes
    frmMainMdi.stbMdi.Panels(3).Text = "Acción"
    
End Sub

Private Sub ConfirmarEvento()

'    Dim intCont As Integer
'
'    With gadoComando
'
'        .CommandText = "UPDATE tblEntregaEventos SET FLG_STAT='C',"
'        If strEvento = "LA" Then .CommandText = .CommandText & "LIB_REAL=" & CDbl(mhrValor.Text) & ","
'        If strEvento = "DE" Then .CommandText = .CommandText & "DIV_REAL=" & CDbl(mhrValor.Text) & ","
'        If strEvento = "VN" Then .CommandText = .CommandText & "VAL_REAL=" & CDbl(mhrValor.Text) & ","
'        .CommandText = .CommandText & "FCH_ACTU='" & strFchOper & "',"
'        .CommandText = .CommandText & "USR_ACTU='" & gstrUID & "' "
'        .CommandText = .CommandText & "WHERE COD_FOND='" & strFondo & "' AND TIP_ACUE='" & strEvento & "' AND "
'
'        gadoConexion.Execute .CommandText
'
'
'    End With
    
End Sub




Private Sub lblCantAcciones_Change()

    Call FormatoMillarEtiqueta(lblCantAcciones, 0)
    
End Sub

Private Sub lblTotal_Change()

    Call FormatoMillarEtiqueta(lblTotal, Decimales_Monto)

End Sub

Private Sub lblValor_Change()

    Call FormatoMillarEtiqueta(lblValor, Decimales_Monto)
    
End Sub

Private Sub tabEvento_Click(PreviousTab As Integer)

    Select Case tabEvento.Tab
        
        Case 1, 2
            'If PreviousTab = 0 And strEstado = Reg_Consulta Then tabEvento.Tab = 0
            If strEstado = Reg_Defecto Then tabEvento.Tab = 0
           
'            If tabEvento.Tab = 2 Then
'                If cboEvento.ListIndex <> Codigo_Evento_Dividendo Then Exit Sub
'                If FP_ReqFormasPagoOK Then
'                    tabEvento.Tab = 2
'                    If cboFormaPago.ListIndex < 1 Then
'                        cboModoFPago_Click
'                        'Seteando el monto total de la confirmación en pantalla
'                        If strModoFPago = "I" Then    'Forma de pago al inicio o contado
'                                txtMontoTotalOpera.Text = lblTotal.Caption
'                        End If
'                    End If
'                End If
'            End If
            
    End Select

End Sub

Private Sub tdgConsulta_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)

    If ColIndex = 3 Then
        Call DarFormatoValor(Value, Decimales_Monto)
    End If
    
    If ColIndex = 4 Then
        Call DarFormatoValor(Value, Decimales_Monto)
    End If
    
End Sub



Private Sub txtTasaComision_Change()

    'Call FormatoCajaTexto(txtTasaComision, Decimales_Tasa)
    
    Call ActualizaComision(txtTasaComision, txtValorComision)

End Sub

Private Sub txtTasaComision_KeyPress(KeyAscii As Integer)
    
    'Call ValidaCajaTexto(KeyAscii, "M", txtTasaComision, Decimales_Tasa)

    If KeyAscii = vbKeyReturn Then
        Call ActualizaComision(txtTasaComision, txtValorComision)
    End If

End Sub


Private Sub ActualizaComision(ctrlPorcentaje As Control, ctrlComision As Control)

    If Not IsNumeric(ctrlPorcentaje.Value) Then Exit Sub
        
    If ctrlPorcentaje.Value > 0 Then
        ctrlComision.Text = CStr(txtValor.Value * ctrlPorcentaje.Value / 100)
    Else
        ctrlComision.Text = "0"
    End If
            
    Call CalculoTotal
            
End Sub



Private Sub txtValor_Change()

    Call ActualizaComision(txtTasaComision, txtValorComision)

End Sub


Private Sub txtValorComision_KeyPress(KeyAscii As Integer)


    If KeyAscii = vbKeyReturn Then
        Call ActualizaPorcentaje(txtValorComision, txtTasaComision)
    End If

End Sub

Private Sub CalculoTotal()

    Dim curValorComision As Currency, curMonTotal As Currency

    If Not IsNumeric(txtValorComision.Value) Then Exit Sub
    
    curValorComision = txtValorComision.Value
    
    curMonTotal = txtValor.Value - curValorComision
 
    lblTotal.Caption = CStr(curMonTotal)
    
End Sub

Private Sub ActualizaPorcentaje(ctrlComision As Control, ctrlPorcentaje As Control)

    If Not IsNumeric(ctrlComision) Or Not IsNumeric(txtValor) Then Exit Sub
                
    If CCur(txtValor.Text) = 0 Then
        ctrlPorcentaje = "0"
    Else
        If CCur(ctrlComision) > 0 Then
            ctrlPorcentaje = CStr((CCur(ctrlComision) / CCur(txtValor.Text) * 100))
        Else
            ctrlPorcentaje = "0"
        End If
    End If
                
End Sub
