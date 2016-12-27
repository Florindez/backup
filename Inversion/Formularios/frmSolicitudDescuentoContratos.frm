VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{830D5A06-9B70-4F8C-98B6-7A19C4E7760B}#1.0#0"; "TAMControls.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmSolicitudDescuentoContratos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Solicitud de Descuentos de Contratos Futuros"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   11715
   ShowInTaskbar   =   0   'False
   Begin TAMControls2.ucBotonEdicion2 cmdOpcion 
      Height          =   735
      Left            =   390
      TabIndex        =   57
      Top             =   6540
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   1296
      Buttons         =   4
      Caption0        =   "&Nuevo"
      Tag0            =   "0"
      Visible0        =   0   'False
      ToolTipText0    =   "Nuevo"
      Caption1        =   "&Eliminar"
      Tag1            =   "4"
      Visible1        =   0   'False
      ToolTipText1    =   "Eliminar"
      Caption2        =   "&Buscar"
      Tag2            =   "5"
      Visible2        =   0   'False
      ToolTipText2    =   "Buscar"
      Caption3        =   "&Modificar"
      Tag3            =   "3"
      Visible3        =   0   'False
      ToolTipText3    =   "Modificar"
      UserControlWidth=   5700
   End
   Begin TAMControls2.ucBotonEdicion2 cmdAccion 
      Height          =   735
      Left            =   6300
      TabIndex        =   59
      Top             =   6540
      Visible         =   0   'False
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
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   10110
      TabIndex        =   58
      Top             =   6570
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
      Caption0        =   "&Salir"
      Tag0            =   "9"
      Visible0        =   0   'False
      ToolTipText0    =   "Salir"
      UserControlWidth=   1200
   End
   Begin MSAdodcLib.Adodc adoConsulta 
      Height          =   330
      Left            =   8730
      Top             =   6540
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
      Height          =   6345
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   11192
      _Version        =   393216
      Style           =   1
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
      TabPicture(0)   =   "frmSolicitudDescuentoContratos.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdOpcionhidden"
      Tab(0).Control(1)=   "fraCriterio"
      Tab(0).Control(2)=   "tdgConsulta"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Datos de la Solicitud"
      TabPicture(1)   =   "frmSolicitudDescuentoContratos.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraDatosBasicos"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraDatosTitulo"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Negociación"
      TabPicture(2)   =   "frmSolicitudDescuentoContratos.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraDatosNegociacion"
      Tab(2).ControlCount=   1
      Begin TAMControls2.ucBotonEdicion2 cmdOpcionhidden 
         Height          =   735
         Left            =   -74880
         TabIndex        =   60
         Top             =   6960
         Width           =   4200
         _ExtentX        =   7408
         _ExtentY        =   1296
         Buttons         =   3
         Caption0        =   "&Nuevo"
         Tag0            =   "0"
         Visible0        =   0   'False
         ToolTipText0    =   "Nuevo"
         Caption1        =   "&Eliminar"
         Tag1            =   "4"
         Visible1        =   0   'False
         ToolTipText1    =   "Eliminar"
         Caption2        =   "&Buscar"
         Tag2            =   "5"
         Visible2        =   0   'False
         ToolTipText2    =   "Buscar"
         UserControlWidth=   4200
      End
      Begin VB.Frame fraDatosTitulo 
         Caption         =   "Datos de la Solicitud"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2565
         Left            =   120
         TabIndex        =   34
         Top             =   1980
         Width           =   11415
         Begin VB.TextBox txtDescripOrden 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1830
            MaxLength       =   45
            TabIndex        =   37
            Top             =   1530
            Width           =   4170
         End
         Begin VB.ComboBox cboMoneda 
            Height          =   315
            Left            =   8760
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   270
            Width           =   2400
         End
         Begin VB.TextBox txtObservacion 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   825
            Left            =   7890
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   35
            Top             =   1590
            Visible         =   0   'False
            Width           =   3270
         End
         Begin MSComCtl2.DTPicker dtpFechaOrden 
            Height          =   315
            Left            =   1830
            TabIndex        =   38
            Top             =   810
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
            Format          =   175570945
            CurrentDate     =   38776
         End
         Begin MSComCtl2.DTPicker dtpFechaLiquidacion 
            Height          =   315
            Left            =   1830
            TabIndex        =   39
            Top             =   1170
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
            Format          =   175570945
            CurrentDate     =   38776
         End
         Begin MSComCtl2.DTPicker dtpFechaVencimiento 
            Height          =   315
            Left            =   7890
            TabIndex        =   40
            Top             =   1200
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
            Format          =   175570945
            CurrentDate     =   38776
         End
         Begin TAMControls.TAMTextBox txtDiasPlazo 
            Height          =   315
            Left            =   7890
            TabIndex        =   41
            Top             =   810
            Width           =   975
            _ExtentX        =   1720
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
            Container       =   "frmSolicitudDescuentoContratos.frx":0054
            Text            =   "0"
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            EnterTab        =   -1  'True
            Apariencia      =   1
            Borde           =   1
         End
         Begin TAMControls.TAMTextBox txtValorFinanciar 
            Height          =   315
            Left            =   1950
            TabIndex        =   53
            Top             =   270
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
            Container       =   "frmSolicitudDescuentoContratos.frx":0070
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            EnterTab        =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   2
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000010&
            BorderStyle     =   6  'Inside Solid
            Index           =   0
            X1              =   120
            X2              =   11160
            Y1              =   690
            Y2              =   690
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Solicitud"
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
            Left            =   210
            TabIndex        =   49
            Top             =   885
            Width           =   750
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Vigencia"
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
            Left            =   210
            TabIndex        =   48
            Top             =   1215
            Width           =   1335
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
            Left            =   210
            TabIndex        =   47
            Top             =   1560
            Width           =   1005
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Plazo (DIAS)"
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
            Left            =   6150
            TabIndex        =   46
            Top             =   870
            Width           =   1095
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
            Left            =   7020
            TabIndex        =   45
            Top             =   345
            Width           =   690
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones"
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
            Left            =   6120
            TabIndex        =   44
            Top             =   1590
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto a Financiar"
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
            Index           =   113
            Left            =   210
            TabIndex        =   43
            Top             =   330
            Width           =   1545
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Vencimiento"
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
            Index           =   117
            Left            =   6120
            TabIndex        =   42
            Top             =   1215
            Width           =   1635
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
         Height          =   1635
         Left            =   -74880
         TabIndex        =   22
         Top             =   420
         Width           =   11415
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
            Left            =   10020
            Picture         =   "frmSolicitudDescuentoContratos.frx":008C
            Style           =   1  'Graphical
            TabIndex        =   56
            ToolTipText     =   "Enviar a BackOffice"
            Top             =   360
            Width           =   1200
         End
         Begin VB.ComboBox cboFondo 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   360
            Width           =   4785
         End
         Begin VB.ComboBox cboTipoInstrumento 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   720
            Width           =   4785
         End
         Begin VB.ComboBox cboEstado 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   1080
            Width           =   4785
         End
         Begin MSComCtl2.DTPicker dtpFechaOrdenDesde 
            Height          =   315
            Left            =   7590
            TabIndex        =   26
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
            Format          =   175570945
            CurrentDate     =   38785
         End
         Begin MSComCtl2.DTPicker dtpFechaOrdenHasta 
            Height          =   315
            Left            =   7590
            TabIndex        =   27
            Top             =   1080
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
            Format          =   175570945
            CurrentDate     =   38785
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Solicitud"
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
            Left            =   6780
            TabIndex        =   33
            Top             =   405
            Width           =   1335
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
            TabIndex        =   32
            Top             =   405
            Width           =   540
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
            Left            =   6780
            TabIndex        =   31
            Top             =   765
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
            Index           =   21
            Left            =   6780
            TabIndex        =   30
            Top             =   1155
            Width           =   510
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
            TabIndex        =   29
            Top             =   795
            Width           =   1005
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
            TabIndex        =   28
            Top             =   1155
            Width           =   600
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
         Height          =   1545
         Left            =   120
         TabIndex        =   13
         Top             =   420
         Width           =   11415
         Begin VB.ComboBox cboComisionista 
            Height          =   315
            Left            =   6990
            Style           =   2  'Dropdown List
            TabIndex        =   61
            Top             =   1095
            Width           =   4185
         End
         Begin VB.ComboBox cboEmisor 
            Height          =   315
            Left            =   6990
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   720
            Width           =   4185
         End
         Begin VB.ComboBox cboClaseInstrumento 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   1095
            Width           =   4185
         End
         Begin VB.ComboBox cboTipoInstrumentoOrden 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   720
            Width           =   4185
         End
         Begin VB.ComboBox cboFondoOrden 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   330
            Width           =   4185
         End
         Begin TAMControls.TAMTextBox txtNum_Solicitud 
            Height          =   315
            Left            =   9660
            TabIndex        =   52
            Top             =   330
            Width           =   1485
            _ExtentX        =   2619
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
            Container       =   "frmSolicitudDescuentoContratos.frx":05E7
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            EnterTab        =   -1  'True
            Apariencia      =   1
            Borde           =   1
         End
         Begin VB.Label lblComisionistaInversion 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Comisionista"
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
            Left            =   5850
            TabIndex        =   62
            Top             =   1170
            Width           =   1095
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "N° Solicitud"
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
            Left            =   8550
            TabIndex        =   51
            Top             =   390
            Width           =   1005
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
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
            Left            =   5850
            TabIndex        =   21
            Top             =   795
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
            Index           =   1
            Left            =   210
            TabIndex        =   20
            Top             =   780
            Width           =   1005
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
            Left            =   210
            TabIndex        =   19
            Top             =   420
            Width           =   540
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
            Left            =   210
            TabIndex        =   18
            Top             =   1170
            Width           =   480
         End
      End
      Begin VB.Frame fraDatosNegociacion 
         Caption         =   "Negociación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2595
         Left            =   -74880
         TabIndex        =   1
         Top             =   420
         Width           =   4155
         Begin VB.ComboBox cboSubClaseInstrumento 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   2100
            Width           =   1905
         End
         Begin VB.ComboBox cboTipoTasa 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   630
            Width           =   1900
         End
         Begin VB.ComboBox cboBaseAnual 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1005
            Width           =   1900
         End
         Begin VB.TextBox txtTasa 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   2040
            MaxLength       =   45
            TabIndex        =   3
            Top             =   270
            Width           =   1900
         End
         Begin VB.ComboBox cboCobroInteres 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1365
            Width           =   1900
         End
         Begin TAMControls.TAMTextBox txtValorNominal 
            Height          =   315
            Left            =   2040
            TabIndex        =   6
            Top             =   1740
            Width           =   1905
            _ExtentX        =   3360
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
            Container       =   "frmSolicitudDescuentoContratos.frx":0603
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            EnterTab        =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   2
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
            Left            =   240
            TabIndex        =   12
            Top             =   2190
            Width           =   810
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Tasa"
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
            Index           =   41
            Left            =   240
            TabIndex        =   11
            Top             =   675
            Width           =   870
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Base Anual"
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
            Index           =   40
            Left            =   240
            TabIndex        =   10
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tasa Facial"
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
            Left            =   240
            TabIndex        =   9
            Top             =   330
            Width           =   1005
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto Aprobado"
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
            Left            =   240
            TabIndex        =   8
            Top             =   1830
            Width           =   1410
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Modo cobro Interés"
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
            Index           =   110
            Left            =   240
            TabIndex        =   7
            Top             =   1425
            Width           =   1650
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmSolicitudDescuentoContratos.frx":061F
         Height          =   4095
         Left            =   -74880
         OleObjectBlob   =   "frmSolicitudDescuentoContratos.frx":0639
         TabIndex        =   55
         Top             =   2130
         Width           =   11415
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   35
         Left            =   -67920
         TabIndex        =   50
         Top             =   5400
         Visible         =   0   'False
         Width           =   405
      End
   End
End
Attribute VB_Name = "frmSolicitudDescuentoContratos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrFondo()               As String, arrFondoOrden()              As String
Dim arrTipoInstrumento()     As String, arrTipoInstrumentoOrden()    As String
Dim arrEstado()              As String
Dim arrEmisor()              As String, arrMoneda()                  As String
Dim arrBaseAnual()           As String, arrTipoTasa()                As String
Dim arrClaseInstrumento()    As String
Dim arrSubClaseInstrumento() As String, arrComisionista()            As String
Dim strCodFondo              As String, strCodFondoOrden             As String
Dim strCodTipoInstrumento    As String, strCodTipoInstrumentoOrden   As String
Dim strCodEstado             As String, strCodTipoSolicitud          As String
Dim strCodEmisor             As String, strCodMoneda                 As String
Dim strCodBaseAnual          As String, strCodTipoTasa               As String
Dim strCodClaseInstrumento   As String
Dim strCodTitulo             As String, strCodSubClaseInstrumento    As String

Dim strEstado                As String, strSQL                       As String
Dim strResponsablePagoCancel As String
Dim arrPagoInteres()         As String

Dim strCodFile               As String, strCodAnalitica              As String

Dim strEstadoSolicitud       As String
Dim strCodRiesgo             As String, strCodSubRiesgo              As String
Dim strIndPacto              As String
Dim strIndNegociable         As String, strCodigosFile               As String
Dim strCodCobroInteres       As String, strViaCobranza               As String

Dim intDiasAdicionales       As Integer

Dim datFechaVctoAdicional    As Date

Dim indCargadoDesdeBandeja   As Boolean
Dim blnCargarCabeceraAnexo   As Boolean
Dim blnCancelaPrepago        As Boolean

Dim strPersonalizaComision   As String, strCodComisionista            As String
Dim numSecCondicion           As Integer

Dim blnFlag                  As Boolean
Dim strCodFondoSol           As String

Public Sub Adicionar()

    If Not EsDiaUtil(gdatFechaActual) Then
        MsgBox "No se puede negociar en un día no útil !", vbCritical, Me.Caption
        Exit Sub
    End If
        
    If cboTipoInstrumento.ListCount > 1 Then
        frmMainMdi.stbMdi.Panels(3).Text = "Adicionar Orden..."
                
        strEstado = Reg_Adicion
              
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        
        If blnCargarCabeceraAnexo = False Then Call HabilitaCombos(True)

        With tabRFCortoPlazo
            .TabEnabled(0) = False
            .TabEnabled(1) = True
            .TabEnabled(2) = False
            .Tab = 1
        End With
        
        cmdOpcion.Visible = False
        cmdAccion.Visible = True
       
    Else
        MsgBox "Acceso a Negociación Denegada", vbCritical, Me.Caption
    End If
    
End Sub

Private Sub LlenarFormulario(ByVal strModo As String, _
                             Optional ByVal strParNumSolicitud As String = "")

    Dim adoRecord       As ADODB.Recordset
    
    Dim intRegistro     As Integer
    
    Dim strNumSolicitud As String
    
    Select Case strModo

        Case Reg_Adicion
        
            txtNum_Solicitud.Text = Valor_Caracter

            intRegistro = ObtenerItemLista(arrFondoOrden(), strCodFondo)

            If intRegistro >= 0 Then cboFondoOrden.ListIndex = intRegistro
        
            txtTasa.Text = "0"
            
            cboBaseAnual.ListIndex = -1

            If cboBaseAnual.ListCount > 0 Then cboBaseAnual.ListIndex = 0
                
            cboTipoTasa.ListIndex = -1

            If cboTipoTasa.ListCount > 0 Then cboTipoTasa.ListIndex = 0
           
            intRegistro = ObtenerItemLista(arrMoneda(), strCodMoneda)

            If intRegistro >= 0 Then cboMoneda.ListIndex = intRegistro
            
            dtpFechaOrden.Value = gdatFechaActual
            dtpFechaLiquidacion.Value = dtpFechaOrden.Value
            
            txtDiasPlazo.Text = "0"
            
            txtDescripOrden.Text = Valor_Caracter

            txtObservacion.Text = Valor_Caracter
            txtValorFinanciar.Text = 0#
            txtValorNominal.Text = "0"
    
            dtpFechaVencimiento.Value = gdatFechaActual
    
        Case Reg_Edicion
            Set adoRecord = New ADODB.Recordset
            
            If strParNumSolicitud = "" Then
                strNumSolicitud = Trim$(tdgConsulta.Columns(0).Value)
            Else
                strNumSolicitud = strParNumSolicitud
            End If

            adoComm.CommandText = "SELECT CodFondo, CodAdministradora  ,NumSolicitud   ,FechaSolicitud    ,CodTitulo" & ",EstadoSolicitud   ,CodFile            ,CodAnalitica   ,CodDetalleFile    ,CodSubDetalleFile" & ",TipoSolicitud     ,DescripSolicitud   ,CodEmisor, CodComisionista,NumSecuencialComisionistaCondicion      ,FechaConfirmacion ,FechaVencimiento" & ",FechaLiquidacion  ,FechaEmision       ,CodMoneda      ,ValorTipoCambio   ,MontoSolicitud" & ",MontoAprobado     ,TipoTasa           ,BaseAnual      ,TasaInteres       ,Observacion " & "FROM InversionSolicitud WHERE CodFondo = '" & strCodFondoSol & "' AND CodAdministradora = '" & gstrCodAdministradora & "' AND NumSolicitud='" & strNumSolicitud & "'"
            Set adoRecord = adoComm.Execute

            If Not adoRecord.EOF Then

                txtNum_Solicitud.Text = strNumSolicitud

                intRegistro = ObtenerItemLista(arrFondoOrden(), adoRecord.Fields("CodFondo"))

                If intRegistro >= 0 Then cboFondoOrden.ListIndex = intRegistro
            
                intRegistro = ObtenerItemLista(arrTipoInstrumentoOrden(), adoRecord.Fields("CodFile"))

                If intRegistro >= 0 Then cboTipoInstrumentoOrden.ListIndex = intRegistro
                                        
                intRegistro = ObtenerItemLista(arrClaseInstrumento(), adoRecord.Fields("CodDetalleFile"))

                If intRegistro >= 0 Then cboClaseInstrumento.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrEmisor(), adoRecord.Fields("CodEmisor"))

                If intRegistro >= 0 Then cboEmisor.ListIndex = intRegistro

                txtTasa.Text = adoRecord.Fields("TasaInteres")
                
                intRegistro = ObtenerItemLista(arrBaseAnual(), adoRecord.Fields("BaseAnual"))

                If intRegistro >= 0 Then cboBaseAnual.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrTipoTasa(), adoRecord.Fields("TipoTasa"))

                If intRegistro >= 0 Then cboTipoTasa.ListIndex = intRegistro

                intRegistro = ObtenerItemLista(arrMoneda(), adoRecord.Fields("CodMoneda"))

                If intRegistro >= 0 Then cboMoneda.ListIndex = intRegistro
                
                'intRegistro = ObtenerItemLista(arrComisionista(), adoRecord.Fields("CodComisionista"))

                'If intRegistro >= 0 Then cboComisionista.ListIndex = intRegistro
                
                
                strCodAnalitica = adoRecord.Fields("CodAnalitica")
                
                dtpFechaOrden.Value = adoRecord.Fields("FechaSolicitud")
                dtpFechaLiquidacion.Value = adoRecord.Fields("FechaLiquidacion")
                dtpFechaVencimiento.Value = adoRecord.Fields("FechaVencimiento")
                '
                txtDescripOrden.Text = adoRecord.Fields("DescripSolicitud")
                '
                txtValorFinanciar.Text = adoRecord.Fields("MontoSolicitud")
                txtValorNominal.Text = adoRecord.Fields("MontoAprobado")
                txtDiasPlazo.Text = adoRecord.Fields("FechaVencimiento") - adoRecord.Fields("FechaSolicitud")
                    
            End If

            adoRecord.Close: Set adoRecord = Nothing
            
    End Select
    
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
        
        strMensaje = "Se procederá a eliminar la ORDEN " & tdgConsulta.Columns(0) & vbNewLine & vbNewLine & vbNewLine & "¿ Seguro de continuar ?"
        
        If MsgBox(strMensaje, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
    
            '*** Anular Orden ***
            adoComm.CommandText = "UPDATE InversionSolicitud SET EstadoSolicitud='" & Estado_Solicitud_Flujos_Anulada & "' WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND  NumSolicitud='" & Trim$(tdgConsulta.Columns(0)) & "'"
                
            adoConn.Execute adoComm.CommandText
            
            '*** Anular Título si corresponde ***
            adoComm.CommandText = "UPDATE InstrumentoInversion SET IndVigente='' WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & "CodTitulo='" & Trim$(tdgConsulta.Columns(2)) & "'"
                
            adoConn.Execute adoComm.CommandText
            
            MsgBox Mensaje_Eliminacion_Exitosa, vbExclamation, Me.Caption
            
            tabRFCortoPlazo.TabEnabled(0) = True
            tabRFCortoPlazo.Tab = 0
            Call Buscar
            
            Exit Sub
        End If
    End If
    
End Sub

Public Sub Grabar()

    Call Accion(vSave)

End Sub

Public Sub GrabarNew()

    Dim adoRegistro         As ADODB.Recordset
    Dim strFechaSolicitud   As String, strFechaLiquidacion      As String
    Dim strFechaVencimiento As String
    Dim intAccion           As Integer
    Dim lngNumError         As Long
    Dim dblTasaInteres      As Double
    
    Dim strMsgError         As String
    On Error GoTo CtrlError

    If strEstado = Reg_Consulta Then Exit Sub
    
    If strEstado = Reg_Adicion Or strEstado = Reg_Edicion Then
        If TodoOK() Then
        
            strEstadoSolicitud = "01"
            
            Me.MousePointer = vbHourglass
            
            strFechaSolicitud = Convertyyyymmdd(dtpFechaOrden.Value)
            strFechaLiquidacion = Convertyyyymmdd(dtpFechaLiquidacion.Value)
            strFechaVencimiento = Convertyyyymmdd(dtpFechaVencimiento.Value)
           
            Set adoRegistro = New ADODB.Recordset

            '*** Guardar Orden de Inversion ***
            With adoComm

                dblTasaInteres = CDbl(txtTasa.Text)

                .CommandText = "{ call up_IVAdicInversionSolicitud('" & strCodFondoOrden & "','" & gstrCodAdministradora & "','" & _
                                txtNum_Solicitud.Text & "','" & strFechaSolicitud & "','" & strEstadoSolicitud & "','" & strCodFile & "','" & _
                                strCodAnalitica & "','" & strCodClaseInstrumento & "','" & strCodSubClaseInstrumento & "','" & _
                                strCodTipoSolicitud & "','" & Trim$(txtDescripOrden.Text) & "','" & strCodEmisor & "','" & _
                                strCodComisionista & "'," & numSecCondicion & ",'" & _
                                strFechaLiquidacion & "','" & strFechaVencimiento & "','" & strFechaLiquidacion & "','" & _
                                strFechaSolicitud & "','" & strCodMoneda & "'," & 0 & "," & txtValorFinanciar.Value & "," & _
                                txtValorNominal.Value & ",'" & strCodTipoTasa & "','" & strCodBaseAnual & "'," & CDec(dblTasaInteres) & _
                                ",'','" & gstrLogin & "'," & IIf(indCargadoDesdeBandeja, 1, 0) & " ) }"
                
                adoConn.Execute .CommandText
                                                                                                      
            End With
                    
            MsgBox Mensaje_Adicion_Exitosa, vbExclamation
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
            If indCargadoDesdeBandeja Then
                Unload Me
                Exit Sub
            End If
            
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
    'JCB -- Incluye Limites
    Me.MousePointer = vbDefault
    strMsgError = "Error " & Str$(err.Number) & vbNewLine
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

End Sub

Private Function TodoOK() As Boolean
 
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
                          
    If cboEmisor.ListIndex <= 0 Then
        MsgBox "Debe seleccionar el Emisor.", vbCritical, Me.Caption

        If cboEmisor.Enabled Then cboEmisor.SetFocus
        Exit Function
    End If
        
    If Trim$(txtDescripOrden.Text) = Valor_Caracter Then
        MsgBox "Debe indicar la Descripción de la SOLICITUD.", vbCritical, Me.Caption

        If txtDescripOrden.Enabled Then txtDescripOrden.SetFocus
        Exit Function
    End If
    
    If CDbl(txtDiasPlazo.Text) = 0 Then
        MsgBox "Debe indicar el número de días de plazo.", vbCritical, Me.Caption

        If txtDiasPlazo.Enabled Then txtDiasPlazo.SetFocus
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

Public Sub Salir()

    Unload Me
    
End Sub

Private Sub cboBaseAnual_Click()

    strCodBaseAnual = Valor_Caracter

    If cboBaseAnual.ListIndex < 0 Then Exit Sub
    
    strCodBaseAnual = Trim$(arrBaseAnual(cboBaseAnual.ListIndex))
    
End Sub

Private Sub cboClaseInstrumento_Click()

    strCodClaseInstrumento = Valor_Caracter

    If cboClaseInstrumento.ListIndex < 0 Then Exit Sub
    
    strCodClaseInstrumento = Trim$(arrClaseInstrumento(cboClaseInstrumento.ListIndex))

End Sub

Private Sub cboComisionista_Click()
    
    strCodComisionista = Valor_Caracter
    numSecCondicion = 0
    
    If cboComisionista.ListIndex < 0 Then Exit Sub
    
    strCodComisionista = Mid$(arrComisionista(cboComisionista.ListIndex), 1, 8)
    numSecCondicion = Mid$(arrComisionista(cboComisionista.ListIndex), 9)


End Sub

Private Sub cboEmisor_Click()

    Dim adoRegistro As ADODB.Recordset
    
    strCodTitulo = Valor_Caracter
    strCodEmisor = Valor_Caracter: strCodAnalitica = Valor_Caracter
    
    If cboEmisor.ListIndex < 0 Then Exit Sub
    
    strCodEmisor = arrEmisor(cboEmisor.ListIndex) '), 8)
    
    'Asignando el nemónico
    Set adoRegistro = New ADODB.Recordset
    adoComm.CommandText = "select DescripNemonico from InstitucionPersona where TipoPersona = '02' and CodPersona = '" & strCodEmisor & "'"
    Set adoRegistro = adoComm.Execute
    
    txtDescripOrden.Text = Trim$(cboTipoInstrumentoOrden.Text)
    adoRegistro.Close
    
    'Obtener lista de comisionistas
    If strCodEmisor <> Valor_Caracter Then
        strSQL = "{ call up_ACLstFondoComisionistaContraparte('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                    Codigo_Tipo_Comisionista_Inversion & "','" & Codigo_Tipo_Persona_Emisor & "','" & strCodEmisor & "','" & _
                    strCodMoneda & "','" & gstrFechaActual & "') }"
        CargarControlLista strSQL, cboComisionista, arrComisionista(), Valor_Caracter
    Else
        cboComisionista.Clear
    End If
    
    If cboComisionista.ListCount = 1 Then
        cboComisionista.ListIndex = 0
    End If

    
    '*** Validar Limites ***
    If strCodTipoInstrumentoOrden = Valor_Caracter Then Exit Sub

    If blnCancelaPrepago = False Then
        strCodTitulo = strCodFondoOrden & strCodFile & strCodAnalitica
    End If
    
    With adoComm
        Set adoRegistro = New ADODB.Recordset
                        
        '*** Categoría del instrumento emitido por el emisor ***
        .CommandText = "SELECT CodCategoriaRiesgo,CodRiesgoFinal,CodSubRiesgoFinal FROM EmisionInstitucionPersona " & "WHERE CodEmisor='" & strCodEmisor & "' AND CodFile='" & strCodTipoInstrumentoOrden & "' AND " & "CodDetalleFile='" & strCodClaseInstrumento & "'"
            
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            strCodRiesgo = Trim$(adoRegistro("CodRiesgoFinal"))
            strCodSubRiesgo = Trim$(adoRegistro("CodSubRiesgoFinal"))
        Else

            If strCodEmisor <> Valor_Caracter And indCargadoDesdeBandeja = False Then
                Exit Sub
            End If
        End If

        adoRegistro.Close
             
        Set adoRegistro = Nothing
    
    End With
    
End Sub

Private Sub cboEstado_Click()

    strCodEstado = Valor_Caracter

    If cboEstado.ListIndex < 0 Then Exit Sub
    
    strCodEstado = Trim$(arrEstado(cboEstado.ListIndex))
    
    Call Buscar
End Sub

Public Sub setFondo(strCodF As String)
    strCodFondoSol = strCodF
    blnFlag = True
End Sub

Private Sub cboFondo_Click()

    Dim adoRegistro As ADODB.Recordset
    
    strCodFondo = Valor_Caracter

    If cboFondo.ListIndex < 0 Then Exit Sub
    strCodFondo = Trim$(arrFondo(cboFondo.ListIndex))
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Fecha Vigente, Moneda ***
        .CommandText = "{ call up_ACSelDatosParametro(23,'" & strCodFondo & "','" & gstrCodAdministradora & "','000') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            gdatFechaActual = CVDate(adoRegistro("FechaCuota"))
            dtpFechaOrdenDesde.Value = gdatFechaActual
            dtpFechaOrdenHasta.Value = dtpFechaOrdenDesde.Value
            strCodMoneda = Trim$(adoRegistro("CodMoneda"))
                        
            frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
        End If

        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    '*** Tipo de Instrumento ***
    strSQL = "SELECT FIF.CodFile CODIGO,DescripFile DESCRIP " & "FROM FondoInversionFile FIF JOIN InversionFile IVF ON(IVF.CodFile=FIF.CodFile) " & "WHERE TipoValor='" & Codigo_Valor_RentaFija & "' AND TipoPlazo='" & Codigo_Valor_CortoPlazo & "' AND IndInstrumento='X' AND IndVigente='X' AND FIF.CodFile = '" & CodFile_Descuento_Flujos_Dinerarios & "' AND " & "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' ORDER BY DescripFile"
    CargarControlLista strSQL, cboTipoInstrumento, arrTipoInstrumento(), Sel_Todos
    
    If cboTipoInstrumento.ListCount > 0 Then cboTipoInstrumento.ListIndex = 0
        
End Sub

Private Sub cboFondoOrden_Click()

    Dim adoRegistro As ADODB.Recordset
    
    strCodFondoOrden = Valor_Caracter

    If cboFondoOrden.ListIndex < 0 Then Exit Sub
    
    strCodFondoOrden = Trim$(arrFondoOrden(cboFondoOrden.ListIndex))

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
            dtpFechaVencimiento.Value = dtpFechaOrden.Value
            strCodMoneda = Trim$(adoRegistro("CodMoneda"))
                        
            frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
            Me.Refresh
        End If

        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    '*** Tipo de Instrumento ***
    strSQL = "SELECT FIF.CodFile CODIGO,DescripFile DESCRIP " & "FROM FondoInversionFile FIF JOIN InversionFile IVF ON(IVF.CodFile=FIF.CodFile) " & "WHERE TipoValor='" & Codigo_Valor_RentaFija & "' AND TipoPlazo='" & Codigo_Valor_CortoPlazo & "' AND IndInstrumento='X' AND IndVigente='X' and FIF.CodFile = '" & CodFile_Descuento_Flujos_Dinerarios & "' AND " & "CodFondo='" & strCodFondoOrden & "' AND CodAdministradora='" & gstrCodAdministradora & "' ORDER BY DescripFile"
    CargarControlLista strSQL, cboTipoInstrumentoOrden, arrTipoInstrumentoOrden(), Sel_Defecto
    If cboTipoInstrumentoOrden.ListCount > 0 Then cboTipoInstrumentoOrden.ListIndex = 0
End Sub

Private Sub cboCobroInteres_Click()

    strCodCobroInteres = Valor_Caracter

    If cboCobroInteres.ListIndex < 0 Then Exit Sub

    strCodCobroInteres = Mid$(Trim$(arrPagoInteres(cboCobroInteres.ListIndex)), 7, 2)

    If strCodCobroInteres = Codigo_Modalidad_Pago_Adelantado Then   'Si es pago de intereses adelantados permitir la edición de int. adicionales
        If (strCodTipoInstrumentoOrden = "015" And strCodClaseInstrumento = "001" And strCodSubClaseInstrumento = "001") Then   'Sòlo en caso de letras
        
            'Los días adicionales se suman a la fecha de vencimieno del documento
            datFechaVctoAdicional = DateAdd("d", intDiasAdicionales, CVDate(dtpFechaVencimiento.Value))

            If Not EsDiaUtil(datFechaVctoAdicional) Then
                datFechaVctoAdicional = ProximoDiaUtil(datFechaVctoAdicional)
            End If
           
            txtDiasPlazo.Text = CStr(DateDiff("d", dtpFechaOrden.Value, dtpFechaVencimiento.Value))
            
        End If

    Else

        If (strCodTipoInstrumentoOrden = "015" And strCodClaseInstrumento = "001" And strCodSubClaseInstrumento = "001") Then   'Sòlo en caso de letras
            
            datFechaVctoAdicional = dtpFechaVencimiento.Value
           
            txtDiasPlazo.Text = CStr(DateDiff("d", dtpFechaOrden.Value, dtpFechaVencimiento.Value))
    
        End If
    End If
           
End Sub

Private Sub cboSubClaseInstrumento_Click()
    strCodSubClaseInstrumento = Valor_Caracter

    If cboSubClaseInstrumento.ListIndex < 0 Then Exit Sub
    
    strCodSubClaseInstrumento = Trim$(arrSubClaseInstrumento(cboSubClaseInstrumento.ListIndex))

End Sub

Private Sub cboTipoInstrumento_Click()

    strCodTipoInstrumento = Valor_Caracter

    If cboTipoInstrumento.ListIndex < 0 Then Exit Sub
    
    strCodTipoInstrumento = Trim$(arrTipoInstrumento(cboTipoInstrumento.ListIndex))
    
    Call Buscar
    
End Sub

Private Sub cboTipoInstrumentoOrden_Click()
    
    strCodTipoInstrumentoOrden = Valor_Caracter
    strIndPacto = Valor_Caracter: strIndNegociable = Valor_Caracter

    If cboTipoInstrumentoOrden.ListIndex < 0 Then Exit Sub
    
    strCodTipoInstrumentoOrden = Trim$(arrTipoInstrumentoOrden(cboTipoInstrumentoOrden.ListIndex))

'    If strCodTipoInstrumentoOrden = "010" Then   'Letras
'        cboCobroInteres.ListIndex = ObtenerItemLista(arrPagoInteres(), "MODPAG" + Codigo_Modalidad_Pago_Vencimiento)
'        cboCobroInteres.Enabled = False
'    Else
'        cboCobroInteres.ListIndex = ObtenerItemLista(arrPagoInteres(), "MODPAG" + Codigo_Modalidad_Pago_Adelantado)
'        cboCobroInteres.Enabled = True
'    End If

    strCodFile = strCodTipoInstrumentoOrden

    '*** Clase de Instrumento ***
    strSQL = "SELECT CodDetalleFile CODIGO,DescripDetalleFile DESCRIP FROM InversionDetalleFile WHERE CodFile='" & strCodTipoInstrumentoOrden & "' AND IndVigente='X' ORDER BY DescripDetalleFile"
    CargarControlLista strSQL, cboClaseInstrumento, arrClaseInstrumento(), Sel_Defecto
    
    If cboClaseInstrumento.ListCount > 0 Then
        cboClaseInstrumento.ListIndex = 0
        cboClaseInstrumento.Enabled = True
    End If
            
End Sub

Private Sub cboMoneda_Click()
    
    If cboMoneda.ListIndex < 0 Then Exit Sub
    
    strCodMoneda = Trim$(arrMoneda(cboMoneda.ListIndex))
    
    'Obtener Comisionistas de la moneda elegida con el obligado seleccionado
    
    strSQL = "{ call up_ACLstFondoComisionistaContraparte('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
                Codigo_Tipo_Comisionista_Inversion & "','" & Codigo_Tipo_Persona_Emisor & "','" & strCodEmisor & "','" & _
                strCodMoneda & "','" & gstrFechaActual & "') }"
    CargarControlLista strSQL, cboComisionista, arrComisionista(), Valor_Caracter
    
    If cboComisionista.ListCount = 1 Then
        cboComisionista.ListIndex = 0
    End If
    
End Sub

Private Sub cboTipoTasa_Click()

    strCodTipoTasa = Valor_Caracter

    If cboTipoTasa.ListIndex < 0 Then Exit Sub
    
    strCodTipoTasa = Trim$(arrTipoTasa(cboTipoTasa.ListIndex))
    
End Sub

Private Sub cmdEnviar_Click()

    Dim strFechaDesde As String, strFechaHasta        As String
    Dim intRegistro   As Integer, intContador         As Integer
    Dim datFecha      As Date, strCodEstadoRegistro As String
    Dim indProceso    As Boolean
    
    If adoConsulta.Recordset.RecordCount = 0 Then Exit Sub
    
    strFechaDesde = Convertyyyymmdd(dtpFechaOrdenDesde.Value)
    datFecha = DateAdd("d", 1, dtpFechaOrdenHasta.Value)
    strFechaHasta = Convertyyyymmdd(datFecha)
    
    intContador = tdgConsulta.SelBookmarks.Count - 1
    
    If intContador < 0 Then
        MsgBox "No se ha seleccionado ningún registro", vbCritical, Me.Caption
        Exit Sub
    End If
    
    indProceso = False
        
    For intRegistro = 0 To intContador
        tdgConsulta.Row = tdgConsulta.SelBookmarks(intRegistro) - 1
        strCodEstadoRegistro = Trim$(tdgConsulta.Columns("EstadoSolicitud"))
        
        If strCodEstadoRegistro = "01" Or strCodEstadoRegistro = "02" Then
            indProceso = True
            
            If strCodEstadoRegistro = "01" Then
                strCodEstadoRegistro = "02"
            ElseIf strCodEstadoRegistro = "02" Then
                strCodEstadoRegistro = "01"
            End If
        
            adoComm.CommandText = "UPDATE InversionSolicitud SET EstadoSolicitud='" & strCodEstadoRegistro & "'," & "UsuarioEdicion='" & gstrLogin & "',FechaEdicion='" & strFechaDesde & Space$(1) & Format$(Time, "hh:mm") & "' " & "WHERE NumSolicitud='" & Trim$(tdgConsulta.Columns(0)) & "' AND CodFondo='" & strCodFondo & "' AND " & "CodAdministradora='" & gstrCodAdministradora & "'"
                
            adoConn.Execute adoComm.CommandText
        End If

    Next
    
    If indProceso Then
        MsgBox "El Proceso terminó en forma exitosa...", vbExclamation, gstrNombreEmpresa
    End If
    
    Call Buscar
    
End Sub

Private Sub dtpFechaLiquidacion_Change()

    If dtpFechaLiquidacion.Value < dtpFechaOrden.Value Then
        dtpFechaLiquidacion.Value = dtpFechaOrden.Value
    End If
        
    If Not EsDiaUtil(dtpFechaLiquidacion.Value) Then
        MsgBox "La Fecha no es un día útil...se cambiará por una fecha correcta !", vbInformation, Me.Caption
        dtpFechaLiquidacion.Value = ProximoDiaUtil(dtpFechaLiquidacion.Value)
    End If
    
    If strCodTipoInstrumentoOrden = "015" Then
        'Si es un instrumento de descuento entonces la fecha de vencimiento es la fecha de vencimiento del dcto. a descontar y NO de la operación
        txtDiasPlazo.Text = CStr(DateDiff("d", dtpFechaOrden.Value, dtpFechaVencimiento.Value))
        
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

Private Sub dtpFechaVencimiento_Change()

    If dtpFechaVencimiento.Value < dtpFechaOrden.Value Then
        dtpFechaVencimiento.Value = dtpFechaOrden.Value
    End If
    
    If dtpFechaVencimiento.Value < dtpFechaLiquidacion.Value Then
        dtpFechaVencimiento.Value = dtpFechaLiquidacion.Value
    End If

    txtDiasPlazo.Text = CStr(DateDiff("d", dtpFechaOrden.Value, dtpFechaVencimiento.Value))
End Sub

Private Sub Form_Activate()

    frmMainMdi.stbMdi.Panels(3).Text = Me.Caption
    '    Call CargarReportes
    
End Sub

Private Sub Form_Deactivate()

    Call OcultarReportes
    
End Sub

Private Sub Form_Load()
   
    indCargadoDesdeBandeja = False
    
    Call InicializarValores
    Call CargarListas
    ' Call CargarReportes
    Call Buscar
    
    Call ValidarPermisoUsoControl(Trim$(gstrLogin), Me, Trim$(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)

    If blnFlag Then
        cboFondo.ListIndex = 0
    End If

    CentrarForm Me
              
    Call ValidaExisteTipoCambio(Codigo_TipoCambio_SBS, gstrFechaActual)
     
End Sub

Public Sub Buscar()

    Dim strFechaSolicitudDesde As String, strFechaSolicitudHasta        As String
    
    Dim datFechaSiguiente      As Date

    Me.MousePointer = vbHourglass
    
    If Not IsNull(dtpFechaOrdenDesde.Value) And Not IsNull(dtpFechaOrdenHasta.Value) Then
        strFechaSolicitudDesde = Convertyyyymmdd(dtpFechaOrdenDesde.Value)
        datFechaSiguiente = DateAdd("d", 1, dtpFechaOrdenHasta.Value)
        strFechaSolicitudHasta = Convertyyyymmdd(datFechaSiguiente)
    End If
        
    strSQL = "SELECT IOR.NumSolicitud,FechaSolicitud,FechaLiquidacion,CodTitulo,EstadoSolicitud,IOR.CodFile,CodAnalitica,TipoSolicitud,IOR.CodMoneda," & "DescripSolicitud,MontoSolicitud,MontoAprobado, " & "CodSigno DescripMoneda, IOR.CodDetalleFile, IOR.CodSubDetalleFile, IOR.CodFondo, " & "IOR.CodEmisor, IP1.DescripPersona DesEmisor,EstadoSolicitud, EST.DescripParametro AS DescripEstado " & "FROM InversionSolicitud IOR JOIN AuxiliarParametro EST ON(EST.CodParametro=IOR.EstadoSolicitud AND EST.CodTipoParametro = 'ESTSCF') " & "JOIN Moneda MON ON(MON.CodMoneda=IOR.CodMoneda) " & "LEFT JOIN InstitucionPersona IP1 ON (IP1.CodPersona = IOR.CodEmisor AND IP1.TipoPersona = '" & Codigo_Tipo_Persona_Emisor & "') " & "WHERE IOR.CodAdministradora='" & gstrCodAdministradora & "' AND IOR.CodFondo='" & strCodFondo & "' "
        
    If strCodTipoInstrumento <> Valor_Caracter Then
        strSQL = strSQL & "AND IOR.CodFile='" & strCodTipoInstrumento & "' "
    Else
        strSQL = strSQL & "AND IOR.CodFile IN " & strCodigosFile & " "
    End If

    If Not IsNull(dtpFechaOrdenDesde.Value) And Not IsNull(dtpFechaOrdenHasta.Value) Then
        strSQL = strSQL & "AND (FechaSolicitud >='" & strFechaSolicitudDesde & "' AND FechaSolicitud <'" & strFechaSolicitudHasta & "') "
    End If

    If strCodEstado <> Valor_Caracter Then
        strSQL = strSQL & "AND EstadoSolicitud='" & strCodEstado & "' "
    End If
        
    strSQL = strSQL & "ORDER BY IOR.NumSolicitud"
    
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

Private Sub CargarListas()

    Dim intRegistro As Integer
    
    '*** Fondos ***
    strSQL = "SELECT CodFondo CODIGO,DescripFondo DESCRIP FROM Fondo WHERE CodAdministradora='" & gstrCodAdministradora & "' AND Estado='01' and CodFondo = '" & gstrCodFondoContable & "' ORDER BY DescripFondo"
    CargarControlLista strSQL, cboFondo, arrFondo(), Valor_Caracter
    CargarControlLista strSQL, cboFondoOrden, arrFondoOrden(), Valor_Caracter
        
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
        
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
            
    '*** Estados de la Solicitud ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE Estado = '01' AND CodTipoParametro='ESTSCF' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboEstado, arrEstado(), Sel_Todos
    
    intRegistro = ObtenerItemLista(arrEstado(), "01")

    If intRegistro >= 0 Then cboEstado.ListIndex = intRegistro
    
    '*** Emisor ***
    strSQL = "SELECT CodPersona CODIGO,DescripPersona DESCRIP FROM InstitucionPersona WHERE TipoPersona='" & Codigo_Tipo_Persona_Emisor & "' AND IndVigente='X' ORDER BY DescripPersona"
    CargarControlLista strSQL, cboEmisor, arrEmisor(), Sel_Defecto
    If cboEmisor.ListCount > 0 Then cboEmisor.ListIndex = 0
    
    '*** Moneda ***
    strSQL = "{ call up_ACSelDatos(2) }"
    CargarControlLista strSQL, cboMoneda, arrMoneda(), Sel_Defecto
    
    '*** Base de Cálculo ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='BASANU' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboBaseAnual, arrBaseAnual(), Valor_Caracter
    
    '*** Tipo Tasa ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='NATTAS' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboTipoTasa, arrTipoTasa(), ""
    
    '*** Momento de cobro de los intereses (por defecto es al inicio) ***
    strSQL = "SELECT (CodTipoParametro + CodParametro) CODIGO, DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='MODPAG' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboCobroInteres, arrPagoInteres(), ""

    If cboCobroInteres.ListCount > 0 Then
           
        If cboCobroInteres.ListCount = 2 Then
            cboCobroInteres.ListIndex = 1
        Else
            cboCobroInteres.ListIndex = 0
        End If
    End If
  
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

        Case vModify
            Call Modificar

        Case vCancel
            blnCancelaPrepago = False
            Call Cancelar

        Case vExit
            Call Salir
        
    End Select
    
End Sub

Public Sub Modificar()

    If strEstado = Reg_Consulta Then
        If Trim$(tdgConsulta.Columns("EstadoSolicitud")) = "01" Then
            strEstado = Reg_Edicion
            LlenarFormulario strEstado
            cmdOpcion.Visible = False

            With tabRFCortoPlazo
                .TabEnabled(0) = False
                .TabEnabled(1) = True
                .TabEnabled(2) = False
                .Tab = 1
            End With

        Else
            MsgBox "Solicitud no se puede modificar...", vbExclamation, gstrNombreEmpresa
        End If
    End If
    
End Sub

Private Sub InicializarValores()
    
    Dim adoRegistro As ADODB.Recordset
    
    '*** Valores Iniciales ***
    strEstado = Reg_Defecto
    tabRFCortoPlazo.Tab = 0
    
    dtpFechaOrdenDesde.Value = gdatFechaActual
    dtpFechaOrdenHasta.Value = dtpFechaOrdenDesde.Value
    
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
            
            strCodigosFile = strCodigosFile & Trim$(adoRegistro("CodFile")) & "'"
        
            adoRegistro.MoveNext
        Loop

        adoRegistro.Close: Set adoRegistro = Nothing
                
        strCodigosFile = "('" & strCodigosFile & ",'009')"
    End With
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    
    'Leer si las comisiones van a ser definidas en la operación o ya viene establecida
    Set adoRegistro = New ADODB.Recordset
    adoComm.CommandText = "SELECT ValorParametro AS PersonalizaComi FROM ParametroGeneral WHERE CodParametro = '32'"
    Set adoRegistro = adoComm.Execute

    If Not (adoRegistro.EOF) Then
        strPersonalizaComision = Trim$(adoRegistro("PersonalizaComi"))
    End If
    
    'Leer el porcentaje de descuento
    Set adoRegistro = New ADODB.Recordset
    adoComm.CommandText = "SELECT ValorParametro AS PorcentajeDscto FROM ParametroGeneral WHERE CodParametro = '33'"
    Set adoRegistro = adoComm.Execute

    adoRegistro.Close: Set adoRegistro = Nothing
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmOrdenRentaFijaCortoPlazo = Nothing
    Call OcultarReportes
    frmMainMdi.stbMdi.Panels(3).Text = "Acción"
    
End Sub

Private Sub tabRFCortoPlazo_Click(PreviousTab As Integer)
    
    Select Case tabRFCortoPlazo.Tab

        Case 1, 2, 3, 4

            If PreviousTab = 0 And indCargadoDesdeBandeja = False And strEstado = Reg_Consulta Then tabRFCortoPlazo.Tab = 0
            If strEstado = Reg_Defecto Then tabRFCortoPlazo.Tab = 0
                                    
            If tabRFCortoPlazo.Tab = 2 Then
                If strCodTipoSolicitud <> Codigo_Orden_PagoCancelacion And strCodTipoSolicitud <> Codigo_Orden_Prepago Then
                    If ValidaRequisitosTab(2, PreviousTab) = True Then
                        fraDatosNegociacion.Caption = "Negociación"
                    Else
                        tabRFCortoPlazo.Tab = 1
                    End If

                Else
                    MsgBox "La ficha de Negociación no está permitida para la Cancelación o Prepago.", vbCritical
                    tabRFCortoPlazo.Tab = PreviousTab
                End If
            End If
            
    End Select
            
End Sub

Private Function ValidaRequisitosTab(intIndTab As Integer, intTabOrigen) As Boolean

    ValidaRequisitosTab = False

    Select Case intIndTab

        Case 2
 
            If (CDbl(txtDiasPlazo.Text) <= 0 Or cboMoneda.ListIndex <= 0) And indCargadoDesdeBandeja = False Then
                MsgBox "Verifique si la moneda y el plazo están ingresados.", vbCritical, Me.Caption
                Exit Function
            End If
     
            If cboEmisor.ListIndex <= 0 And indCargadoDesdeBandeja = False Then
                MsgBox "Debe seleccionar el Emisor.", vbCritical, Me.Caption

                If cboEmisor.Enabled And tabRFCortoPlazo.TabEnabled(1) = True Then
                    tabRFCortoPlazo.Tab = 1
                    cboEmisor.SetFocus
                End If

                Exit Function
            End If
    
    End Select

    ValidaRequisitosTab = True

End Function

Private Sub tdgConsulta_FormatText(ByVal ColIndex As Integer, _
                                   Value As Variant, _
                                   Bookmark As Variant)

    If ColIndex = 4 Then
        Call DarFormatoValor(Value, Decimales_Monto)
    End If
    
    If ColIndex = 7 Then
        Call DarFormatoValor(Value, Decimales_Monto)
    End If
    
End Sub

Private Sub txtDiasPlazo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
    
        If IsNumeric(txtDiasPlazo.Text) Then
            dtpFechaVencimiento.Value = DateAdd("d", txtDiasPlazo.Text, CVDate(dtpFechaOrden.Value))
        Else
            dtpFechaVencimiento.Value = dtpFechaOrden.Value
        End If
            
        dtpFechaVencimiento_Change
    
    End If
    
End Sub

Private Sub txtDiasPlazo_LostFocus()

    txtDiasPlazo_KeyPress (vbKeyReturn)
    cboEmisor_Click
    
End Sub

Private Sub txtTasa_Change()

    Call FormatoCajaTexto(txtTasa, Decimales_Tasa)

End Sub

Private Sub txtTasa_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtTasa, Decimales_Tasa)
    
End Sub

Public Sub HabilitaCombos(ByVal pBloquea As Boolean)

    cboFondoOrden.Enabled = pBloquea
    cboTipoInstrumentoOrden.Enabled = pBloquea
    cboClaseInstrumento.Enabled = pBloquea
    cboEmisor.Enabled = pBloquea

End Sub

Public Sub mostrarForm(ByVal strNumSolicitud As String)

    Load Me
    
    indCargadoDesdeBandeja = True
        
    LlenarFormulario Reg_Edicion, strNumSolicitud
    
    fraDatosBasicos.Enabled = False
    fraDatosTitulo.Enabled = False
    
    txtValorNominal.Text = txtValorFinanciar.Value
    
    '*** SubClase de Instrumento ***
    strSQL = "SELECT CodSubDetalleFile CODIGO,DescripSubDetalleFile DESCRIP FROM InversionSubDetalleFile WHERE " & "CodDetalleFile='" & strCodClaseInstrumento & "' AND CodFile='" & strCodTipoInstrumentoOrden & "' AND IndVigente='X' ORDER BY DescripSubDetalleFile"
        
    CargarControlLista strSQL, cboSubClaseInstrumento, arrSubClaseInstrumento(), ""
    
    If cboSubClaseInstrumento.ListCount > 0 Then
        If cboSubClaseInstrumento.ListCount = 2 Then
            cboSubClaseInstrumento.ListIndex = 1
        Else
            cboSubClaseInstrumento.ListIndex = 0
        End If
    End If
    
    strEstado = Reg_Edicion
    
    With tabRFCortoPlazo
        .TabEnabled(0) = False
        .TabEnabled(1) = True
        .TabEnabled(2) = True
        .Tab = 2
    End With
    
    cmdOpcion.Visible = False
        
    Me.Show
    Me.SetFocus
    
    
End Sub

Private Sub txtValorNominal_Change()

    If Val(txtValorNominal.Value) > Val(txtValorFinanciar.Value) Then
        MsgBox "El monto no puede exceder el monto a financiar", vbInformation, Me.Caption
        txtValorNominal.Text = txtValorFinanciar.Value
    End If

End Sub
