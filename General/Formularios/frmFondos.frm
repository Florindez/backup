VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmFondos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fondos Administrados"
   ClientHeight    =   7980
   ClientLeft      =   1020
   ClientTop       =   705
   ClientWidth     =   11745
   FillColor       =   &H00800000&
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7980
   ScaleWidth      =   11745
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   9840
      TabIndex        =   18
      Top             =   7140
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
      Caption0        =   "&Salir"
      Tag0            =   "9"
      ToolTipText0    =   "Salir"
      UserControlWidth=   1200
   End
   Begin TAMControls2.ucBotonEdicion2 cmdAccion 
      Height          =   735
      Left            =   6480
      TabIndex        =   70
      Top             =   7140
      Visible         =   0   'False
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
   Begin TAMControls2.ucBotonEdicion2 cmdOpcion 
      Height          =   735
      Left            =   360
      TabIndex        =   17
      Top             =   7140
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   1296
      Buttons         =   4
      Caption0        =   "&Nuevo"
      Tag0            =   "0"
      ToolTipText0    =   "Nuevo"
      Caption1        =   "&Modificar"
      Tag1            =   "3"
      ToolTipText1    =   "Modificar"
      Caption2        =   "&Eliminar"
      Tag2            =   "4"
      ToolTipText2    =   "Eliminar"
      Caption3        =   "&Buscar"
      Tag3            =   "5"
      ToolTipText3    =   "Buscar"
      UserControlWidth=   5700
   End
   Begin TabDlg.SSTab tabFondo 
      Height          =   6915
      Left            =   0
      TabIndex        =   19
      Top             =   90
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   12197
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
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
      TabPicture(0)   =   "frmFondos.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tdgConsulta"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Datos Generales"
      TabPicture(1)   =   "frmFondos.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraFechas"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "fraFondo(0)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Parámetros"
      TabPicture(2)   =   "frmFondos.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraFondo(1)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "fraFondo(2)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Regulación"
      TabPicture(3)   =   "frmFondos.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraDatosRegulacion"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.Frame fraFondo 
         Caption         =   "Definición"
         Height          =   3555
         Index           =   0
         Left            =   -74790
         TabIndex        =   20
         Top             =   420
         Width           =   11205
         Begin VB.CommandButton cmdBusqueda 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   8340
            TabIndex        =   62
            Top             =   2940
            Width           =   315
         End
         Begin VB.TextBox txtDescripCustodio 
            Height          =   315
            Left            =   2730
            Locked          =   -1  'True
            TabIndex        =   61
            Top             =   2940
            Width           =   5625
         End
         Begin VB.CheckBox chkIndicadorFondoRegulado 
            Caption         =   "Fondo Regulado"
            Height          =   255
            Left            =   6150
            TabIndex        =   60
            Top             =   840
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.TextBox txtRucFondo 
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2730
            MaxLength       =   11
            TabIndex        =   4
            Top             =   1650
            Width           =   2760
         End
         Begin VB.ComboBox cboTipoCartera 
            Height          =   315
            Left            =   2730
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   2505
            Width           =   2820
         End
         Begin VB.TextBox txtDescripFondo 
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2730
            MaxLength       =   200
            TabIndex        =   2
            Top             =   1230
            Width           =   5910
         End
         Begin VB.ComboBox cboTipoFondo 
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2730
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   813
            Width           =   3060
         End
         Begin VB.ComboBox cboMoneda 
            Height          =   315
            Left            =   2730
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   2070
            Width           =   2775
         End
         Begin VB.Label lblFondo 
            AutoSize        =   -1  'True
            Caption         =   "Num. RUC"
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
            Left            =   360
            TabIndex        =   43
            Top             =   1710
            Width           =   900
         End
         Begin VB.Label lblFondo 
            AutoSize        =   -1  'True
            Caption         =   "Custodio"
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
            TabIndex        =   39
            Top             =   2985
            Width           =   750
         End
         Begin VB.Label lblFondo 
            Caption         =   "Cartera"
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
            Left            =   360
            TabIndex        =   28
            Top             =   2550
            Width           =   1125
         End
         Begin VB.Label lblFondo 
            Caption         =   "Denominación"
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
            Index           =   2
            Left            =   360
            TabIndex        =   25
            Top             =   1290
            Width           =   1245
         End
         Begin VB.Label lblFondo 
            Caption         =   "Tipo de Fondo"
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
            Index           =   1
            Left            =   360
            TabIndex        =   24
            Top             =   870
            Width           =   1245
         End
         Begin VB.Label lblFondo 
            Caption         =   "Código"
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
            Index           =   0
            Left            =   360
            TabIndex        =   23
            Top             =   435
            Width           =   645
         End
         Begin VB.Label lblFondo 
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
            Height          =   285
            Index           =   25
            Left            =   360
            TabIndex        =   22
            Top             =   2130
            Width           =   765
         End
         Begin VB.Label lblCodFondo 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2730
            TabIndex        =   21
            Top             =   390
            Width           =   1155
         End
      End
      Begin VB.Frame fraFechas 
         Caption         =   "Fechas y Plazos"
         Height          =   2655
         Left            =   -74790
         TabIndex        =   63
         Top             =   4050
         Width           =   5595
         Begin VB.TextBox txtPlazo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2730
            MaxLength       =   3
            TabIndex        =   64
            Top             =   1170
            Width           =   1605
         End
         Begin MSComCtl2.DTPicker dtpFechaOperativa 
            Height          =   315
            Left            =   2730
            TabIndex        =   65
            Top             =   780
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CalendarForeColor=   8388608
            DateIsNull      =   -1  'True
            Format          =   175374337
            CurrentDate     =   38068
         End
         Begin MSComCtl2.DTPicker dtpFechaPreOperativa 
            Height          =   315
            Left            =   2730
            TabIndex        =   66
            Top             =   390
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            CalendarForeColor=   8388608
            Format          =   175374337
            CurrentDate     =   38068
         End
         Begin VB.Label lblFondo 
            AutoSize        =   -1  'True
            Caption         =   "Plazo (Años)"
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
            TabIndex        =   69
            Top             =   1200
            Width           =   1080
         End
         Begin VB.Label lblFondo 
            AutoSize        =   -1  'True
            Caption         =   "Etapa PreOperativa"
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
            Index           =   24
            Left            =   360
            TabIndex        =   68
            Top             =   450
            Width           =   1680
         End
         Begin VB.Label lblFondo 
            AutoSize        =   -1  'True
            Caption         =   "Inicio Actividades"
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
            TabIndex        =   67
            Top             =   840
            Width           =   1530
         End
      End
      Begin VB.Frame fraDatosRegulacion 
         Caption         =   "Informacion de Regulación"
         Height          =   2235
         Left            =   -74760
         TabIndex        =   51
         Top             =   630
         Width           =   9795
         Begin VB.ComboBox cboRegulador 
            Height          =   315
            Left            =   2925
            Style           =   2  'Dropdown List
            TabIndex        =   58
            Top             =   420
            Width           =   5850
         End
         Begin VB.TextBox txtNumConasev 
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2940
            MaxLength       =   15
            TabIndex        =   53
            Top             =   1200
            Width           =   1620
         End
         Begin VB.TextBox txtCodConasev 
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2940
            MaxLength       =   5
            TabIndex        =   52
            Top             =   1605
            Width           =   3150
         End
         Begin MSComCtl2.DTPicker dtpFechaConasev 
            Height          =   315
            Left            =   2940
            TabIndex        =   54
            Top             =   810
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            CalendarForeColor=   8388608
            Format          =   175374337
            CurrentDate     =   38068
         End
         Begin VB.Label lblFondo 
            AutoSize        =   -1  'True
            Caption         =   "Regulador"
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
            Left            =   360
            TabIndex        =   59
            Top             =   450
            Width           =   885
         End
         Begin VB.Label lblFondo 
            AutoSize        =   -1  'True
            Caption         =   "Resolución Conasev Nº"
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
            TabIndex        =   57
            Top             =   1245
            Width           =   2025
         End
         Begin VB.Label lblFondo 
            Caption         =   "Código Conasev"
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
            Left            =   360
            TabIndex        =   56
            Top             =   1635
            Width           =   1725
         End
         Begin VB.Label lblFondo 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Inscripción Fondo"
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
            Left            =   360
            TabIndex        =   55
            Top             =   855
            Width           =   2100
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Valor Cuota"
         Height          =   2655
         Left            =   -69060
         TabIndex        =   44
         Top             =   4050
         Width           =   5475
         Begin VB.CheckBox chkIndicadorTipoValuacionModificable 
            Caption         =   "Asignación Modificable"
            Height          =   255
            Left            =   360
            TabIndex        =   71
            Top             =   2100
            Value           =   1  'Checked
            Width           =   2505
         End
         Begin VB.TextBox txtValorNominal 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2730
            TabIndex        =   47
            Top             =   1170
            Width           =   2130
         End
         Begin VB.ComboBox cboFrecuencia 
            Height          =   315
            Left            =   2730
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Top             =   780
            Width           =   2430
         End
         Begin VB.ComboBox cboTipoValuacion 
            Height          =   315
            Left            =   2730
            Style           =   2  'Dropdown List
            TabIndex        =   45
            Top             =   390
            Width           =   2430
         End
         Begin MSComCtl2.DTPicker dtpHoraCorte 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "HH:mm"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   4
            EndProperty
            Height          =   285
            Left            =   2730
            TabIndex        =   72
            Top             =   1560
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   175374339
            UpDown          =   -1  'True
            CurrentDate     =   38831
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Hora de Corte"
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
            Height          =   315
            Index           =   4
            Left            =   360
            TabIndex        =   73
            Top             =   1590
            Width           =   1200
         End
         Begin VB.Label lblFondo 
            Caption         =   "Frecuencia Asignación"
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
            Height          =   405
            Index           =   21
            Left            =   360
            TabIndex        =   50
            Top             =   825
            Width           =   2085
         End
         Begin VB.Label lblFondo 
            Caption         =   "Asignación"
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
            Index           =   23
            Left            =   360
            TabIndex        =   49
            Top             =   435
            Width           =   2085
         End
         Begin VB.Label lblFondo 
            Caption         =   "Valor Nominal"
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
            Index           =   6
            Left            =   360
            TabIndex        =   48
            Top             =   1200
            Width           =   1845
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmFondos.frx":0070
         Height          =   5895
         Left            =   300
         OleObjectBlob   =   "frmFondos.frx":008A
         TabIndex        =   29
         Top             =   600
         Width           =   11025
      End
      Begin VB.Frame fraFondo 
         Caption         =   "Participación"
         Height          =   2295
         Index           =   2
         Left            =   -74790
         TabIndex        =   27
         Top             =   2580
         Width           =   10815
         Begin VB.TextBox txtSaldoMinMantenerMonto 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3030
            TabIndex        =   14
            Top             =   1620
            Width           =   1905
         End
         Begin VB.TextBox txtSaldoMinMantenerCuotas 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3030
            TabIndex        =   13
            Top             =   1230
            Width           =   1905
         End
         Begin VB.TextBox txtRedMinMonto 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   8265
            TabIndex        =   16
            Top             =   825
            Width           =   1905
         End
         Begin VB.TextBox txtRedMinCuotas 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   8265
            TabIndex        =   15
            Top             =   420
            Width           =   1905
         End
         Begin VB.TextBox txtSusMinMonto 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3030
            TabIndex        =   12
            Top             =   825
            Width           =   1905
         End
         Begin VB.TextBox txtSusMinCuotas 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3030
            TabIndex        =   11
            Top             =   420
            Width           =   1905
         End
         Begin VB.Label lblFondo 
            Caption         =   "Saldo Min. Mantener Cuotas"
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
            Index           =   11
            Left            =   360
            TabIndex        =   38
            Top             =   1275
            Width           =   2505
         End
         Begin VB.Label lblFondo 
            Caption         =   "Saldo Min. Mantener Monto"
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
            Index           =   12
            Left            =   360
            TabIndex        =   37
            Top             =   1650
            Width           =   2745
         End
         Begin VB.Label lblFondo 
            Caption         =   "Redención Min. en Cuotas"
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
            Index           =   15
            Left            =   5490
            TabIndex        =   36
            Top             =   465
            Width           =   2265
         End
         Begin VB.Label lblFondo 
            Caption         =   "Redención Min. en Monto"
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
            Left            =   5490
            TabIndex        =   35
            Top             =   870
            Width           =   2715
         End
         Begin VB.Label lblFondo 
            Caption         =   "Susc. Min. en Cuotas"
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
            Index           =   13
            Left            =   360
            TabIndex        =   34
            Top             =   465
            Width           =   2265
         End
         Begin VB.Label lblFondo 
            Caption         =   "Susc. Min. en Monto"
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
            Index           =   14
            Left            =   360
            TabIndex        =   33
            Top             =   870
            Width           =   2265
         End
      End
      Begin VB.Frame fraFondo 
         Caption         =   "Fondo"
         Height          =   1965
         Index           =   1
         Left            =   -74790
         TabIndex        =   26
         Top             =   540
         Width           =   10815
         Begin VB.ComboBox cboSiNo 
            Height          =   315
            Left            =   8250
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1185
            Width           =   1905
         End
         Begin VB.TextBox txtMontoEmitido 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3030
            TabIndex        =   7
            Top             =   1215
            Width           =   1905
         End
         Begin VB.TextBox txtMontoEmision 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3030
            TabIndex        =   6
            Top             =   825
            Width           =   1905
         End
         Begin VB.TextBox txtPorcenParticipacion 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   8250
            TabIndex        =   9
            Top             =   795
            Width           =   1905
         End
         Begin VB.TextBox txtCapital 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   8235
            TabIndex        =   8
            Top             =   420
            Width           =   1905
         End
         Begin VB.TextBox txtPatrimonio 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3030
            TabIndex        =   5
            Top             =   420
            Width           =   1905
         End
         Begin VB.Label lblFondo 
            AutoSize        =   -1  'True
            Caption         =   "Monto Emitido"
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
            TabIndex        =   42
            Top             =   1260
            Width           =   1695
         End
         Begin VB.Label lblFondo 
            AutoSize        =   -1  'True
            Caption         =   "Monto Emisión"
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
            Index           =   17
            Left            =   360
            TabIndex        =   41
            Top             =   885
            Width           =   1725
         End
         Begin VB.Label lblFondo 
            AutoSize        =   -1  'True
            Caption         =   "Pagos Parciales"
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
            Index           =   26
            Left            =   5490
            TabIndex        =   40
            Top             =   1230
            Width           =   1380
         End
         Begin VB.Label lblFondo 
            Caption         =   "% Máximo Participación con relación al Patrimonio Neto"
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
            Height          =   420
            Index           =   8
            Left            =   5490
            TabIndex        =   32
            Top             =   735
            Width           =   2445
         End
         Begin VB.Label lblFondo 
            Caption         =   "Patrimonio Neto Mínimo Inicio Actividades"
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
            Height          =   435
            Index           =   4
            Left            =   360
            TabIndex        =   31
            Top             =   405
            Width           =   2415
         End
         Begin VB.Label lblFondo 
            Caption         =   "Patrimonio Neto"
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
            Index           =   5
            Left            =   5490
            TabIndex        =   30
            Top             =   450
            Width           =   2235
         End
      End
   End
End
Attribute VB_Name = "frmFondos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strTipoFondo                    As String, arrTipoFondo()                   As String
Dim strCodMoneda                    As String, arrCodMoneda()                   As String
Dim strTipoValuacion                As String, arrTipoValuacion()               As String
Dim strFrecuenciaValorizacion       As String, arrFrecuenciaValorizacion()      As String
Dim strTipoComision                 As String, arrTipoComision()                As String
Dim strTipoRescate                  As String, arrTipoRescate()                 As String
Dim strTipoCartera                  As String, arrTipoCartera()                 As String
Dim strCodCustodio                  As String, arrCustodio()                    As String
Dim strCodSiNo                      As String, arrSiNo()                        As String
Dim strCodRegulador                 As String, arrRegulador()                   As String
Dim strEstado                       As String, strCodFondo                      As String
Dim strIndActivo                    As String, strIndRegulado                   As String
Dim intNumPagos                     As Integer, strIndTipoValuacionModificable  As String
Dim strIndComision                  As String
Dim adoConsulta                     As ADODB.Recordset
Dim indSortAsc                      As Boolean, indSortDesc                     As Boolean

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

Public Sub Buscar()
        
    Dim strSQL As String
    
    Set adoConsulta = New ADODB.Recordset
    
    strSQL = "SELECT CodFondo,DescripFondo,CodSigno,ValorCuotaNominal," & _
        "ASIGVC.DescripParametro TipoValuacion " & _
        "FROM Fondo JOIN AuxiliarParametro ASIGVC ON (ASIGVC.CodParametro=TipoValuacion AND ASIGVC.CodTipoParametro='ASIGVC') " & _
        "JOIN Moneda TIPMON ON (TIPMON.CodMoneda=Fondo.CodMoneda) " & _
        "WHERE CodAdministradora='" & gstrCodAdministradora & "' AND Fondo.Estado='" & Estado_Activo & "' " & _
        "ORDER BY CodFondo"
                        
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




Public Sub Cancelar()

    cmdOpcion.Visible = True
    With tabFondo
        .TabEnabled(0) = True
        .Tab = 0
        .TabEnabled(1) = False
        .TabEnabled(2) = False
        .TabEnabled(3) = False
    End With
    strEstado = Reg_Consulta
    
End Sub

Private Sub CargarListas()

    Dim strSQL As String
                  
    '*** Tipo de Fondo ***
    strSQL = "{ call up_ACSelDatos(1) }"
    CargarControlLista strSQL, cboTipoFondo, arrTipoFondo(), Sel_Defecto
            
    '*** Moneda ***
    strSQL = "{ call up_ACSelDatos(2) }"
    CargarControlLista strSQL, cboMoneda, arrCodMoneda(), Sel_Defecto
    
    '*** Tipo de Cartera ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPCAR' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboTipoCartera, arrTipoCartera(), Sel_Defecto
    
  
    '*** Criterio Asignación Valor Cuota ***
    strSQL = "{ call up_ACSelDatos(3) }"
    CargarControlLista strSQL, cboTipoValuacion, arrTipoValuacion(), Sel_Defecto
    
    '*** Frecuencias de Valorización ***
    strSQL = "{ call up_ACSelDatos(25) }"
    CargarControlLista strSQL, cboFrecuencia, arrFrecuenciaValorizacion(), Sel_Defecto
    
    '*** Organismos reguladores ***
    strSQL = "{ call up_ACSelDatos(28) }"
    CargarControlLista strSQL, cboRegulador, arrRegulador(), Sel_Defecto
    
    '*** Afirmación Si/No ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='RESPSN' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboSiNo, arrSiNo(), Valor_Caracter
                        
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

Private Sub Deshabilita()
    
    fraFondo(1).Enabled = False
    fraFondo(2).Enabled = False
    
End Sub



Public Sub Imprimir()

    Call SubImprimir(1)
    
End Sub

Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Listado de Fondos"
    
End Sub

Private Sub InicializarValores()

    '*** Valores Iniciales ***
    strEstado = Reg_Defecto
    tabFondo.Tab = 0
    tabFondo.TabEnabled(1) = False
    tabFondo.TabEnabled(2) = False
    tabFondo.TabEnabled(3) = False
    
    
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgConsulta.Columns(0).Width = tdgConsulta.Width * 0.01 * 10
    tdgConsulta.Columns(1).Width = tdgConsulta.Width * 0.01 * 40
    tdgConsulta.Columns(2).Width = tdgConsulta.Width * 0.01 * 10
    tdgConsulta.Columns(3).Width = tdgConsulta.Width * 0.01 * 20
    tdgConsulta.Columns(4).Width = tdgConsulta.Width * 0.01 * 10
    
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    Set cmdOpcion.FormularioActivo = Me
    
End Sub

Private Sub LlenarFormulario(strModo As String)

    Dim adoRegistro   As ADODB.Recordset
    Dim intRegistro   As Integer
    Dim adoAuxiliar   As ADODB.Recordset
    Dim adoFondo      As ADODB.Recordset
    
    Select Case strModo
        Case Reg_Adicion
            Set adoRegistro = New ADODB.Recordset
            
            adoComm.CommandText = "SELECT COUNT(*) SecuencialFondo FROM Fondo"
            Set adoRegistro = adoComm.Execute
            
            If Not adoRegistro.EOF Then
                lblCodFondo.Caption = Format(adoRegistro("SecuencialFondo") + 1, "000")
            Else
                lblCodFondo.Caption = "001"
            End If
            adoRegistro.Close: Set adoRegistro = Nothing
            
            txtDescripFondo.Text = Valor_Caracter
            
            chkIndicadorFondoRegulado.Value = vbUnchecked
           
            chkIndicadorTipoValuacionModificable.Value = vbUnchecked
            
'            If gstrTipoAdministradora = Codigo_Tipo_Fondo_Mutuo Then
'                intRegistro = ObtenerItemLista(arrTipoFondo(), Codigo_Fondo_Abierto)
'                If intRegistro >= 0 Then cboTipoFondo.ListIndex = intRegistro
'                cboTipoFondo.Enabled = False
'            ElseIf gstrTipoAdministradora = Codigo_Tipo_Fondo_Inversion Then
'                intRegistro = ObtenerItemLista(arrTipoFondo(), Codigo_Fondo_Cerrado)
'                If intRegistro >= 0 Then cboTipoFondo.ListIndex = intRegistro
'                cboTipoFondo.Enabled = False
'            Else
'                intRegistro = ObtenerItemLista(arrTipoFondo(), Codigo_Fondo_Abierto)
'                If intRegistro >= 0 Then cboTipoFondo.ListIndex = intRegistro
'                cboTipoFondo.Enabled = True
'            End If

          ' ---------- Nueva Validacion Tipos de Fondos
            Set adoFondo = New ADODB.Recordset
            
            adoComm.CommandText = "SELECT COUNT(*) NumAdministradora FROM Fondo WHERE TipoFondo='03'"
            Set adoFondo = adoComm.Execute
            
            If adoFondo.EOF = False Then
               
               If CInt(adoFondo("NumAdministradora")) >= 1 Then
               
                    If gstrTipoAdministradora = Codigo_Tipo_Fondo_Mutuo Then
                        intRegistro = ObtenerItemLista(arrTipoFondo(), Codigo_Fondo_Abierto)
                        If intRegistro >= 0 Then cboTipoFondo.ListIndex = intRegistro
                        cboTipoFondo.Enabled = False
                    ElseIf gstrTipoAdministradora = Codigo_Tipo_Fondo_Inversion Then
                        intRegistro = ObtenerItemLista(arrTipoFondo(), Codigo_Fondo_Cerrado)
                        If intRegistro >= 0 Then cboTipoFondo.ListIndex = intRegistro
                        cboTipoFondo.Enabled = False
                    End If
                
               Else
               
                    intRegistro = ObtenerItemLista(arrTipoFondo(), Administradora_Fondos)
                    If intRegistro >= 0 Then cboTipoFondo.ListIndex = intRegistro
                    cboTipoFondo.Enabled = True
               
               End If
            
            End If
            adoFondo.Close: Set adoFondo = Nothing
          ' ----------------------------------------------------------------------------------
            
            
                        
            txtValorNominal.Text = "0"
            
            cboMoneda.ListIndex = -1
            If cboMoneda.ListCount > 0 Then cboMoneda.ListIndex = 0
            
            dtpFechaConasev.Value = gdatFechaActual
            If Not EsDiaUtil(dtpFechaConasev.Value) Then
                If dtpFechaConasev.Value >= gdatFechaActual Then
                    dtpFechaConasev.Value = AnteriorDiaUtil(dtpFechaConasev.Value)
                Else
                    dtpFechaConasev.Value = ProximoDiaUtil(dtpFechaConasev.Value)
                End If
            End If
            
            dtpFechaPreOperativa.Value = dtpFechaConasev.Value
            dtpFechaOperativa.Enabled = False
            dtpFechaOperativa.Value = dtpFechaPreOperativa.Value
            
            cboRegulador.ListIndex = -1
            
            txtNumConasev.Text = Valor_Caracter
            txtPlazo.Enabled = True
            txtPlazo.Text = "0"
                                    
            intRegistro = ObtenerItemLista(arrSiNo(), Codigo_Respuesta_No)
            If intRegistro >= 0 Then cboSiNo.ListIndex = intRegistro
            
            cboTipoCartera.ListIndex = -1
            If cboTipoCartera.ListCount > 0 Then cboTipoCartera.ListIndex = 0
            
            cboTipoValuacion.ListIndex = -1
            If cboTipoValuacion.ListCount > 0 Then cboTipoValuacion.ListIndex = 0
            
            cboRegulador.ListIndex = -1
            If cboRegulador.ListCount > 0 Then cboRegulador.ListIndex = 0
            
            cboFrecuencia.Enabled = True
            cboFrecuencia.ListIndex = -1
            If cboFrecuencia.ListCount > 0 Then cboFrecuencia.ListIndex = 0
                        
            If gstrTipoAdministradora = Codigo_Tipo_Fondo_Mutuo Then
                intRegistro = ObtenerItemLista(arrFrecuenciaValorizacion(), Codigo_Tipo_Frecuencia_Diaria)
                If intRegistro >= 0 Then cboFrecuencia.ListIndex = intRegistro
                cboFrecuencia.Enabled = False
                txtPlazo.Enabled = False
                cboSiNo.Enabled = False
            End If
            
            strIndComision = ""
                        
            txtPorcenParticipacion.Text = "0"
            txtPatrimonio.Text = "0": txtCapital.Text = "0"
            txtMontoEmision.Text = "0": txtMontoEmitido.Text = "0"
            txtSusMinMonto.Text = "0"
            txtSaldoMinMantenerMonto.Text = "0": txtRedMinMonto.Text = "0"
            txtRedMinCuotas.Text = "0": txtSaldoMinMantenerCuotas.Text = "0"
            txtSusMinCuotas.Text = "0"
            txtCodConasev.Text = Valor_Caracter
            txtDescripCustodio.Text = Valor_Caracter
            strCodCustodio = ""
            
            dtpHoraCorte.Value = "00:00" 'Format(dtpHoraCorte.Value, "hh:mm")
            
            cboMoneda.SetFocus
                        
        Case Reg_Edicion
            Set adoRegistro = New ADODB.Recordset
            
            strCodFondo = Trim(tdgConsulta.Columns(0))
            
            adoComm.CommandText = "SELECT * FROM Fondo WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
            Set adoRegistro = adoComm.Execute
            
            If Not adoRegistro.EOF Then
                lblCodFondo.Caption = strCodFondo
                
                intRegistro = ObtenerItemLista(arrTipoFondo(), adoRegistro("TipoFondo"))
                If intRegistro >= 0 Then cboTipoFondo.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrCodMoneda(), adoRegistro("CodMoneda"))
                If intRegistro >= 0 Then cboMoneda.ListIndex = intRegistro
                
                strIndComision = adoRegistro("IndComision")
                
                strIndComision = adoRegistro("IndComision")
                
                txtDescripFondo.Text = adoRegistro("DescripFondo")
                txtNumConasev.Text = adoRegistro("NumResolucionConasev")
                txtPatrimonio.Text = adoRegistro("MontoPatrimonioNetoMinInicio")
                txtCapital.Text = adoRegistro("MontoPatrimonioNeto")
                txtMontoEmision.Text = adoRegistro("MontoEmision")
                txtMontoEmitido.Text = adoRegistro("MontoEmitido")
                txtValorNominal.Text = adoRegistro("ValorCuotaNominal")
                txtRucFondo.Text = adoRegistro("NumRucFondo")
                
                If adoRegistro("IndFondoRegulado") = Valor_Indicador Then
                    chkIndicadorFondoRegulado.Value = vbChecked
                Else
                    chkIndicadorFondoRegulado.Value = vbUnchecked
                End If
                
                Call chkIndicadorFondoRegulado_Click
                
                If adoRegistro("IndTipoValuacionModificable") = Valor_Indicador Then
                    chkIndicadorTipoValuacionModificable.Value = vbChecked
                Else
                    chkIndicadorTipoValuacionModificable.Value = vbUnchecked
                End If
                
                Call chkIndicadorTipoValuacionModificable_Click
                
                intRegistro = ObtenerItemLista(arrRegulador(), adoRegistro("CodRegulador"))
                If intRegistro >= 0 Then cboRegulador.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrSiNo(), adoRegistro("CodAfirmacion"))
                If intRegistro >= 0 Then cboSiNo.ListIndex = intRegistro
                
                dtpFechaConasev.Value = adoRegistro("FechaResolucionConasev")
                dtpFechaPreOperativa.Value = adoRegistro("FechaInicioEtapaPreOperativa")
                dtpFechaOperativa.Enabled = True
                If adoRegistro("FechaInicioEtapaOperativa") = Valor_Fecha Then
                    dtpFechaOperativa.Value = dtpFechaPreOperativa.Value
                Else
                    dtpFechaOperativa.Value = adoRegistro("FechaInicioEtapaOperativa")
                End If
                
                txtPorcenParticipacion.Text = adoRegistro("PorcenMaxParticipe")
                txtPlazo.Text = adoRegistro("DuracionFondo")
                
                intRegistro = ObtenerItemLista(arrTipoCartera(), adoRegistro("TipoCartera"))
                If intRegistro >= 0 Then cboTipoCartera.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrTipoValuacion(), adoRegistro("TipoValuacion"))
                If intRegistro >= 0 Then cboTipoValuacion.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrFrecuenciaValorizacion(), adoRegistro("FrecuenciaValorizacion"))
                If intRegistro >= 0 Then cboFrecuencia.ListIndex = intRegistro
                                               
                strCodCustodio = adoRegistro("CodCustodio")
                
                Set adoAuxiliar = New ADODB.Recordset
                adoComm.CommandText = "SELECT IP.DescripPersona " & _
                    "FROM InstitucionPersona IP " & _
                    "WHERE CodPersona='" & strCodCustodio & "' AND TipoPersona='" & Codigo_Tipo_Persona_Emisor & "'"
                Set adoAuxiliar = adoComm.Execute
                
                If Not adoAuxiliar.EOF Then
                    txtDescripCustodio.Text = Trim(adoAuxiliar("DescripPersona"))
                End If
                adoAuxiliar.Close: Set adoAuxiliar = Nothing
                                
                txtSaldoMinMantenerCuotas.Text = adoRegistro("CantMinCuotaMantener")
                txtSaldoMinMantenerMonto.Text = adoRegistro("MontoMinSaldoMantener")
                
                txtSusMinCuotas.Text = adoRegistro("CantMinCuotaSuscripcion")
                txtSusMinMonto.Text = adoRegistro("MontoMinSuscripcion")
                
                txtRedMinCuotas.Text = adoRegistro("CantCuotaMinRedencion")
                txtRedMinMonto.Text = adoRegistro("MontoMinRedencion")
                                                                
                txtCodConasev.Text = Trim(adoRegistro("CodConasev"))
                            
                dtpHoraCorte.Value = Format(adoRegistro("HoraCorte"), "hh:mm")
                            
                            
            End If
            adoRegistro.Close: Set adoRegistro = Nothing
            
    End Select
    
End Sub


Public Sub SubImprimir(Index As Integer)

    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()

    Select Case Index
        Case 1
            gstrNameRepo = "Fondo"
                        
            Set frmReporte = New frmVisorReporte

            ReDim aReportParamS(2)
            ReDim aReportParamFn(2)
            ReDim aReportParamF(2)

            aReportParamFn(0) = "Usuario"
            aReportParamFn(1) = "Hora"
            aReportParamFn(2) = "NombreEmpresa"
            
            aReportParamF(0) = gstrLogin
            aReportParamF(1) = Format(Time(), "hh:mm:ss")
            aReportParamF(2) = gstrNombreEmpresa & Space(1)
                        
            aReportParamS(0) = strCodFondo
            aReportParamS(1) = gstrCodAdministradora
            If tabFondo.Tab = 1 And strEstado = Reg_Edicion Then
                aReportParamS(2) = Codigo_Listar_Individual
            Else
                aReportParamS(2) = Codigo_Listar_Todos
            End If
            
    End Select

    gstrSelFrml = Valor_Caracter
    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"

    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())

    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
    frmReporte.Show vbModal

    Set frmReporte = Nothing

    Screen.MousePointer = vbNormal
    
End Sub

Public Sub Modificar()

    If strEstado = Reg_Consulta Then
        strEstado = Reg_Edicion
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabFondo
            .TabEnabled(0) = False
            .Tab = 1
            .TabEnabled(1) = True
            .TabEnabled(2) = True
            .TabEnabled(3) = True
        End With
        Call Habilita
    End If
        
End Sub

Public Sub Eliminar()

    On Error GoTo CtrlError           '/**/ HMC Habilitamos la rutina de Errores Existente.
    
    Dim intAccion As Integer, lngNumError   As Long

    If strEstado = Reg_Consulta Or strEstado = Reg_Edicion Then
        If MsgBox(Mensaje_Eliminacion, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
    
            adoComm.CommandText = "UPDATE Fondo SET Estado='" & Estado_Eliminado & "' " & _
                "WHERE CodFondo='" & Trim(tdgConsulta.Columns(0)) & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
            adoConn.Execute adoComm.CommandText
            
            tabFondo.TabEnabled(0) = True
            tabFondo.Tab = 0
            
            Call Buscar
        End If
    End If

    Exit Sub
    
CtrlError:                                  '/**/
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


End Sub

Public Sub Grabar()

    Dim adoresult               As ADODB.Recordset, adoRec      As ADODB.Recordset
    Dim intAccion               As Integer, lngNumError         As Long
    Dim strCodClaseTipoCambio   As String
    Dim dblTipCambio            As Double
    Dim strFechaAnterior        As String, strFechaSiguiente    As String
    Dim datFechaFinPeriodo      As Date
    Dim adoError As ADODB.Error
    Dim strErrMsg As String
    'Dim numClientTranCount As Integer

    
    If strEstado = Reg_Consulta Then Exit Sub
    
    On Error GoTo CtrlError
    
    'numClientTranCount = 0
    
    If strEstado = Reg_Adicion Then
        Set adoRec = New ADODB.Recordset
                
        strFechaAnterior = Convertyyyymmdd(DateAdd("d", -1, dtpFechaPreOperativa.Value))
        gstrFechaActual = Convertyyyymmdd(dtpFechaPreOperativa.Value)
        gdatFechaActual = dtpFechaPreOperativa.Value
        strFechaSiguiente = Convertyyyymmdd(DateAdd("d", 1, dtpFechaPreOperativa.Value))
        
        intNumPagos = 1
        If strCodSiNo = Codigo_Respuesta_Si Then intNumPagos = 2
        
        adoComm.CommandText = "SELECT CodParametro FROM AuxiliarParametro WHERE CodTipoParametro='TIPCAM' AND ValorParametro='" & gstrClaseTipoCambioFondo & "'"
        Set adoRec = adoComm.Execute
        
        If Not adoRec.EOF Then
            strCodClaseTipoCambio = adoRec("CodParametro")
        End If
        adoRec.Close
        
        '*** Consultar Tipo de Cambio ***
        adoComm.CommandText = "SELECT ValorTipoCambioCompra,ValorTipoCambioVenta " & _
            "FROM TipoCambioFondo " & _
            "WHERE (FechaTipoCambio>='" & strFechaAnterior & "' AND FechaTipoCambio<'" & gstrFechaActual & "') AND " & _
            "CodTipoCambio='" & strCodClaseTipoCambio & "' AND CodMoneda='" & strCodMoneda & "'"
        Set adoRec = adoComm.Execute
        
        If Not adoRec.EOF Then
            If gstrValorTipoCambioOperacion = Valor_TipoCambio_Compra Then
                dblTipCambio = adoRec("ValorTipoCambioCompra")
            Else
                dblTipCambio = adoRec("ValorTipoCambioVenta")
            End If
            gdblTipoCambio = dblTipCambio
        Else
            If strCodMoneda <> Codigo_Moneda_Local Then
                MsgBox "No existe tipo de cambio vigente", vbCritical, "Fondos"
                Exit Sub
            End If
        End If
        adoRec.Close: Set adoRec = Nothing
                
        If TodoOK() Then
            Me.MousePointer = vbHourglass
            strCodFondo = Trim(lblCodFondo.Caption)
            
            'numClientTranCount = numClientTranCount + adoConn.BeginTrans()
            
'                    .CommandText = "UPDATE Fondo " & _
'            "SET HoraCorte='" & Format(dtpHoraCorte.Value, "hh:mm") & "'," & _
'            "HoraInicio='" & Format(dtpHoraInicio.Value, "hh:mm") & "'," & _
'            "HoraTermino='" & Format(dtpHoraTermino.Value, "hh:mm") & "' " & _
'            "WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
'        adoConn.Execute .CommandText
'
            
            '*** Guardar Fondo ***
            With adoComm
                .CommandText = "{ call up_GNManFondo('" & _
                    strCodFondo & "','" & gstrCodAdministradora & "','" & _
                    Trim(txtDescripFondo.Text) & "','" & strTipoFondo & "','" & _
                    strTipoCartera & "','" & Trim(txtRucFondo.Text) & "','" & _
                    strTipoValuacion & "','" & strIndTipoValuacionModificable & "','" & strFrecuenciaValorizacion & "'," & _
                    CInt(txtPlazo.Text) & ",'" & strCodCustodio & "','" & _
                    strCodMoneda & "','" & strIndRegulado & "','" & _
                    strCodRegulador & "','" & _
                    Convertyyyymmdd(dtpFechaConasev.Value) & "','" & _
                    Trim(txtNumConasev.Text) & "','" & Trim(txtCodConasev.Text) & "','" & _
                    "'," & CDec(txtValorNominal.Text) & ",'" & _
                    Convertyyyymmdd(dtpFechaPreOperativa.Value) & "','" & _
                    Convertyyyymmdd(dtpFechaOperativa.Value) & "','" & Estado_Activo & "'," & _
                    CDec(txtPorcenParticipacion.Text) & "," & CDec(txtPatrimonio.Text) & "," & _
                    CDec(txtCapital.Text) & "," & CDec(txtMontoEmision.Text) & "," & _
                    CDec(txtMontoEmitido.Text) & ",'" & Convertyyyymmdd(Date) & "',0,0," & _
                    CDec(txtRedMinCuotas.Text) & "," & CDec(txtRedMinMonto.Text) & "," & _
                    CDec(txtSaldoMinMantenerCuotas.Text) & "," & CDec(txtSaldoMinMantenerMonto.Text) & "," & _
                    CDec(txtSusMinCuotas.Text) & "," & CDec(txtSusMinMonto.Text) & "," & _
                    "0,0," & intNumPagos & ",'" & strCodSiNo & "','" & _
                    "00:00','00:00','00:00','" & Format(dtpHoraCorte.Value, "hh:mm") & "','" & _
                    gstrLogin & "','" & Convertyyyymmdd(Date) & "','" & _
                    gstrLogin & "','" & Convertyyyymmdd(Date) & "','I') }"
                adoConn.Execute .CommandText
                
                datFechaFinPeriodo = Convertddmmyyyy(Format(Year(dtpFechaPreOperativa), "0000") & "1231")
                '*** Generar Periodo Contable del Fondo y Fechas de Corte y Pago a la Administradora ***
                frmMainMdi.stbMdi.Panels(3).Text = "Generando Periodo Contable..."
                Call GenerarFondoEstructura(gstrTipoAdministradora, gstrCodAdministradora, strCodFondo, strCodMoneda, dtpFechaPreOperativa, datFechaFinPeriodo)
                'Call GenerarPeriodoContable(gstrTipoAdministradora, gstrCodAdministradora, strCodFondo, strCodMoneda, dtpFechaPreOperativa, datFechaFinPeriodo, frmMainMdi.stbMdi)
                
                '*** Parámetros del Fondo ***
                frmMainMdi.stbMdi.Panels(3).Text = "Generando Parámetros del Fondo..."
                'Call GenerarParametrosFondo(gstrCodAdministradora, strCodFondo)
                
                '*** Restricciones del Fondo ***
                frmMainMdi.stbMdi.Panels(3).Text = "Generando Restricciones del Fondo..."
                'Call GenerarRestriccionFondo(strCodFondo)
                
                '*** Periodo de Pagos del Fondo - Suscripciones ***
                frmMainMdi.stbMdi.Panels(3).Text = "Generando Periodos de Pago para Suscripciones del Fondo..."
                Call GenerarPeriodosPagoFondo(strCodFondo, intNumPagos)
                
                '*** Habilita Período Contable ***
                frmMainMdi.stbMdi.Panels(3).Text = "Habilitando Periodo Contable"
            
                .CommandText = "UPDATE PeriodoContable SET IndVigente='X' WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND PeriodoContable='" & Mid(gstrFechaActual, 1, 4) & "' AND MesContable='" & Mid(gstrFechaActual, 5, 2) & "'"
                adoConn.Execute .CommandText
            
                .CommandText = "UPDATE PeriodoContable SET IndCierre='X' WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND PeriodoContable='" & Mid(gstrFechaActual, 1, 4) & "' AND MesContable='99'"
                adoConn.Execute .CommandText
            
                '*** Habilita Fecha Vigente del Fondo ***
                frmMainMdi.stbMdi.Panels(3).Text = "Habilitando Fecha Vigente del Fondo"
                
                .CommandText = "UPDATE FondoValorCuota SET IndAbierto='X',ValorCuotaInicial=" & CDbl(txtValorNominal.Text) & "," & _
                    "ValorCuotaInicialReal=" & CDbl(txtValorNominal.Text) & ",ValorTipoCambio=" & dblTipCambio & " " & _
                    "WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                    "(FechaCuota>='" & gstrFechaActual & "' AND FechaCuota<'" & strFechaSiguiente & "')"
                adoConn.Execute .CommandText
                
                '*** Actualizar vigencia de valores ***
                'Inicio comentarios ACR: 06/04/2009
                '.CommandText = "{ call up_GNActVigenciaValores('" & strFechaSiguiente & "') }"
                'adoConn.Execute .CommandText
                'Fin comentarios ACR: 06/04/2009
            End With
            
            'numClientTranCount = numClientTranCount + adoConn.CommitTrans()
            
            '*** Fecha Vigente ***
            gdatFechaActual = Convertddmmyyyy(gstrFechaActual)
                                    
            Me.MousePointer = vbDefault
                        
            MsgBox Mensaje_Adicion_Exitosa, vbExclamation
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
            cmdOpcion.Visible = True
            With tabFondo
                .TabEnabled(0) = True
                .Tab = 0
            End With
            Call Buscar
        End If
    End If
    
    If strEstado = Reg_Edicion Then
        If TodoOK() Then
            Me.MousePointer = vbHourglass
            strCodFondo = Trim(lblCodFondo.Caption)
            
            
            With adoComm
                '*** Actualizar Fondo ***
                .CommandText = "{ call up_GNManFondo('" & _
                    strCodFondo & "','" & gstrCodAdministradora & "','" & _
                    Trim(txtDescripFondo.Text) & "','" & strTipoFondo & "','" & _
                    strTipoCartera & "','" & Trim(txtRucFondo.Text) & "','" & _
                    strTipoValuacion & "','" & strIndTipoValuacionModificable & "','" & strFrecuenciaValorizacion & "'," & _
                    CInt(txtPlazo.Text) & ",'" & strCodCustodio & "','" & _
                    strCodMoneda & "','" & strIndRegulado & "','" & _
                    strCodRegulador & "','" & _
                    Convertyyyymmdd(dtpFechaConasev.Value) & "','" & _
                    Trim(txtNumConasev.Text) & "','" & Trim(txtCodConasev.Text) & "','" & _
                    "'," & CDec(txtValorNominal.Text) & ",'" & _
                    Convertyyyymmdd(dtpFechaPreOperativa.Value) & "','" & _
                    Convertyyyymmdd(dtpFechaOperativa.Value) & "','" & Estado_Activo & "'," & _
                    CDec(txtPorcenParticipacion.Text) & "," & CDec(txtPatrimonio.Text) & "," & _
                    CDec(txtCapital.Text) & "," & CDec(txtMontoEmision.Text) & "," & _
                    CDec(txtMontoEmitido.Text) & ",'" & Convertyyyymmdd(Date) & "',0,0," & _
                    CDec(txtRedMinCuotas.Text) & "," & CDec(txtRedMinMonto.Text) & "," & _
                    CDec(txtSaldoMinMantenerCuotas.Text) & "," & CDec(txtSaldoMinMantenerMonto.Text) & "," & _
                    CDec(txtSusMinCuotas.Text) & "," & CDec(txtSusMinMonto.Text) & "," & _
                    "0,0," & intNumPagos & ",'" & strCodSiNo & "','" & _
                    "00:00','00:00','00:00','" & Format(dtpHoraCorte.Value, "hh:mm") & "','" & _
                    gstrLogin & "','" & Convertyyyymmdd(Date) & "','" & _
                    gstrLogin & "','" & Convertyyyymmdd(Date) & "','U') }"
                adoConn.Execute .CommandText
                
       
            End With

            Me.MousePointer = vbDefault
                        
            MsgBox Mensaje_Edicion_Exitosa, vbExclamation
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
            cmdOpcion.Visible = True
            With tabFondo
                .TabEnabled(0) = True
                .Tab = 0
            End With
            Call Buscar
        End If
    End If
    Exit Sub
                
CtrlError:
    If adoConn.Errors.Count > 0 Then
        For Each adoError In adoConn.Errors
            strErrMsg = strErrMsg & adoError.Description & " (" & adoError.NativeError & ") "
        Next
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

Private Sub Habilita()
    
    fraFondo(1).Enabled = True
    fraFondo(2).Enabled = True
    
End Sub

Public Sub Salir()

    Unload Me
    
End Sub



Private Sub cboFrecuencia_Click()

    strFrecuenciaValorizacion = Valor_Caracter
    If cboFrecuencia.ListIndex < 0 Then Exit Sub
    
    strFrecuenciaValorizacion = Trim(arrFrecuenciaValorizacion(cboFrecuencia.ListIndex))
    
End Sub


Private Sub cboMoneda_Click()

    strCodMoneda = Valor_Caracter
    If cboMoneda.ListIndex < 0 Then Exit Sub
    
    strCodMoneda = Trim(arrCodMoneda(cboMoneda.ListIndex))
    
    lblFondo(4).Caption = "Patrimonio Neto Mínimo Inicio Actividades" & Space(1) & ObtenerSignoMoneda(strCodMoneda)
    lblFondo(17).Caption = "Monto Emisión" & Space(1) & ObtenerSignoMoneda(strCodMoneda)
    lblFondo(9).Caption = "Monto Emitido" & Space(1) & ObtenerSignoMoneda(strCodMoneda)
    lblFondo(14).Caption = "Susc. Min. en Monto" & Space(1) & ObtenerSignoMoneda(strCodMoneda)
    lblFondo(16).Caption = "Redención Min. en Monto" & Space(1) & ObtenerSignoMoneda(strCodMoneda)
    lblFondo(12).Caption = "Saldo Min. Mantener Monto" & Space(1) & ObtenerSignoMoneda(strCodMoneda)
    lblFondo(5).Caption = "Patrimonio Neto" & Space(1) & ObtenerSignoMoneda(Codigo_Moneda_Local)
    
End Sub

Private Sub cboRegulador_Click()

    strCodRegulador = Valor_Caracter
    If cboRegulador.ListIndex < 0 Then Exit Sub
    
    strCodRegulador = Trim(arrRegulador(cboRegulador.ListIndex))
    
End Sub

Private Sub cboSiNo_Click()

    strCodSiNo = Valor_Caracter
    If cboSiNo.ListIndex < 0 Then Exit Sub
    
    strCodSiNo = Trim(arrSiNo(cboSiNo.ListIndex))
    
End Sub

Private Sub cboTipoCartera_Click()

    strTipoCartera = Valor_Caracter
    If cboTipoCartera.ListIndex < 0 Then Exit Sub
    
    strTipoCartera = Trim(arrTipoCartera(cboTipoCartera.ListIndex))
    
End Sub

Private Sub cboTipoFondo_Click()

    strTipoFondo = Valor_Caracter
    If cboTipoFondo.ListIndex < 0 Then Exit Sub
    
    strTipoFondo = Trim(arrTipoFondo(cboTipoFondo.ListIndex))
    
    If strTipoFondo = "03" Then
    
        cboFrecuencia.ListIndex = 2
        cboFrecuencia.Enabled = False
    Else
        cboFrecuencia.ListIndex = 0
        cboFrecuencia.Enabled = True
    End If
        
End Sub

Private Sub cboTipoValuacion_Click()

    strTipoValuacion = ""
    If cboTipoValuacion.ListIndex < 0 Then Exit Sub
    
    strTipoValuacion = Trim(arrTipoValuacion(cboTipoValuacion.ListIndex))
    
End Sub

Private Sub chkIndicadorFondoRegulado_Click()

    If chkIndicadorFondoRegulado.Value = vbChecked Then
        tabFondo.TabVisible(3) = True
        strIndRegulado = Valor_Indicador
    Else
        tabFondo.TabVisible(3) = False
        strIndRegulado = Valor_Caracter
    End If

End Sub

Private Sub chkIndicadorTipoValuacionModificable_Click()

    If chkIndicadorTipoValuacionModificable.Value = vbChecked Then
        strIndTipoValuacionModificable = Valor_Indicador
    Else
        strIndTipoValuacionModificable = Valor_Caracter
    End If

End Sub




Private Sub cmdBusqueda_Click()
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
        
        frmBus.Caption = " Relación de Instituciones Financieras"
        .sSql = "{ call up_ACSelDatos(22) }"
        .OutputColumns = "1,2"
        .HiddenColumns = ""
        
        .BuscarTabla
        
        Screen.MousePointer = vbNormal
        frmBus.Show 1
       
        If .iParams.Count = 0 Then Exit Sub
        
        'If .sCodigo <> "" Then
        If .iParams(1).Valor <> "" Then
            strCodCustodio = .iParams(1).Valor  '.sCodigo
            txtDescripCustodio.Text = .iParams(2).Valor '.sDescripcion
        End If
            
       
    End With
    
    Set frmBus = Nothing
    
End Sub


Private Sub dtpFechaConasev_Change()

    If Not EsDiaUtil(dtpFechaConasev.Value) Then
        MsgBox "La Fecha no es un día útil...se cambiará por una fecha correcta !", vbInformation, Me.Caption
        If dtpFechaConasev.Value >= gdatFechaActual Then
            dtpFechaConasev.Value = AnteriorDiaUtil(dtpFechaConasev.Value)
        Else
            dtpFechaConasev.Value = ProximoDiaUtil(dtpFechaConasev.Value)
        End If
    End If
    dtpFechaPreOperativa.Value = dtpFechaConasev.Value
        
End Sub

Private Sub dtpFechaOperativa_Change()

    If Not EsDiaUtil(dtpFechaOperativa.Value) Then
        MsgBox "La Fecha no es un día útil...se cambiará por una fecha correcta !", vbInformation, Me.Caption
        If dtpFechaOperativa.Value >= gdatFechaActual Then
            dtpFechaOperativa.Value = AnteriorDiaUtil(dtpFechaOperativa.Value)
        Else
            dtpFechaOperativa.Value = ProximoDiaUtil(dtpFechaOperativa.Value)
        End If
    End If
    
End Sub

Private Sub dtpFechaPreOperativa_Change()

    If Not EsDiaUtil(dtpFechaPreOperativa.Value) Then
        MsgBox "La Fecha no es un día útil...se cambiará por una fecha correcta !", vbInformation, Me.Caption
        If dtpFechaPreOperativa.Value >= gdatFechaActual Then
            dtpFechaPreOperativa.Value = AnteriorDiaUtil(dtpFechaPreOperativa.Value)
        Else
            dtpFechaPreOperativa.Value = ProximoDiaUtil(dtpFechaPreOperativa.Value)
        End If
    End If
    
End Sub

Private Sub dtpFechaPreOperativa_LostFocus()

    If Not EsDiaUtil(dtpFechaPreOperativa.Value) Then
        MsgBox "La Fecha no es un día útil...se cambiará por una fecha correcta !", vbInformation, Me.Caption
        If dtpFechaPreOperativa.Value >= gdatFechaActual Then
            dtpFechaPreOperativa.Value = AnteriorDiaUtil(dtpFechaPreOperativa.Value)
        Else
            dtpFechaPreOperativa.Value = ProximoDiaUtil(dtpFechaPreOperativa.Value)
        End If
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
    Call DarFormato
    
    
    Call ValidarPermisoUsoControl(Trim(gstrLoginUS), Me, Trim(App.Title) + Separador_Codigo_Objeto + _
    gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)
    
    CentrarForm Me
        
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call OcultarReportes
    Set frmFondos = Nothing
    
End Sub

Private Function TodoOK() As Boolean
        
    TodoOK = False
    
    If Trim(lblCodFondo.Caption) = Valor_Caracter Then
        MsgBox "El Código de Fondo no es válido.", vbCritical
        Exit Function
    End If
    
    If Trim(strTipoFondo) = Valor_Caracter Then
        MsgBox "Debe seleccionar el tipo de Fondo.", vbCritical
        tabFondo.Tab = 1
        cboTipoFondo.SetFocus
        Exit Function
    End If
    
    If Trim(strCodMoneda) = Valor_Caracter Then
        MsgBox "Debe seleccionar la moneda del Fondo.", vbCritical
        tabFondo.Tab = 1
        cboMoneda.SetFocus
        Exit Function
    End If
            
    If Trim(txtDescripFondo.Text) = Valor_Caracter Then
        MsgBox "Debe indicar la descripción del Fondo.", vbCritical
        tabFondo.Tab = 1
        txtDescripFondo.SetFocus
        Exit Function
    End If
    
    If chkIndicadorFondoRegulado.Value = vbChecked And Trim(txtNumConasev.Text) = Valor_Caracter Then
        MsgBox "Debe indicar la resolución CONASEV del Fondo", vbCritical
        tabFondo.Tab = 1
        txtNumConasev.SetFocus
        Exit Function
    End If
        
    If CDec(txtPatrimonio.Text) <= 0 Then
        If MsgBox("El Patrimonio Neto Mínimo para el inicio de actividades es 0 desea Continuar?.", vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes Then
            txtPatrimonio.SetFocus
            Exit Function
        End If
    End If
            
    If CDec(txtValorNominal.Text) <= 0 Then
        MsgBox "Ingresar el Valor Nominal de Cuota.", vbCritical
        tabFondo.Tab = 1
        txtValorNominal.SetFocus
        Exit Function
    End If
    
    If CDec(txtPorcenParticipacion.Text) <= 0 Then
        MsgBox "Ingresar porcentaje máximo de participación.", vbCritical
        txtPorcenParticipacion.SetFocus
        Exit Function
    End If
            
    If Trim(strTipoValuacion) = Valor_Caracter Then
        MsgBox "Debe seleccionar el tipo de asignación del valor cuota.", vbCritical
        tabFondo.Tab = 1
        cboTipoValuacion.SetFocus
        Exit Function
    End If
    
    If Trim(strFrecuenciaValorizacion) = Valor_Caracter Then
        MsgBox "Debe seleccionar la frecuencia de asignación de cuota.", vbCritical
        tabFondo.Tab = 1
        cboFrecuencia.SetFocus
        Exit Function
    End If
                                            
    If CDec(txtSaldoMinMantenerCuotas.Text) <> 0 And CDec(txtSaldoMinMantenerMonto.Text) <> 0 Then
        MsgBox "Ingreso de saldo a mantener debe ser en cuotas o en monto", vbCritical
        txtSaldoMinMantenerCuotas.SetFocus
        Exit Function
    End If
    
    If CDbl(txtSusMinCuotas.Text) <> 0 And CDbl(txtSusMinMonto.Text) <> 0 Then
        MsgBox "Ingreso de suscripción mínima debe ser en cuotas o en monto", vbCritical
        txtSusMinCuotas.SetFocus
        Exit Function
    End If
    
    If CDec(txtRedMinCuotas.Text) <> 0 And CDec(txtRedMinMonto.Text) <> 0 Then
        MsgBox "Ingreso de rescate mínimo debe ser en cuotas o en monto", vbCritical
        txtRedMinCuotas.SetFocus
        Exit Function
    End If
              
    If chkIndicadorFondoRegulado.Value = vbChecked And Trim(txtCodConasev.Text) = Valor_Caracter Then
        MsgBox "Ingresar el Código Conasev", vbCritical
        tabFondo.Tab = 1
        txtCodConasev.SetFocus
        Exit Function
    End If
        
    '*** Si todo paso OK ***
    TodoOK = True
  
End Function

Public Sub Adicionar()
                
    frmMainMdi.stbMdi.Panels(3).Text = "Adicionar Fondo..."
                
    strEstado = Reg_Adicion
    LlenarFormulario strEstado
    cmdOpcion.Visible = False
    With tabFondo
        .TabEnabled(0) = False
        .Tab = 1
        .TabEnabled(1) = True
        .TabEnabled(2) = True
        .TabEnabled(3) = True
    End With
    Call Habilita
      
End Sub

Private Sub tabFondo_Click(PreviousTab As Integer)

    Select Case tabFondo.Tab
        Case 1, 2
            cmdAccion.Visible = True
            If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vModify)
            If strEstado = Reg_Defecto Then tabFondo.Tab = 0
        Case 0
            cmdAccion.Visible = False
    End Select
    
End Sub

Private Sub tdgConsulta_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)

    If ColIndex = 3 Then
        Call DarFormatoValor(Value, Decimales_ValorCuota)
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

    Call OrdenarDBGrid(ColIndex, adoConsulta, tdgConsulta)
    
    numColindex = ColIndex
    
    '****
    strPrevColumTDB = strColNameTDB
    '***
    
End Sub

Private Sub txtCapital_Change()

    Call FormatoCajaTexto(txtCapital, Decimales_Monto)
    
End Sub

Private Sub txtCapital_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtCapital, Decimales_Monto)
    
End Sub

Private Sub txtCodConasev_KeyPress(KeyAscii As Integer)

     KeyAscii = ValiText(KeyAscii, "AN", 1) 'HMC

End Sub

Private Sub txtDescripFondo_KeyPress(KeyAscii As Integer)

     KeyAscii = ValiText(KeyAscii, "AN", 1) 'HMC

End Sub

Private Sub txtMontoEmision_Change()

    Call FormatoCajaTexto(txtMontoEmision, Decimales_Monto)
    
End Sub

Private Sub txtMontoEmision_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtMontoEmision, Decimales_Monto)
    
End Sub

Private Sub txtMontoEmitido_Change()

    Call FormatoCajaTexto(txtMontoEmitido, Decimales_Monto)
    
End Sub

Private Sub txtMontoEmitido_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtMontoEmitido, Decimales_Monto)
    
End Sub



Private Sub txtNumConasev_KeyPress(KeyAscii As Integer)

     KeyAscii = ValiText(KeyAscii, "AN", 1) 'HMC

End Sub

Private Sub txtPatrimonio_Change()

    Call FormatoCajaTexto(txtPatrimonio, Decimales_Monto)
    
End Sub

Private Sub txtPatrimonio_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtPatrimonio, Decimales_Monto)
    
End Sub

Private Sub txtPlazo_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "N")
    
End Sub

Private Sub txtPorcenParticipacion_Change()

    Call FormatoCajaTexto(txtPorcenParticipacion, Decimales_Tasa2)
    
End Sub

Private Sub txtPorcenParticipacion_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtPorcenParticipacion, Decimales_Tasa2)
    
End Sub

Private Sub txtRedMinCuotas_Change()

    Call FormatoCajaTexto(txtRedMinCuotas, Decimales_CantCuota)
    
End Sub

Private Sub txtRedMinCuotas_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtRedMinCuotas, Decimales_CantCuota)
    
End Sub

Private Sub txtRedMinMonto_Change()

    Call FormatoCajaTexto(txtRedMinMonto, Decimales_Monto)
    
End Sub

Private Sub txtRedMinMonto_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtRedMinMonto, Decimales_Monto)
    
End Sub



Private Sub txtRucFondo_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "N")

End Sub

Private Sub txtSaldoMinMantenerCuotas_Change()

    Call FormatoCajaTexto(txtSaldoMinMantenerCuotas, Decimales_CantCuota)
    
End Sub

Private Sub txtSaldoMinMantenerCuotas_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtSaldoMinMantenerCuotas, Decimales_CantCuota)
    
End Sub

Private Sub txtSaldoMinMantenerMonto_Change()

    Call FormatoCajaTexto(txtSaldoMinMantenerMonto, Decimales_Monto)
    
End Sub

Private Sub txtSaldoMinMantenerMonto_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtSaldoMinMantenerMonto, Decimales_Monto)
    
End Sub

Private Sub txtSusMinCuotas_Change()

    Call FormatoCajaTexto(txtSusMinCuotas, Decimales_CantCuota)
    
End Sub

Private Sub txtSusMinCuotas_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtSusMinCuotas, Decimales_CantCuota)
    
End Sub

Private Sub txtSusMinMonto_Change()

    Call FormatoCajaTexto(txtSusMinMonto, Decimales_Monto)
    
End Sub

Private Sub txtSusMinMonto_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtSusMinMonto, Decimales_Monto)
    
End Sub

Private Sub txtValorNominal_Change()

    Call FormatoCajaTexto(txtValorNominal, Decimales_ValorCuota)
    
End Sub

Private Sub txtValorNominal_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtValorNominal, Decimales_ValorCuota)
    
End Sub



