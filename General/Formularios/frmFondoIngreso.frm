VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{830D5A06-9B70-4F8C-98B6-7A19C4E7760B}#1.0#0"; "TAMControls.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmFondoIngreso 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingresos del Fondo"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
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
      Left            =   8520
      Picture         =   "frmFondoIngreso.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   7320
      Width           =   1200
   End
   Begin TabDlg.SSTab tabGasto 
      Height          =   7095
      Left            =   120
      TabIndex        =   2
      Top             =   30
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   12515
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
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
      TabPicture(0)   =   "frmFondoIngreso.frx":05EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tdgConsulta"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraGastos(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "frmFondoIngreso.frx":0608
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdAcciones"
      Tab(1).Control(1)=   "fraGastos(1)"
      Tab(1).Control(2)=   "cmdAcciones2"
      Tab(1).ControlCount=   3
      Begin TAMControls2.ucBotonEdicion2 cmdAcciones 
         Height          =   735
         Left            =   -66600
         TabIndex        =   39
         Top             =   6120
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
      Begin VB.Frame fraGastos 
         Caption         =   "General"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5055
         Index           =   1
         Left            =   -74790
         TabIndex        =   9
         Top             =   630
         Width           =   11070
         Begin VB.TextBox txtNumOperacion 
            Height          =   315
            Left            =   2640
            Locked          =   -1  'True
            MaxLength       =   60
            TabIndex        =   48
            Top             =   4320
            Visible         =   0   'False
            Width           =   1785
         End
         Begin VB.CommandButton cmdOperacion 
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
            Left            =   10080
            TabIndex        =   46
            ToolTipText     =   "Buscar Proveedor"
            Top             =   4320
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtDescripOperacion 
            Height          =   315
            Left            =   4500
            Locked          =   -1  'True
            MaxLength       =   60
            TabIndex        =   45
            Top             =   4320
            Visible         =   0   'False
            Width           =   5505
         End
         Begin VB.CheckBox chkNotaCredito 
            Caption         =   "Nota de Crédito"
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
            Left            =   2640
            TabIndex        =   44
            Top             =   3870
            Width           =   2175
         End
         Begin VB.ComboBox cboEstado 
            Height          =   315
            Left            =   7350
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   2970
            Visible         =   0   'False
            Width           =   3285
         End
         Begin VB.ComboBox cboFormula 
            Height          =   315
            ItemData        =   "frmFondoIngreso.frx":0624
            Left            =   7350
            List            =   "frmFondoIngreso.frx":0631
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   3870
            Visible         =   0   'False
            Width           =   3285
         End
         Begin VB.ComboBox cboTipoCalculo 
            Height          =   315
            ItemData        =   "frmFondoIngreso.frx":065A
            Left            =   7350
            List            =   "frmFondoIngreso.frx":065C
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   3420
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.ComboBox cboMoneda 
            Height          =   315
            Left            =   2640
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   2970
            Width           =   2655
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
            Left            =   10230
            TabIndex        =   12
            ToolTipText     =   "Buscar Proveedor"
            Top             =   1800
            Width           =   375
         End
         Begin VB.ComboBox cboIngreso 
            Height          =   315
            Left            =   2640
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   960
            Width           =   7965
         End
         Begin VB.TextBox txtDescripIngreso 
            Height          =   315
            Left            =   2640
            MaxLength       =   60
            TabIndex        =   10
            Top             =   1395
            Width           =   7935
         End
         Begin MSComCtl2.DTPicker dtpFechaIngreso 
            Height          =   315
            Left            =   2640
            TabIndex        =   17
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
            Format          =   175833089
            CurrentDate     =   38068
         End
         Begin TAMControls.TAMTextBox txtMontoIngreso 
            Height          =   315
            Left            =   2640
            TabIndex        =   38
            Top             =   3420
            Width           =   1785
            _ExtentX        =   3149
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
            Container       =   "frmFondoIngreso.frx":065E
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
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Operación Relacionada"
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
            TabIndex        =   47
            Top             =   4395
            Visible         =   0   'False
            Width           =   2100
         End
         Begin VB.Label lblCodContraparte 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   7920
            TabIndex        =   34
            Top             =   2250
            Width           =   2385
         End
         Begin VB.Line Line1 
            X1              =   330
            X2              =   10740
            Y1              =   2730
            Y2              =   2730
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
            Index           =   17
            Left            =   5900
            TabIndex        =   32
            Top             =   3050
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.Label lblFormula 
            AutoSize        =   -1  'True
            Caption         =   "Formula"
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
            Left            =   5900
            TabIndex        =   31
            Top             =   3950
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Calculo"
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
            Index           =   35
            Left            =   5900
            TabIndex        =   30
            Top             =   3500
            Visible         =   0   'False
            Width           =   1080
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
            Index           =   10
            Left            =   360
            TabIndex        =   29
            Top             =   3050
            Width           =   690
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Importe del Ingreso"
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
            TabIndex        =   28
            Top             =   3500
            Width           =   1650
         End
         Begin VB.Line Line3 
            Index           =   0
            X1              =   300
            X2              =   9870
            Y1              =   810
            Y2              =   810
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
            TabIndex        =   27
            Top             =   2280
            Width           =   1230
         End
         Begin VB.Label lblNumDocID 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   5220
            TabIndex        =   26
            Top             =   2250
            Width           =   2655
         End
         Begin VB.Label lblTipoDocID 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2640
            TabIndex        =   25
            Top             =   2250
            Width           =   2535
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Contratante"
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
            TabIndex        =   24
            Top             =   1845
            Width           =   1005
         End
         Begin VB.Label lblContraparte 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2640
            TabIndex        =   23
            Top             =   1800
            Width           =   7500
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
            TabIndex        =   22
            Top             =   1020
            Width           =   825
         End
         Begin VB.Label lblAnalitica 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "060-00000000"
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
            Left            =   8490
            TabIndex        =   21
            Top             =   390
            Visible         =   0   'False
            Width           =   1395
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
            Index           =   2
            Left            =   360
            TabIndex        =   20
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
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
            TabIndex        =   19
            Top             =   405
            Width           =   540
         End
         Begin VB.Label lblMoneda 
            AutoSize        =   -1  'True
            Caption         =   "PEN"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   4560
            TabIndex        =   18
            Top             =   3480
            Width           =   330
         End
      End
      Begin VB.Frame fraGastos 
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
         ForeColor       =   &H00000000&
         Height          =   1695
         Index           =   0
         Left            =   390
         TabIndex        =   3
         Top             =   510
         Width           =   10995
         Begin VB.CheckBox chkVerSoloIngresosContabilizados 
            Caption         =   "Ver sólo ingresos contabilizados"
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
            Left            =   6960
            TabIndex        =   42
            Top             =   1050
            Width           =   3435
         End
         Begin VB.ComboBox cboFondoSerie 
            Enabled         =   0   'False
            Height          =   315
            Left            =   6270
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   1770
            Visible         =   0   'False
            Width           =   4185
         End
         Begin VB.CheckBox chkVerSoloIngresosVigentes 
            Caption         =   "Ver sólo ingresos vigentes"
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
            Left            =   420
            TabIndex        =   5
            Top             =   1050
            Width           =   2595
         End
         Begin VB.ComboBox cboFondo 
            Height          =   315
            Left            =   1710
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   450
            Width           =   6315
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Serie"
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
            ForeColor       =   &H00800000&
            Height          =   315
            Index           =   34
            Left            =   4740
            TabIndex        =   8
            Top             =   1770
            Visible         =   0   'False
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
            Index           =   0
            Left            =   360
            TabIndex        =   7
            Top             =   510
            Width           =   540
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmFondoIngreso.frx":067A
         Height          =   4575
         Left            =   390
         OleObjectBlob   =   "frmFondoIngreso.frx":0694
         TabIndex        =   33
         Top             =   2340
         Width           =   10995
      End
      Begin TAMControls.ucBotonEdicion cmdAcciones2 
         Height          =   390
         Left            =   -66600
         TabIndex        =   37
         Top             =   6330
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
   End
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   10080
      TabIndex        =   41
      Top             =   7320
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
      Caption0        =   "&Salir"
      Tag0            =   "9"
      ToolTipText0    =   "Salir"
      UserControlWidth=   1200
   End
   Begin TAMControls2.ucBotonEdicion2 cmdOpcion 
      Height          =   735
      Left            =   480
      TabIndex        =   40
      Top             =   7320
      Width           =   4200
      _ExtentX        =   7408
      _ExtentY        =   1296
      Buttons         =   3
      Caption0        =   "&Nuevo "
      Tag0            =   "0"
      ToolTipText0    =   "Nuevo"
      Caption1        =   "&Modificar"
      Tag1            =   "3"
      ToolTipText1    =   "Modificar"
      Caption2        =   "&Anular"
      Tag2            =   "4"
      ToolTipText2    =   "Anular"
      UserControlWidth=   4200
   End
   Begin TAMControls.ucBotonEdicion cmdOpcion2 
      Height          =   390
      Left            =   480
      TabIndex        =   0
      Top             =   8400
      Visible         =   0   'False
      Width           =   4200
      _ExtentX        =   7408
      _ExtentY        =   688
      Buttons         =   3
      Caption0        =   "&Nuevo"
      Tag0            =   "0"
      ToolTipText0    =   "Nuevo"
      Caption1        =   "&Modificar"
      Tag1            =   "3"
      ToolTipText1    =   "Modificar"
      Caption2        =   "&Eliminar"
      Tag2            =   "4"
      ToolTipText2    =   "Eliminar"
      UserControlHeight=   390
      UserControlWidth=   4200
   End
   Begin TAMControls.ucBotonEdicion cmdSalir2 
      Height          =   390
      Left            =   10080
      TabIndex        =   1
      Top             =   8400
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   688
      Caption0        =   "&Salir"
      Tag0            =   "9"
      ToolTipText0    =   "Salir"
      UserControlHeight=   390
      UserControlWidth=   1200
   End
   Begin TAMControls.ucBotonEdicion cmdAccion 
      Height          =   1125
      Left            =   6570
      TabIndex        =   35
      Top             =   5100
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
      UserControlHeight=   1125
      UserControlWidth=   2700
   End
   Begin TAMControls.ucBotonEdicion ucBotonEdicion2 
      Height          =   1125
      Left            =   7800
      TabIndex        =   36
      Top             =   5790
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
      UserControlHeight=   1125
      UserControlWidth=   2700
   End
End
Attribute VB_Name = "frmFondoIngreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrFondo()                  As String, arrTipoProvision()           As String
Dim arrMoneda()                 As String, arrFrecuenciaDevengo()       As String
Dim arrCuentaGasto()            As String, arrGasto()                   As String
Dim arrTipoPago()               As String, arrCreditoFiscal()           As String
Dim arrTipoValor()              As String, arrEstado()                  As String
Dim arrPeriodoPago()            As String, arrTipoDesplazamiento()      As String
Dim arrModalidadPago()          As String, arrTipoDevengo()             As String
Dim arrAplicacionDevengo()      As String, arrPeriodoTasa()             As String
Dim arrTipoTasa()               As String, arrTipoGasto()               As String
Dim arrBaseCalculo()            As String, arrFondoSerie()              As String
Dim arrFrecuenciaGasto()        As String, arrFormula()                 As String
Dim arrTipoCalculo()            As String, arrModalidadDevengo()        As String
Dim arrPeriodoGasto()           As String, arrPeriodoDevengo()          As String

Dim strCodFondo                 As String, strCodTipoProvision          As String
Dim strCodMoneda                As String, strCodFrecuenciaDevengo      As String
Dim strCodCuenta                As String, strCodGasto                  As String
Dim strCodFile                  As String, strCodAnalitica              As String
Dim strCodTipoPago              As String, strCodDetalleGasto           As String
Dim strCodCreditoFiscal         As String, strCodTipoDesplazamiento     As String
Dim strEstado                   As String, strSQL                       As String
Dim strCodTipoValor             As String, strCodPeriodoPago            As String
Dim strCodModalidadPago         As String, strCodTipoDevengo            As String
Dim strCodAplicacionDevengo     As String, strCodPeriodoTasa            As String
Dim strCodTipoTasa              As String, strCodTipoGasto              As String
Dim strCodBaseCalculo           As String, strEstadoIngreso             As String
Dim strCodFondoSerie            As String, strCodFormula                As String
Dim strCodFrecuenciaGasto       As String, strIndGastoIterativo         As String
Dim strCodPeriodoGasto          As String, strCodPeriodoDevengo         As String

Dim intNumPeriodo               As Integer, strFechaInicio              As String
Dim strFechaFin                 As String, strFechaPago                 As String
Dim intCantDias                 As Long, strIndVigente                  As String
Dim intSecuencialIngreso        As Integer, intNumSecuencial            As Integer
Dim strCodTipoCalculo           As String, strCodModalidadDevengo       As String
Dim strNumAnexo                 As String, strCodLimiteCli              As String
Dim strCodEstructura            As String, dblIGVImporte                As Double
Dim strNumContrato              As String, strNumDocumentoFisico        As String
'******BMM NUEVOS CAMBIOS******
Dim adoConsulta As ADODB.Recordset
'******************************
Public Sub Buscar()
    
    Set adoConsulta = New ADODB.Recordset
               
    strSQL = "SELECT FI.CodCuenta,FI.NumIngreso,FI.CodAnalitica,FI.FechaDefinicion,PCG.DescripCuenta, FI.DescripIngreso,FI.MontoIngreso,FI.CodFile,INP.DescripPersona as DescripContratante,NumOrdenCobro,Estado " & _
        "FROM FondoIngreso FI " & _
        "JOIN OrdenCobro OC ON (FI.CodFondo=OC.CodFondo AND FI.CodAdministradora=OC.CodAdministradora AND FI.NumIngreso=OC.NumIngreso) " & _
        "LEFT JOIN FondoConceptoIngreso FCG ON(FCG.CodCuenta=FI.CodCuenta AND FCG.CodAdministradora=FI.CodAdministradora AND FCG.CodFondo=FI.CodFondo) " & _
        "LEFT JOIN PlanContable PCG ON(PCG.CodCuenta=FI.CodCuenta) " & _
        "JOIN InstitucionPersona INP ON(INP.CodPersona=FI.CodContratante AND INP.TipoPersona = '" & Codigo_Tipo_Persona_Emisor & "') " & _
        "WHERE FI.CodFondo='" & strCodFondo & "' AND FI.CodAdministradora='" & gstrCodAdministradora & "'"

    If chkVerSoloIngresosVigentes.Value = vbChecked Then
        strSQL = strSQL & " AND (FI.IndVigente = 'X') "
        If chkVerSoloIngresosContabilizados.Value = vbChecked Then
            strSQL = strSQL & " AND (Estado = '04') "
        Else
            strSQL = strSQL & " AND (Estado <> '04') "
        End If
    Else
        strSQL = strSQL & " AND (FI.IndVigente='') "
    End If

    strSQL = strSQL & " ORDER BY FI.NumIngreso"

    strEstado = Reg_Defecto
    
    With adoConsulta
        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open strSQL
    End With
    
    tdgConsulta.DataSource = adoConsulta
    
    Call AutoAjustarGrillas
    
    tdgConsulta.Refresh

    If adoConsulta.RecordCount > 0 Then strEstado = Reg_Consulta

            
End Sub

Private Sub CargarIngresos()
    '*** Gastos del Fondo ***
    strSQL = "SELECT FCI.CodCuenta CODIGO,(RTRIM(DescripCuenta)) DESCRIP " & _
        "FROM FondoConceptoIngreso FCI JOIN PlanContable PCG ON(PCG.CodCuenta=FCI.CodCuenta AND PCG.CodAdministradora=FCI.CodAdministradora) " & _
        "WHERE CodFondo='" & strCodFondo & "' AND FCI.CodAdministradora='" & gstrCodAdministradora & "' " & _
        "ORDER BY DescripCuenta"
    CargarControlLista strSQL, cboIngreso, arrGasto(), Sel_Defecto
    
End Sub

Public Sub Salir()

    Unload Me
    
End Sub

Private Sub cboEstado_Click()

    strEstadoIngreso = Valor_Caracter
    If cboEstado.ListIndex < 0 Then Exit Sub
    
    strEstadoIngreso = Trim(arrEstado(cboEstado.ListIndex))

End Sub

Private Sub cboFondoSerie_Click()
    strCodFondoSerie = ""
    If cboFondoSerie.ListIndex < 0 Then Exit Sub
    
    strCodFondoSerie = Trim(arrFondoSerie(cboFondoSerie.ListIndex))
    
    Call Buscar
End Sub

'Private Sub cboFormula_Click()
'    strCodFormula = Valor_Caracter
'    If cboFormula.ListIndex < 0 Then Exit Sub
'
'    strCodFormula = Trim(arrFormula(cboFormula.ListIndex))
'End Sub

Private Sub cboFondo_Click()

    Dim adoRegistro     As ADODB.Recordset
    
    strCodFondo = Valor_Caracter
    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(arrFondo(cboFondo.ListIndex))
    
    'Cargamos las series del fondo
'    strSQL = "{ call up_ACSelDatosParametro(50,'" & gstrCodAdministradora & "','" & strCodFondo & "') }"
'    CargarControlLista strSQL, cboFondoSerie, arrFondoSerie(), Valor_Caracter
    
    If cboFondoSerie.ListCount > 0 Then cboFondoSerie.ListIndex = 0
    
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Fecha Vigente, Valor Cuota, Valor Cuota T-1, Moneda y Cantidad Inicial de Cuotas del Fondo ***
        .CommandText = "{ call up_ACSelDatosParametro(23,'" & strCodFondo & "','" & gstrCodAdministradora & "','000') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            dtpFechaIngreso.Value = adoRegistro("FechaCuota")
            strCodMoneda = adoRegistro("CodMoneda")
            
            gdatFechaActual = adoRegistro("FechaCuota")
            frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    Call Buscar
    
End Sub


'Private Sub cboFrecuenciaDevengo_Click()
'
'    strCodFrecuenciaDevengo = Valor_Caracter
'    If cboFrecuenciaDevengo.ListIndex < 0 Then Exit Sub
'
'    strCodFrecuenciaDevengo = Trim(arrFrecuenciaDevengo(cboFrecuenciaDevengo.ListIndex))
'
'End Sub


Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Listado de Ingresos del Fondo Vigentes"
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Text = "Listado de Ingresos del Fondo No Vigentes"
    
End Sub

Private Sub cboIngreso_Click()

    strCodGasto = Valor_Caracter ': strCodAnalitica = Valor_Caracter
    If cboIngreso.ListIndex <= 0 Then Exit Sub
    
    strCodGasto = Trim(arrGasto(cboIngreso.ListIndex))

End Sub


Private Sub cboMoneda_Click()

    strCodMoneda = Valor_Caracter
    If cboMoneda.ListIndex < 0 Then Exit Sub
    
    strCodMoneda = Trim(arrMoneda(cboMoneda.ListIndex))
    
    lblMoneda.Caption = ObtenerCodSignoMoneda(strCodMoneda)
    
End Sub

Private Sub chkNotaCredito_Click()

    txtMontoIngreso.Text = txtMontoIngreso.Value * -1
    If chkNotaCredito.Value = vbChecked Then
        lblDescrip(3).Caption = "Importe Nota de Crédito"
    Else
        lblDescrip(3).Caption = "Importe del Ingreso"
        
        txtNumOperacion.Text = Valor_Caracter
        txtDescripOperacion.Text = Valor_Caracter
        strNumAnexo = Valor_Caracter
        strCodFile = "060"
        strCodAnalitica = Valor_Caracter
        lblAnalitica.Caption = "060-????????"
        strCodLimiteCli = Valor_Caracter
        strCodEstructura = Valor_Caracter
        strNumContrato = Valor_Caracter
        strNumDocumentoFisico = Valor_Caracter

    End If
    
    lblDescrip(4).Visible = chkNotaCredito.Value = vbChecked
    txtDescripOperacion.Visible = chkNotaCredito.Value = vbChecked
    cmdOperacion.Visible = chkNotaCredito.Value = vbChecked
    txtNumOperacion.Visible = chkNotaCredito.Value = vbChecked
    
End Sub

Private Sub chkVerSoloIngresosContabilizados_Click()
    If chkVerSoloIngresosContabilizados.Value = vbChecked Then
        chkVerSoloIngresosVigentes.Value = vbChecked
    End If
    
    Call Buscar
End Sub

'Private Sub cboTipoCalculo_Click()
'
'    strCodTipoCalculo = Valor_Caracter
'
'    If cboTipoCalculo.ListIndex < 0 Then Exit Sub
'
'    strCodTipoCalculo = Trim(arrTipoCalculo(cboTipoCalculo.ListIndex))
'
'    If strCodTipoCalculo = Tipo_Calculo_Variable Then
'        lblFormula.Visible = True
''        cboFormula.Visible = True
'        txtMontoIngreso.Enabled = False
'        txtMontoIngreso.Text = "0"
'    Else
'        lblFormula.Visible = False
''        cboFormula.Visible = False
'        cboFormula.ListIndex = -1
'        txtMontoIngreso.Enabled = True
'    End If
'
'End Sub

Private Sub chkVerSoloIngresosVigentes_Click()
    If chkVerSoloIngresosVigentes.Value = vbUnchecked Then
        chkVerSoloIngresosContabilizados.Value = Unchecked
        chkVerSoloIngresosContabilizados.Enabled = False
    Else
        chkVerSoloIngresosContabilizados.Enabled = True
    End If
    
    Call Buscar
End Sub



Private Sub cmdOperacion_Click()

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
        
        frmBus.Caption = " Relación de Operaciones"
        .sSql = "select NumOperacion,DescripOperacion, ValorNominalDscto, IP.DescripPersona, IO.NumAnexo, " & _
                " CodFile, CodAnalitica, CodLimiteCli, CodEstructura, NumContrato, NumDocumentoFisico " & _
                " from InversionOperacion IO " & _
                " JOIN InstitucionPersona IP on (IO.CodEmisor = IP.CodPersona and IP.TipoPersona = '02') " & _
                " where CodFondo = '" & gstrCodFondoContable & "' and TipoOperacion = '01' and CodEmisor = '" & lblCodContraparte.Caption & "'"
        
        .OutputColumns = "1,2,3,4,5,6,7,8,9,10,11"
        .HiddenColumns = "6,7,8,9,10,11"
        
        .BuscarTabla
        
        Screen.MousePointer = vbNormal
        frmBus.Show 1
       
        If .iParams.Count = 0 Then Exit Sub
        
        If .iParams(1).Valor <> "" Then
            txtNumOperacion.Text = .iParams(1).Valor
            txtDescripOperacion.Text = .iParams(2).Valor
            strNumAnexo = .iParams(5).Valor
            strCodFile = .iParams(6).Valor
            strCodAnalitica = .iParams(7).Valor
            lblAnalitica.Caption = strCodFile & "-" & strCodAnalitica
            strCodLimiteCli = .iParams(8).Valor
            strCodEstructura = .iParams(9).Valor
            strNumContrato = .iParams(10).Valor
            strNumDocumentoFisico = .iParams(11).Valor
       Else
            txtNumOperacion.Text = Valor_Caracter
            txtDescripOperacion.Text = Valor_Caracter
            strNumAnexo = Valor_Caracter
            strCodFile = "060"
            strCodAnalitica = Valor_Caracter
            lblAnalitica.Caption = "060-????????"
            strCodLimiteCli = Valor_Caracter
            strCodEstructura = Valor_Caracter
            strNumContrato = Valor_Caracter
            strNumDocumentoFisico = Valor_Caracter
        End If
            
       
    End With
    
    Set frmBus = Nothing

End Sub

Private Sub cmdProveedor_Click()

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
        
        frmBus.Caption = " Relación de Clientes"
        .sSql = "{ call up_ACSelDatos(32) }"
        
        .OutputColumns = "1,2,3,4,5,6"
        .HiddenColumns = "1,2,6"
        
        .BuscarTabla
        
        Screen.MousePointer = vbNormal
        frmBus.Show 1
       
        If .iParams.Count = 0 Then Exit Sub
        
        If .iParams(1).Valor <> "" Then
            lblContraparte.Caption = .iParams(5).Valor
            lblTipoDocID.Caption = .iParams(3).Valor
            lblNumDocID.Caption = .iParams(4).Valor
            'lblDireccion.Caption = .iParams(6).Valor
            'lblCodProveedor.Caption = .iParams(1).Valor
            lblCodContraparte.Caption = .iParams(1).Valor
            strCodAnalitica = lblCodContraparte.Caption
            lblAnalitica.Caption = "060-" & lblCodContraparte.Caption
       Else
            strCodAnalitica = ""
            lblAnalitica.Caption = "060-????????"
            
        End If
            
       
    End With
    
    Set frmBus = Nothing
    


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
    Call DarFormato
    Call Buscar
    
    Call ValidarPermisoUsoControl(Trim(gstrLoginUS), Me, Trim(App.Title) + Separador_Codigo_Objeto + _
    gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)
    
    CentrarForm Me
        
    Call ValidaExisteTipoCambio(Codigo_TipoCambio_SBS, gstrFechaActual)
 

End Sub

Private Sub DarFormato()

    Dim intCont As Integer
    
'    For intCont = 0 To (lblDescrip.Count - 1)
'        Call FormatoEtiqueta(lblDescrip(intCont), vbLeftJustify)
'    Next
    Dim c As Object
    Dim elemento As Object

    For Each c In Me.Controls
        
        If TypeOf c Is Label Then
            Call FormatoEtiqueta(c, vbLeftJustify)
        End If
    Next
    
    For intCont = 0 To (fraGastos.Count - 1)
        Call FormatoMarco(fraGastos(intCont))
    Next
    
    For Each elemento In Me.Controls
    
        If TypeOf elemento Is TDBGrid Then
            Call FormatoGrilla(elemento)
        End If
    
    Next
            
End Sub
Private Sub CargarListas()
        
    Dim intRegistro         As Integer
    
    '*** Fondos ***
    strSQL = "{ call up_ACSelDatosParametro(74,'" & gstrCodAdministradora & "','" & gstrCodFondoContable & "') }"
    CargarControlLista strSQL, cboFondo, arrFondo(), Valor_Caracter
    
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
        
    '*** Tipo de Desplazamiento ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPCAC' ORDER BY CodParametro"
    CargarControlLista strSQL, cboTipoCalculo, arrTipoCalculo(), Valor_Caracter
    
    '*** Moneda ***
    strSQL = "{ call up_ACSelDatos(2) }"
    CargarControlLista strSQL, cboMoneda, arrMoneda(), Valor_Caracter
        
    '*** Estados ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP From AuxiliarParametro WHERE CodTipoParametro='INDREG' AND CodParametro<>'03' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboEstado, arrEstado(), Valor_Caracter
    
    '*** Formulas
'    strSQL = "SELECT CodFormula CODIGO,DescripFormula DESCRIP From Formula ORDER BY DescripFormula"
'    CargarControlLista strSQL, cboFormula, arrFormula(), Valor_Caracter
        
End Sub
Private Sub InicializarValores()
                        
    '*** Valores Iniciales ***
    tabGasto.Tab = 0
    strCodFile = "060"
    
    tabGasto.TabEnabled(1) = False
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgConsulta.Columns(2).Width = tdgConsulta.Width * 0.01 * 6
    tdgConsulta.Columns(3).Width = tdgConsulta.Width * 0.01 * 10
    tdgConsulta.Columns(5).Width = tdgConsulta.Width * 0.01 * 15
    tdgConsulta.Columns(6).Width = tdgConsulta.Width * 0.01 * 30
    
'    Set cmdSalir.FormularioActivo = Me
'    Set cmdAccion.FormularioActivo = Me
'    Set cmdOpciones.FormularioActivo = Me
'
    chkVerSoloIngresosVigentes.Value = 1
    
    Set cmdSalir.FormularioActivo = Me
    Set cmdAcciones.FormularioActivo = Me
    Set cmdOpcion.FormularioActivo = Me
    
    
End Sub

Private Sub Form_Resize()
    Call AutoAjustarGrillas
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call OcultarReportes
    Set frmFondoGastos = Nothing
    
End Sub

Public Sub Accion(nAccion As ButtonAction)
    
    Select Case nAccion
        
        Case vNew
            Call Adicionar
        Case vModify
            Call Modificar
        Case vDelete
            Call Eliminar
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

Public Sub Cancelar()

    cmdOpcion.Visible = True
    With tabGasto
        .TabEnabled(0) = True
        .TabEnabled(1) = False
        .Tab = 0
    End With
    strEstado = Reg_Consulta
    
End Sub

Public Sub Grabar()

    Dim adoRegistro                     As ADODB.Recordset, adoRec      As ADODB.Recordset
    Dim adoConsulta                     As ADODB.Recordset
    Dim intAccion                       As Integer, lngNumError         As Long
    Dim dblTipCambio                    As Double
    Dim strFechaAnterior                As String, strFechaSiguiente    As String
    Dim datFechaFinPeriodo              As Date
    Dim numMontoIngreso                 As Double
    Dim strIndNoIncluyeBalancePreCierre As String
    Dim strIndFinMes                    As String
    Dim strIndSinVencimiento            As String
    Dim intDiasProvision                As String
    Dim intDiasBaseAnual                As Integer
    Dim intNumPeriodoAnualTasa          As Integer
    Dim strEstadoReg                    As String
    Dim mensaje                         As String
    Dim strTipoPersona                  As String
    Dim dblPorcenIGV                    As Double
    Dim dblIGV                          As Double
    Dim indAfectoIGV                    As Integer
    
    If strEstado = Reg_Consulta Then Exit Sub
    If Not TodoOK() Then Exit Sub
    
'    numMontoIngreso = CDbl(txtMontoIngreso.Text)
    
    strEstadoReg = "I"
    mensaje = Mensaje_Adicion
    
    If strEstado = Reg_Edicion Then
        strEstadoReg = "U"
        mensaje = Mensaje_Edicion
    End If
    
    Set adoRegistro = New ADODB.Recordset
                           
    If MsgBox(mensaje, vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption) <> vbYes Then Exit Sub
                           
    Me.MousePointer = vbHourglass
    
'    intSecuencialIngreso = 0
    
    '*** Guardar ***
    With adoComm
        '*** Obtener el número secuencial *** Se paso este codigo al SP
'        .CommandText = "SELECT MAX(NumIngreso) NumSecuencial FROM FondoIngreso " & _
'            "WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
'        Set adoRegistro = .Execute
'
'        If Not adoRegistro.EOF Then
'            If IsNull(adoRegistro("NumSecuencial")) Then
'                intSecuencialIngreso = 1
'            Else
'                intSecuencialIngreso = CInt(adoRegistro("NumSecuencial")) + 1
'            End If
'        Else
'            intSecuencialIngreso = 1
'        End If
        
'        adoRegistro.Close: Set adoRegistro = Nothing
    
'        .CommandText = "{ call up_GNManFondoIngreso('" & strCodFondo & "','" & _
'            gstrCodAdministradora & "'," & intSecuencialIngreso & ",'" & _
'            strCodFondoSerie & "','" & Convertyyyymmdd(dtpFechaIngreso.Value) & "','" & strCodGasto & "','" & _
'            strCodFile & "','" & strCodAnalitica & "','" & Codigo_Tipo_Persona_Contratante & "','" & _
'            Trim(lblCodContraparte.Caption) & "','" & Trim(txtDescripIngreso.Text) & "','" & _
'            Convertyyyymmdd(CVDate(Valor_Fecha)) & "','','" & strEstadoIngreso & "'," & _
'            gdblTipoCambio & ",'" & strCodMoneda & "','" & strCodTipoCalculo & "','" & _
'            strCodFormula & "'," & numMontoIngreso & ",'" & strEstadoReg & "') }"
        If chkNotaCredito.Value = vbChecked Then
            strCodGasto = Valor_Caracter
            strTipoPersona = Codigo_Tipo_Persona_Emisor
            dblPorcenIGV = gdblTasaIgv * 100
            dblIGVImporte = gdblTasaIgv * CDec(txtMontoIngreso.Text)
            indAfectoIGV = 1
        Else
            strTipoPersona = Codigo_Tipo_Persona_Cliente
            txtNumOperacion.Text = Valor_Caracter
            txtDescripOperacion.Text = Valor_Caracter
            strNumAnexo = Valor_Caracter
            strCodFile = "060"
            strCodLimiteCli = Valor_Caracter
            strCodEstructura = Valor_Caracter
            strNumContrato = Valor_Caracter
            dblPorcenIGV = 0
            indAfectoIGV = 0
            strNumDocumentoFisico = Valor_Caracter

        End If
        
        .CommandText = "{ call up_GNManFondoIngreso('" & strCodFondo & "','" & _
            gstrCodAdministradora & "'," & intSecuencialIngreso & ",'" & _
            Convertyyyymmdd(dtpFechaIngreso.Value) & "','" & strCodGasto & "','" & _
            strCodFile & "','" & strCodAnalitica & "','" & strTipoPersona & "','" & _
            Trim(lblCodContraparte.Caption) & "','" & Trim(txtDescripIngreso.Text) & "','" & _
            Convertyyyymmdd(CVDate(Valor_Fecha)) & "','','" & strEstadoIngreso & "'," & _
            gdblTipoCambio & ",'" & strCodMoneda & "'," & _
            CDec(txtMontoIngreso.Text) & "," & dblPorcenIGV & "," & indAfectoIGV & ", " & dblIGVImporte & ", '" & strNumAnexo & "', '" & strCodLimiteCli & "', '" & strCodEstructura & "', '" & strNumContrato & "', '" & strNumDocumentoFisico & "', '" & txtNumOperacion.Text & "','" & IIf(strEstadoReg = "I", "", tdgConsulta.Columns(9)) & "','" & strEstadoReg & "') }"
        
       
        adoConn.Execute .CommandText

    End With
                                                                                                        
    Me.MousePointer = vbDefault
                
    If strEstado = Reg_Adicion Then
        MsgBox Mensaje_Adicion_Exitosa, vbExclamation
    Else
        MsgBox Mensaje_Edicion_Exitosa, vbExclamation
    End If
    
    frmMainMdi.stbMdi.Panels(3).Text = "Acción"
    
    cmdOpcion.Visible = True
    With tabGasto
        .TabEnabled(0) = True
        .TabEnabled(1) = False
        .Tab = 0
    End With
   
    Call cboFondo_Click
    
    Exit Sub
                
CtrlError:
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

Public Sub Imprimir()

    Call SubImprimir(1)
    
End Sub

Public Sub SubImprimir(Index As Integer)

    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    Dim strFechaDesde           As String, strFechaHasta        As String
    Dim strSeleccionRegistro    As String
    
    gstrNameRepo = "FondoIngreso"
    
    strSeleccionRegistro = "{Participe.FechaIngreso} IN 'Fch1' TO 'Fch2'"
    gstrSelFrml = strSeleccionRegistro
    frmRangoFecha.Show vbModal
            
    
    If gstrSelFrml <> "0" Then
    
        Set frmReporte = New frmVisorReporte

        ReDim aReportParamS(4)
        ReDim aReportParamFn(4)
        ReDim aReportParamF(4)

        aReportParamFn(0) = "Usuario"
        aReportParamFn(1) = "Hora"
        aReportParamFn(2) = "NombreEmpresa"
        aReportParamFn(3) = "Fondo"
        aReportParamFn(4) = "Titulo"
    
        aReportParamF(0) = gstrLogin
        aReportParamF(1) = Format(Time(), "hh:mm:ss")
        aReportParamF(2) = gstrNombreEmpresa & Space(1)
        aReportParamF(3) = Trim(cboFondo.Text)
                
        aReportParamS(0) = strCodFondo
        aReportParamS(1) = gstrCodAdministradora
    
            Select Case Index
                Case 1
                    aReportParamF(4) = "LISTADO DE INGRESOS VIGENTES"
                    aReportParamS(2) = Valor_Indicador
                Case 2
                    aReportParamF(4) = "LISTADO DE INGRESOS NO VIGENTES"
                    aReportParamS(2) = Valor_Caracter
            End Select
    
   ' aReportParamS(3) = Convertyyyymmdd(gstrFchDel)
   ' aReportParamS(4) = Convertyyyymmdd(gstrFchAl)
    
    aReportParamS(3) = Convertyyyymmdd(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10))
    aReportParamS(4) = Convertyyyymmdd(DateAdd("d", 1, CVDate(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10))))
    
    End If
    
   If gstrSelFrml <> "0" Then
   
   gstrSelFrml = Valor_Caracter
   
   frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"

   Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())

   frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
   frmReporte.Show vbModal

   Set frmReporte = Nothing

   Screen.MousePointer = vbNormal
   
   End If
       
End Sub

Private Sub cmdImprimir_Click()
    
    Call SubImprimir2(1)
    
End Sub

Public Sub SubImprimir2(Index As Integer)
    
    
Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    
    gstrNameRepo = "FondoIngresoGrilla"
    
        Set frmReporte = New frmVisorReporte

        ReDim aReportParamS(5)
        ReDim aReportParamFn(4)
        ReDim aReportParamF(4)
    
        aReportParamFn(0) = "Usuario"
        aReportParamFn(1) = "Hora"
        aReportParamFn(2) = "NombreEmpresa"
        aReportParamFn(3) = "Fondo"
        aReportParamFn(4) = "Titulo"
        
        aReportParamF(0) = gstrLogin
        aReportParamF(1) = Format(Time(), "hh:mm:ss")
        aReportParamF(2) = gstrNombreEmpresa & Space(1)
        aReportParamF(3) = Trim(cboFondo.Text)
                    
        aReportParamS(0) = strCodFondo
        aReportParamS(1) = gstrCodAdministradora
        aReportParamS(2) = Codigo_Tipo_Persona_Emisor
       
       
         If chkVerSoloIngresosVigentes.Value = vbChecked Then
         
            aReportParamS(3) = "X"
            
            If chkVerSoloIngresosContabilizados.Value = vbChecked Then
                aReportParamS(4) = "1"
                aReportParamS(5) = "04"
            Else
                aReportParamS(4) = "0"
                aReportParamS(5) = "04"
             End If
             
        Else
             aReportParamS(3) = ""
             aReportParamS(4) = ""
             aReportParamS(5) = "%"
        End If
       
          
    
    gstrSelFrml = Valor_Caracter
    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"

    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())

    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
    frmReporte.Show vbModal

    Set frmReporte = Nothing

    Screen.MousePointer = vbNormal

    
End Sub


Public Sub Eliminar()
    
    
    Dim adoRegistro As New ADODB.Recordset
    
    If chkVerSoloIngresosVigentes.Value = 0 Then
        MsgBox "Este registro ya esta anulado", vbOKOnly + vbCritical, Me.Caption
        Exit Sub
    End If
        
    If tdgConsulta.Columns(10).Value = "04" Then
        MsgBox "No se puede anular un registro que ya ah sido confirmado en un Registro de Venta", vbOKOnly + vbCritical, Me.Caption
        Exit Sub
    End If
    
    If strEstado = Reg_Consulta Or strEstado = Reg_Edicion Then
        If MsgBox("Se procederá a eliminar el Ingreso del Fondo." & vbNewLine & vbNewLine & "Seguro de continuar ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            adoComm.CommandText = "UPDATE FondoIngreso SET IndVigente='" & Valor_Caracter & "' " & _
                        "WHERE NumIngreso=" & CInt(tdgConsulta.Columns(2)) & " AND CodCuenta='" & tdgConsulta.Columns(1) & "' AND " & _
                        "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
                        adoConn.Execute adoComm.CommandText
            
            
            adoComm.CommandText = "UPDATE OrdenCobro SET Estado='03' WHERE NumIngreso='" & CInt(tdgConsulta.Columns(2)) & "' " & _
                                "AND CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
                        adoConn.Execute adoComm.CommandText
                                
            
            tabGasto.TabEnabled(0) = True
            tabGasto.TabEnabled(1) = False
            tabGasto.Tab = 0
            Call Buscar
            Exit Sub
        End If
    End If
End Sub
Public Sub Modificar()
    
    Dim adoRegistro As New ADODB.Recordset
    
    If strEstado = Reg_Consulta Then
    
        If chkVerSoloIngresosVigentes.Value = 0 Then
            MsgBox "No se puede modificar un registro anulado", vbOKOnly + vbCritical, Me.Caption
            Exit Sub
        End If
        
        If tdgConsulta.Columns(10).Value = "04" Then
            MsgBox "No se puede modificar un registro que ya ah sido confirmado en un Registro de Venta", vbOKOnly + vbCritical, Me.Caption
            Exit Sub
        End If
        
        strEstado = Reg_Edicion
        
        LlenarFormulario strEstado
        
        cmdOpcion.Visible = False
        
        With tabGasto
            .TabEnabled(0) = False
            .TabEnabled(1) = True
            .Tab = 1
        End With
        
    End If
        
End Sub

Private Sub LlenarFormulario(strModo As String)

    Dim adoRegistro   As ADODB.Recordset
    Dim intRegistro As Integer
    
    Select Case strModo
        
        Case Reg_Adicion
            
            fraGastos(1).Caption = "Fondo : " & Trim(cboFondo.Text) & Space(1) & "-" & Space(1) & "Serie : " & Trim(cboFondoSerie.Text) '& "Tipo Gasto : " & Trim(cboTipoProvision.Text)
            intSecuencialIngreso = 0
            lblAnalitica.Caption = "060-????????"
            txtDescripIngreso.Text = Valor_Caracter
            txtMontoIngreso.Text = "0"
            dtpFechaIngreso.Value = gdatFechaActual
            dtpFechaIngreso.Enabled = False
                        
            lblContraparte.Caption = Valor_Caracter
            lblCodContraparte.Caption = Valor_Caracter
            lblTipoDocID.Caption = Valor_Caracter
            lblNumDocID.Caption = Valor_Caracter
            
            txtNumOperacion.Text = Valor_Caracter
            txtDescripOperacion.Text = Valor_Caracter
            strCodFile = "060"
            strCodAnalitica = Valor_Caracter
            
            Call CargarIngresos
            
            cboIngreso.ListIndex = -1
            If cboIngreso.ListCount > 0 Then cboIngreso.ListIndex = 0
                                  
            'Gastos
            intRegistro = ObtenerItemLista(arrMoneda(), strCodMoneda)
            If intRegistro >= 0 Then cboMoneda.ListIndex = intRegistro

            intRegistro = ObtenerItemLista(arrEstado(), Valor_Indicador)
            If intRegistro >= 0 Then cboEstado.ListIndex = intRegistro
                        
'            intRegistro = ObtenerItemLista(arrTipoCalculo(), Tipo_Calculo_Fijo)
'            If intRegistro >= 0 Then cboTipoCalculo.ListIndex = intRegistro
            
            cboIngreso.SetFocus
                        
        Case Reg_Edicion
        
            Call CargarIngresos
            
            Set adoRegistro = New ADODB.Recordset
            
            If tdgConsulta.AllowRowSelect = True Then
                adoComm.CommandText = "SELECT FI.NumIngreso, FI.FechaDefinicion, FI.CodCuenta, FI.CodFile, FI.CodAnalitica, " & _
                    "FI.TipoPersona, FI.CodContratante, FI.DescripIngreso, FI.FechaConfirma, FI.IndConfirma, " & _
                    "FI.IndVigente, FI.ValorTipoCambio, FI.CodMoneda, FI.MontoIngreso, " & _
                    "AP.DescripParametro TipoIdentidad,INP.DescripPersona,INP.NumIdentidad " & _
                    "FROM FondoIngreso FI " & _
                    "JOIN InstitucionPersona INP ON(INP.CodPersona=FI.CodContratante AND INP.TipoPersona = FI.TipoPersona) " & _
                    "JOIN AuxiliarParametro AP ON (AP.CodParametro = INP.TipoIdentidad AND AP.CodTipoParametro = 'TIPIDE') " & _
                    "WHERE NumIngreso=" & CInt(tdgConsulta.Columns("NumIngreso")) & " AND CodCuenta='" & tdgConsulta.Columns("CodCuenta") & "' AND " & _
                    "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
                Set adoRegistro = adoComm.Execute
            Else
                MsgBox "Por lo menos debe haber un ingreso registrado", vbExclamation, gstrNombreSistema
                Exit Sub
            End If
            
            If Not adoRegistro.EOF Then
                fraGastos(1).Caption = "Fondo : " & Trim(cboFondo.Text) & Space(1) & "-" & Space(1) & "Serie : " & Trim(cboFondoSerie.Text) ' & "Tipo Gasto : " & Trim(cboTipoProvision.Text)
                
                intSecuencialIngreso = adoRegistro("NumIngreso")
                
                strCodFile = Trim(adoRegistro("CodFile"))
                strCodAnalitica = Trim(adoRegistro("CodAnalitica"))
                lblAnalitica.Caption = Trim(adoRegistro("CodFile")) & "-" & Trim(adoRegistro("CodAnalitica"))
                txtDescripIngreso.Text = Trim(adoRegistro("DescripIngreso"))
            
                intRegistro = ObtenerItemLista(arrGasto(), adoRegistro("CodCuenta"))
                If intRegistro >= 0 Then cboIngreso.ListIndex = intRegistro
                
                lblContraparte.Caption = Trim(adoRegistro("DescripPersona"))
                lblTipoDocID.Caption = adoRegistro("TipoIdentidad")
                lblNumDocID.Caption = adoRegistro("NumIdentidad")
                lblCodContraparte.Caption = adoRegistro("CodContratante")
               
                dtpFechaIngreso.Value = adoRegistro("FechaDefinicion")
                                   
                intRegistro = ObtenerItemLista(arrEstado(), adoRegistro("IndVigente"))
                If intRegistro >= 0 Then cboEstado.ListIndex = intRegistro
                                   
                'Montos
                intRegistro = ObtenerItemLista(arrMoneda(), adoRegistro("CodMoneda"))
                If intRegistro >= 0 Then cboMoneda.ListIndex = intRegistro
                                   
                txtMontoIngreso.Text = CStr(adoRegistro("MontoIngreso"))
                
'                intRegistro = ObtenerItemLista(arrTipoCalculo(), adoRegistro("CodTipoCalculo"))
'                If intRegistro >= 0 Then cboTipoCalculo.ListIndex = intRegistro
            
'                intRegistro = ObtenerItemLista(arrFormula(), "" & adoRegistro("CodFormula"))
'                If intRegistro >= 0 Then cboFormula.ListIndex = intRegistro
                                                               
            End If
            adoRegistro.Close: Set adoRegistro = Nothing
    
    End Select
    
End Sub

Public Sub Adicionar()
                
    If strCodFondo = Valor_Caracter Then
        MsgBox "No existen fondos definidos...", vbCritical, Me.Caption
        Exit Sub
    End If
    
    frmMainMdi.stbMdi.Panels(3).Text = "Adicionar Ingresos a la SIV..."
                
    strEstado = Reg_Adicion
    LlenarFormulario strEstado
    cmdOpcion.Visible = False
    With tabGasto
        .TabEnabled(0) = False
        .TabEnabled(1) = True
        .Tab = 1
    End With
      
End Sub
Private Function TodoOK()

    TodoOK = False
    
    If Trim(strCodGasto) = Valor_Caracter Then
        MsgBox "Debe Seleccionar el Ingreso.", vbCritical
        cboIngreso.SetFocus
        Exit Function
    End If
      
'    If Trim(strCodCreditoFiscal) = Valor_Caracter Then
'        MsgBox "Debe Seleccionar el Tipo de Crédito Fiscal.", vbCritical
'        cboCreditoFiscal.SetFocus
'        Exit Function
'    End If
    
    If Trim(txtDescripIngreso.Text) = Valor_Caracter Then
        MsgBox "Debe Ingresar la Descripción del Ingreso.", vbCritical
        txtDescripIngreso.SetFocus
        Exit Function
    End If
    
    If cboMoneda.ListIndex = -1 Then
        MsgBox "Debe Ingresar la Moneda del Ingreso.", vbCritical
        cboMoneda.SetFocus
        Exit Function
    End If
                
    If CDec(txtMontoIngreso.Text) = 0 Then
        MsgBox "El Valor del Ingreso no Puede Ser 0.", vbCritical
        txtMontoIngreso.SetFocus
        Exit Function
    End If
    
    If Trim(lblCodContraparte.Caption) = "" Then
        MsgBox "Debe Indicar el Relacionado al Ingreso.", vbCritical
        txtMontoIngreso.SetFocus
        Exit Function
    End If
    
    If chkNotaCredito.Value = vbChecked Then
        If txtNumOperacion.Text = Valor_Caracter Then
            MsgBox "Debe seleccionar una operación relacionada.", vbCritical
            txtMontoIngreso.SetFocus
            Exit Function
        End If
    End If
    
'    If cboAplicacionDevengo.ListIndex = -1 Then
'        MsgBox "Debe Seleccionar el Tipo de Aplicación de Devengo del Gasto.", vbCritical
'        cboAplicacionDevengo.SetFocus
'        Exit Function
'    End If
    
'    If strCodAplicacionDevengo = Codigo_Aplica_Devengo_Periodica And cboFrecuenciaDevengo.ListIndex = -1 Then
'        MsgBox "Debe Seleccionar la Frecuencia de Aplicación de Devengo del Gasto.", vbCritical
'        cboFrecuenciaDevengo.SetFocus
'        Exit Function
'    End If
    
    '*** Si todo paso OK ***
    TodoOK = True
  
End Function


Private Sub tabGasto_Click(PreviousTab As Integer)

    Select Case tabGasto.Tab
        Case 1, 2
            cmdAccion.Visible = True
            If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vQuery)
            If strEstado = Reg_Defecto Then tabGasto.Tab = 0
        Case 0
            'If PreviousTab > 0 Then
            cmdAccion.Visible = False
            'End If
    End Select
    
End Sub

Private Sub tdgConsulta_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)

    If ColIndex = 7 Then
        Call DarFormatoValor(Value, Decimales_Monto)
    End If
    
End Sub

'Private Sub txtMontoIngreso_Change()
'
'    Call FormatoCajaTexto(txtMontoIngreso, Decimales_Monto)
'
'End Sub


'Private Sub txtMontoIngreso_KeyPress(KeyAscii As Integer)
'
'    Call ValidaCajaTexto(KeyAscii, "M", txtMontoIngreso, Decimales_Monto)
'
'End Sub

Private Sub tdgConsulta_HeadClick(ByVal ColIndex As Integer)

    Static numColindex As Integer

    tdgConsulta.Splits(0).Columns(numColindex).HeadingStyle.ForegroundPicture = Null

    Call OrdenarDBGrid(ColIndex, adoConsulta, tdgConsulta)
    
    numColindex = ColIndex

End Sub


Private Sub AutoAjustarGrillas()
    
    Dim i As Integer
    
    If Not adoConsulta.EOF Then
        If adoConsulta.RecordCount > 0 Then
            For i = 1 To tdgConsulta.Columns.Count - 1
                tdgConsulta.Columns(i).AutoSize
            Next
            
            tdgConsulta.Columns(9).AutoSize
        End If
    End If
    
    tdgConsulta.Refresh

End Sub

Private Sub txtMontoIngreso_LostFocus()
    If chkNotaCredito.Value = vbChecked And txtMontoIngreso.Value > 0 Then
        txtMontoIngreso.Text = txtMontoIngreso.Value * -1
    End If
End Sub

