VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{830D5A06-9B70-4F8C-98B6-7A19C4E7760B}#1.0#0"; "TAMControls.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmAbonoRetiroCtaProveedor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Abono y Retiro de Cuenta de Proveedor"
   ClientHeight    =   8700
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8700
   ScaleWidth      =   10995
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
      Left            =   7830
      Picture         =   "frmAbonoRetiroCtaProveedor.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   7920
      Width           =   1215
   End
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   9270
      TabIndex        =   42
      Top             =   7920
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
      Top             =   7920
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
      Caption2        =   "&Buscar"
      Tag2            =   "5"
      ToolTipText2    =   "Buscar"
      Caption3        =   "&Anular"
      Tag3            =   "4"
      ToolTipText3    =   "Anular"
      UserControlWidth=   5700
   End
   Begin TAMControls.ucBotonEdicion cmdSalir2 
      Height          =   390
      Left            =   8640
      TabIndex        =   2
      Top             =   10080
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   688
      Caption0        =   "&Salir"
      Tag0            =   "9"
      ToolTipText0    =   "Salir"
      UserControlHeight=   390
      UserControlWidth=   1200
   End
   Begin TAMControls.ucBotonEdicion cmdOpcion2 
      Height          =   390
      Left            =   480
      TabIndex        =   1
      Top             =   10080
      Visible         =   0   'False
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   688
      Buttons         =   4
      Caption0        =   "&Nuevo"
      Tag0            =   "0"
      ToolTipText0    =   "Nuevo"
      Caption1        =   "&Modificar"
      Tag1            =   "3"
      ToolTipText1    =   "Modificar"
      Caption2        =   "&Buscar"
      Tag2            =   "5"
      ToolTipText2    =   "Buscar"
      Caption3        =   "&Anular"
      Tag3            =   "4"
      ToolTipText3    =   "Anular"
      UserControlHeight=   390
      UserControlWidth=   5700
   End
   Begin TabDlg.SSTab tabPagos 
      Height          =   7755
      Left            =   0
      TabIndex        =   4
      Top             =   120
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   13679
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
      TabPicture(0)   =   "frmAbonoRetiroCtaProveedor.frx":0671
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraCriterio"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "tdgConsulta"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "frmAbonoRetiroCtaProveedor.frx":068D
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdAccion"
      Tab(1).Control(1)=   "fraDatos"
      Tab(1).ControlCount=   2
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   -67290
         TabIndex        =   41
         Top             =   6780
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
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmAbonoRetiroCtaProveedor.frx":06A9
         Height          =   3675
         Left            =   360
         OleObjectBlob   =   "frmAbonoRetiroCtaProveedor.frx":06C3
         TabIndex        =   13
         Top             =   3630
         Width           =   10245
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
         Height          =   2985
         Left            =   360
         TabIndex        =   8
         Top             =   570
         Width           =   10245
         Begin VB.CommandButton cmdLimpiar 
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
            Left            =   9480
            Picture         =   "frmAbonoRetiroCtaProveedor.frx":6B77
            Style           =   1  'Graphical
            TabIndex        =   49
            Top             =   1020
            Width           =   375
         End
         Begin VB.CommandButton cmdBuscarProveedor 
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
            Left            =   8970
            TabIndex        =   35
            ToolTipText     =   "Buscar Proveedor"
            Top             =   1020
            Width           =   375
         End
         Begin VB.CommandButton cmdEnviarBackOffice 
            Caption         =   "Enviar"
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
            Left            =   8640
            Picture         =   "frmAbonoRetiroCtaProveedor.frx":6F35
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   1950
            Width           =   1200
         End
         Begin VB.ComboBox cboEstadoOperacionBusqueda 
            Height          =   315
            Left            =   2100
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   1830
            Width           =   2745
         End
         Begin VB.ComboBox cboFondo 
            Height          =   315
            Left            =   2100
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   540
            Width           =   7755
         End
         Begin MSComCtl2.DTPicker dtpFechaOperacionDesde 
            Height          =   315
            Left            =   2910
            TabIndex        =   43
            Top             =   2400
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
            Format          =   175964161
            CurrentDate     =   38785
         End
         Begin MSComCtl2.DTPicker dtpFechaOperacionHasta 
            Height          =   315
            Left            =   5280
            TabIndex        =   44
            Top             =   2400
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
            Format          =   175964161
            CurrentDate     =   38785
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Operacion"
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
            Left            =   390
            TabIndex        =   47
            Top             =   2430
            Width           =   1470
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
            Left            =   2130
            TabIndex        =   46
            Top             =   2430
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
            Index           =   25
            Left            =   4560
            TabIndex        =   45
            Top             =   2460
            Width           =   510
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor"
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
            TabIndex        =   39
            Top             =   1020
            Width           =   885
         End
         Begin VB.Label lblNumDocProveedor 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   4680
            TabIndex        =   38
            Top             =   1380
            Width           =   2655
         End
         Begin VB.Label lblTipoDocProv 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2100
            TabIndex        =   37
            Top             =   1380
            Width           =   2535
         End
         Begin VB.Label lblNomProveedor 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2100
            TabIndex        =   36
            Top             =   960
            Width           =   6780
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
            Index           =   5
            Left            =   360
            TabIndex        =   24
            Top             =   1860
            Width           =   1515
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
            TabIndex        =   9
            Top             =   630
            Width           =   540
         End
      End
      Begin VB.Frame fraDatos 
         Caption         =   "Datos"
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
         Height          =   6045
         Left            =   -74640
         TabIndex        =   5
         Top             =   540
         Width           =   10065
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
            Left            =   8880
            TabIndex        =   31
            ToolTipText     =   "Buscar Proveedor"
            Top             =   1170
            Width           =   375
         End
         Begin VB.TextBox txtDescripObservaciones 
            Height          =   795
            Left            =   1980
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   22
            Top             =   5010
            Width           =   7425
         End
         Begin VB.ComboBox cboTipoMov 
            Height          =   315
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   2100
            Width           =   3135
         End
         Begin VB.ComboBox cboMoneda 
            Height          =   315
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   3630
            Width           =   3105
         End
         Begin VB.TextBox txtNroVoucher 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1980
            MaxLength       =   12
            TabIndex        =   14
            Text            =   " "
            Top             =   4560
            Width           =   1830
         End
         Begin MSComCtl2.DTPicker dtpFechaActual 
            Height          =   315
            Left            =   1980
            TabIndex        =   3
            Top             =   2550
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   175964161
            CurrentDate     =   38949
         End
         Begin VB.TextBox txtMontoPago 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1980
            MaxLength       =   12
            TabIndex        =   6
            Text            =   " "
            Top             =   4110
            Width           =   1800
         End
         Begin MSComCtl2.DTPicker dtpFechaObligacion 
            Height          =   315
            Left            =   1980
            TabIndex        =   27
            Top             =   3000
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   556
            _Version        =   393216
            Format          =   175964161
            CurrentDate     =   38949
         End
         Begin VB.Label lblProveedor 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1980
            TabIndex        =   34
            Top             =   1170
            Width           =   6780
         End
         Begin VB.Label lblTipoDocID 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1980
            TabIndex        =   33
            Top             =   1590
            Width           =   2535
         End
         Begin VB.Label lblNumDocID 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   4560
            TabIndex        =   32
            Top             =   1590
            Width           =   2655
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Nro. Operación"
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
            Left            =   5820
            TabIndex        =   30
            Top             =   360
            Width           =   1290
         End
         Begin VB.Label lblNumOperacion 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "GENERADO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   7440
            TabIndex        =   29
            Top             =   330
            Width           =   1755
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Obligacion"
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
            Left            =   300
            TabIndex        =   28
            Top             =   3030
            Width           =   1500
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
            Left            =   3870
            TabIndex        =   26
            Top             =   4440
            Width           =   465
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor"
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
            Left            =   330
            TabIndex        =   21
            Top             =   1200
            Width           =   885
         End
         Begin VB.Line Line1 
            X1              =   240
            X2              =   8310
            Y1              =   3450
            Y2              =   3450
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
            Left            =   270
            TabIndex        =   20
            Top             =   2160
            Width           =   1410
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
            Index           =   8
            Left            =   330
            TabIndex        =   18
            Top             =   5070
            Width           =   1275
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
            Left            =   300
            TabIndex        =   17
            Top             =   3660
            Width           =   690
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
            Index           =   0
            Left            =   330
            TabIndex        =   15
            Top             =   4590
            Width           =   1140
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
            Index           =   4
            Left            =   330
            TabIndex        =   12
            Top             =   4170
            Width           =   540
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
            Left            =   390
            TabIndex        =   11
            Top             =   750
            Width           =   540
         End
         Begin VB.Label lblDescripFondo 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1980
            TabIndex        =   10
            Top             =   720
            Width           =   7215
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
            Left            =   300
            TabIndex        =   7
            Top             =   2580
            Width           =   1305
         End
      End
   End
End
Attribute VB_Name = "frmAbonoRetiroCtaProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strCodFondo         As String, strCodParticipe      As String
Dim strEstado           As String, strSQL               As String
Dim strTipMov           As String
Dim curMontoEmitido     As Currency, strNroVocuher      As String
Dim arrMoneda()         As String, strCodMoneda         As String, strSignoMoneda   As String, strCodSignoMoneda As String
Dim arrCuenta()         As String, strCodCuenta         As String
Dim arrTipMov()         As String, strEstadoOperacion   As String
Dim strTipoContraparte As String, strCodContraparte As String
Dim adoConsulta         As ADODB.Recordset
Dim adoRegistroAux      As ADODB.Recordset
Dim arrEstadoOperacionBusqueda()  As String
Dim strEstadoOperacionBusqueda As String
Dim strCodParticipeBusqueda As String
Dim strNumOperacion As String

Private Sub CargarReportes()

    'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Abono y Retiros de Proveedores"

End Sub


Private Sub cmdImprimir_Click()
    Call Imprimir
End Sub

Private Sub cmdLimpiar_Click()

    lblNomProveedor.Caption = ""
    lblTipoDocProv.Caption = ""
    lblNumDocProveedor.Caption = ""
    
    strCodContraparte = Valor_Caracter
    strTipoContraparte = Valor_Caracter
    
End Sub

'Private Sub cboCuenta_Click()
'    strCodCuenta = Valor_Caracter
'    If cboCuenta.ListIndex < 0 Then Exit Sub
'    strCodCuenta = Trim(arrCuenta(cboCuenta.ListIndex))
'
'End Sub
Private Sub dtpFechaOperacionDesde_Change()
    Call Buscar
End Sub

Private Sub dtpFechaOperacionHasta_Change()
    Call Buscar
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
            dtpFechaObligacion.Value = adoRegistro("FechaCuota")
        Else
            MsgBox "Periodo contable no vigente ! Debe aperturar primero un periodo contable!", vbExclamation + vbOKOnly, Me.Caption
            Exit Sub
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
        
    End With

End Sub

'Private Sub cboParticipe_Click()
'
'    gstrCodParticipe = Valor_Caracter
'    If cboParticipe.ListIndex < 0 Then Exit Sub
'
'    gstrCodParticipe = Trim(garrParticipe(cboParticipe.ListIndex))
'
'    Call Buscar
'
'End Sub



Private Sub cboMoneda_Click()
    strCodMoneda = Valor_Caracter: strSignoMoneda = Valor_Caracter
    If cboMoneda.ListIndex < 0 Then Exit Sub
    
    strCodMoneda = Trim(arrMoneda(cboMoneda.ListIndex))
    
    strSignoMoneda = ObtenerSignoMoneda(strCodMoneda)
    strCodSignoMoneda = ObtenerCodSignoMoneda(strCodMoneda)
    
    lblSignoMoneda.Caption = strCodSignoMoneda
  
'    If strCodMoneda <> Valor_Caracter And strTipMov = "01" Then
'        '*** Cuentas ***
'        strSQL = "SELECT CodCuenta CODIGO, DescripCuenta DESCRIP FROM PlanContable  " & _
'                "WHERE CodAdministradora='" & gstrCodAdministradora & "' AND CodCuenta LIKE '46911[1-2]%' AND IndMovimiento='X' AND CodMoneda ='" & strCodMoneda & "'"
'        CargarControlLista strSQL, cboCuenta, arrCuenta(), Valor_Caracter
'
'        If cboCuenta.ListCount > 0 Then cboCuenta.ListIndex = 0
'    End If
    
End Sub

Private Sub cboTipoMov_Click()
     
     strTipMov = Valor_Caracter
     If cboTipoMov.ListIndex < 0 Then Exit Sub
     strTipMov = Trim(arrTipMov(cboTipoMov.ListIndex))

End Sub

Private Sub cmdBuscarProveedor_Click()

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
            lblNomProveedor.Caption = .iParams(5).Valor
            lblTipoDocProv.Caption = .iParams(3).Valor
            lblNumDocProveedor.Caption = .iParams(4).Valor
            strCodContraparte = .iParams(1).Valor
            strTipoContraparte = .iParams(2).Valor
            'lblDireccion.Caption = .iParams(6).Valor
            'lblCodProveedor.Caption = .iParams(1).Valor
        End If
            
       
    End With
    
    Call Buscar
    
    Set frmBus = Nothing

End Sub

Private Sub cmdEnviarBackOffice_Click()

    
    Dim intContador                 As Integer
    Dim intRegistro                 As Integer
    Dim strTesoreriaOperacionXML    As String
    Dim objTesoreriaOperacionXML    As DOMDocument60
    Dim strFechaGrabar              As String
    Dim strMsgError                 As String
    Dim adoRegistro                 As ADODB.Recordset
        
    On Error GoTo ErrorHandler
                
        
    If TodoOkBackOffice() Then
        '*** Realizar proceso de contabilización ***
        If MsgBox("Datos correctos. ¿ Procedemos a enviar estas operaciones a Backoffice de Tesoreria?", vbQuestion + vbYesNo, "Observación") = vbNo Then Exit Sub
    
        intContador = tdgConsulta.SelBookmarks.Count - 1
               
        strFechaGrabar = Convertyyyymmdd(dtpFechaActual.Value) & Space(1) & Format(Time, "hh:mm")
                   
        Call ConfiguraRecordsetAuxiliar
            
        'adoRegistroAux.Open
               
        Set adoRegistro = New ADODB.Recordset
        
        With adoComm
        
            Set objTesoreriaOperacionXML = Nothing
            strTesoreriaOperacionXML = ""
                      
                      
            For intRegistro = 0 To intContador
                
                adoConsulta.MoveFirst
                
                adoConsulta.Move CLng(tdgConsulta.SelBookmarks(intRegistro) - 1), 0
                                
                tdgConsulta.Refresh
                                
                adoRegistroAux.AddNew
                adoRegistroAux.Fields("CodFondo") = strCodFondo
                adoRegistroAux.Fields("CodAdministradora") = gstrCodAdministradora
                adoRegistroAux.Fields("NumOperacion") = tdgConsulta.Columns("NumOperacion")
          
            Next
           
            Call XMLADORecordset(objTesoreriaOperacionXML, "TesoreriaOperacion", "Operacion", adoRegistroAux, strMsgError)
            strTesoreriaOperacionXML = objTesoreriaOperacionXML.xml
                
            .CommandText = "{ call up_TEProcMovimientoFondoProveedor('" & _
                strCodFondo & "','" & gstrCodAdministradora & "','" & _
                strFechaGrabar & "','" & _
                strTesoreriaOperacionXML & "') }"
            
            
            .Execute .CommandText
            
                                             
        End With
        
        MsgBox Mensaje_Proceso_Exitoso, vbExclamation
        Me.MousePointer = vbDefault
        
        frmMainMdi.stbMdi.Panels(3).Text = "Acción"
        
        'cmdOpcion.Visible = True
        With tabPagos
            .TabEnabled(0) = True
            .TabEnabled(1) = False
            .Tab = 0
        End With
        
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
            strCodContraparte = .iParams(1).Valor
            strTipoContraparte = .iParams(2).Valor
            'lblDireccion.Caption = .iParams(6).Valor
            'lblCodProveedor.Caption = .iParams(1).Valor
        End If
            
       
    End With
    
    Set frmBus = Nothing


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
        
    End Select
    
End Sub
Public Sub Imprimir()
    Call SubImprimir2(1)
End Sub

Public Sub Eliminar()

    If strEstado = Reg_Consulta Or strEstado = Reg_Edicion Then
    
        If strEstadoOperacionBusqueda = "02" Then
            MsgBox "No se puede anular un registro ya procesado ", vbOKOnly + vbCritical, Me.Caption
            Exit Sub
        ElseIf strEstadoOperacionBusqueda = "03" Then
            MsgBox "Este registro ya esta anulado ", vbOKOnly + vbCritical, Me.Caption
            Exit Sub
        End If

        strEstadoOperacion = "03"
        If MsgBox("Se procederá a anular el Abono del Proveedor." & vbNewLine & vbNewLine & "Seguro de continuar ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            adoComm.CommandText = "UPDATE TesoreriaOperacion SET EstadoOperacion='" & strEstadoOperacion & "' " & _
            "WHERE NumOperacion='" & tdgConsulta.Columns(0) & "' AND " & _
            "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
            adoConn.Execute adoComm.CommandText
            tabPagos.TabEnabled(0) = True
            tabPagos.Tab = 0
            Call Buscar
            Exit Sub
        End If
    End If
End Sub

Public Sub Salir()

    Unload Me
    
End Sub
Public Sub Cancelar()

    cmdOpcion.Visible = True
    With tabPagos
        .TabEnabled(0) = True
        .TabEnabled(1) = False
        .Tab = 0
    End With
    strEstado = Reg_Consulta
    
End Sub
Public Sub Grabar()
                        
    Dim intAccion As Integer, lngNumError   As Integer
    Dim dblMontoPago As Double, strNumSolicitud As String
    Dim strDescripOperacion As String, strNumDocumento As String
    Dim strAccion As String, strDescripObservacion As String
    Dim strFechaObligacion As String, strFechaGrabar As String
    Dim strTipoOperacion As String
    Dim mensaje As String
        
    If strEstado = Reg_Consulta Then Exit Sub
            
    On Error GoTo ErrorHandler
        
    If TodoOk() Then

        If strTipMov = "01" Then
            strDescripOperacion = "Abono de Proveedor - " + Trim(lblProveedor.Caption)
            strTipoOperacion = "35"
        ElseIf strTipMov = "02" Then
            strDescripOperacion = "Retiro de Proveedor - " + Trim(lblProveedor.Caption)
            strTipoOperacion = "36"
        ElseIf strTipMov = "03" Then
            strDescripOperacion = "Anticipo Proveedor - " + Trim(lblProveedor.Caption)
            strTipoOperacion = "39"
        End If
        
        dblMontoPago = CDbl(txtMontoPago.Text)
        
        strNumDocumento = Trim(txtNroVoucher.Text)
        
        strDescripObservacion = Trim(txtDescripObservaciones.Text)
    
        If strEstado = Reg_Adicion Then
            strAccion = "I"
            strEstadoOperacion = "01"
            mensaje = Mensaje_Adicion
        End If
        
        If strEstado = Reg_Edicion Then
            strAccion = "U"
            strEstadoOperacion = "02"
            mensaje = Mensaje_Edicion
        End If
        
        strFechaGrabar = Convertyyyymmdd(dtpFechaActual.Value) & Space(1) & Format(Time, "hh:mm")
        strFechaObligacion = Convertyyyymmdd(dtpFechaObligacion.Value)
        
        strNumOperacion = lblNumOperacion.Caption
        
    
        If MsgBox(mensaje, vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption) <> vbYes Then Exit Sub
               
        Me.MousePointer = vbHourglass
        
        ''------------------------------
        ''---Registro en BD-------------
                
        With adoComm
            
            
            .CommandText = "{ call up_TEManTesoreriaOperacion ('" & _
                                 strCodFondo & "','" & gstrCodAdministradora & "','" & strNumOperacion & "'," & _
                                 "'000', '00000000', '000'," & _
                                 "'000','" & strTipoOperacion & "','','" & strTipMov & "','" & strCodParticipe & "','" & strTipoContraparte & "','" & strCodContraparte & "'," & _
                                 "'" & strDescripOperacion & "','" & strFechaGrabar & "','" & strFechaObligacion & "'," & _
                                 "'19000101','" & strCodMoneda & "'," & dblMontoPago & "," & _
                                 "'', 0," & _
                                 "'','', 0,0,0, " & _
                                 "'','',0, " & _
                                 "'13','" & strNumDocumento & "','" & strDescripObservacion & "'," & _
                                 "'','', ''," & _
                                 "'', '',''," & _
                                 "'','', ''," & _
                                 "'', '',''," & _
                                 "'','', '', '', '', " & _
                                 "'01','" & strAccion & "') }"
                                                                                       
                                 
           .Execute
        
        End With
        Me.MousePointer = vbDefault

        MsgBox Mensaje_Adicion_Exitosa, vbExclamation
        
        'MsgBox Mensaje_Adicion_Exitosa, vbExclamation
        strCodContraparte = Valor_Caracter
        
        frmMainMdi.stbMdi.Panels(3).Text = "Acción"
        
        cmdOpcion.Visible = True
        With tabPagos
            .TabEnabled(0) = True
            .TabEnabled(1) = False
            .Tab = 0
        End With
         
        Call Buscar

 End If
 
 
    Exit Sub
 
 
ErrorHandler:
    
    If err.Number <> 0 Then
        MsgBox err.Number & " " & err.Description, vbCritical + vbOKOnly, Me.Caption
        Me.MousePointer = vbDefault
    End If
 
 
        
'    End If
'
'    frmMainMdi.stbMdi.Panels(3).Text = "Acción"
'
'    cmdOpcion.Visible = True
'    With tabPagos
'        .TabEnabled(0) = True
'        .Tab = 0
'    End With
'
'    Call Buscar
'
'    Exit Sub
'
'
'ErrorHandler:
'
'    If err.Number <> 0 Then
'        MsgBox err.Number & " " & err.Description, vbCritical + vbOKOnly, Me.Caption
'        Me.MousePointer = vbDefault
'    End If
    
'CtrlError:
'    Me.MousePointer = vbDefault
'    intAccion = ControlErrores
'    Select Case intAccion
'        Case 0: Resume
'        Case 1: Resume Next
'        Case 2: Exit Sub
'        Case Else
'            lngNumError = err.Number
'            err.Raise Number:=lngNumError
'            err.Clear
'    End Select
    
End Sub

Private Function TodoOk() As Boolean
        
    TodoOk = False
    Dim adoRegistro As ADODB.Recordset
      
    If dtpFechaActual.Value > dtpFechaObligacion.Value Then
        MsgBox "La fecha de obligacion no puede ser menor a la fecha actual!.", vbCritical, gstrNombreEmpresa
        Exit Function
    End If
        
        
    If Trim(lblProveedor.Caption) = "" Then
        MsgBox "Seleccione el Proveedor!.", vbCritical, vbCritical
        If cmdProveedor.Enabled Then cmdProveedor.SetFocus
        Exit Function
    End If
                      
        
    If CCur(txtMontoPago.Text) = 0 Then
        MsgBox "El Monto de Pago no puede ser cero!.", vbCritical, gstrNombreEmpresa
        Exit Function
    End If
    
    
    If Trim(txtNroVoucher.Text) = "" Then
        MsgBox "Ingrese el Numero de Voucher o Transaccion!.", vbCritical, gstrNombreEmpresa
        Exit Function
    End If
    
    
    If cboTipoMov.ListIndex <= -1 Then
        MsgBox "Seleccione el Tipo de Movimiento", vbCritical, gstrNombreEmpresa
        cboTipoMov.SetFocus
        Exit Function
    End If
    
    
    If cboMoneda.ListIndex <= 0 Then
        MsgBox "Seleccione una Moneda", vbCritical, gstrNombreEmpresa
        cboMoneda.SetFocus
        Exit Function
    End If
        
    
    '*** Si todo paso OK ***
    TodoOk = True
  
End Function
Private Function TodoOkBackOffice() As Boolean
        
    TodoOkBackOffice = False
        
    If tdgConsulta.SelBookmarks.Count - 1 = -1 Then
        MsgBox "Debe seleccionar registros para enviar a Backoffice!", vbCritical, gstrNombreEmpresa
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

Public Sub SubImprimir(index As Integer)

    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    Dim strFechaDesde           As String, strFechaHasta        As String
    Dim strSeleccionRegistro    As String
 
    
    Select Case index
        Case 1
        
            gstrNameRepo = "AbonoRetiroCtaProveedor"
            
            strSeleccionRegistro = "{MovimientoFondo.FechaRegistro} IN 'Fch1' TO 'Fch2'"
            gstrSelFrml = strSeleccionRegistro
            frmRangoFecha.Show vbModal
                        
            If gstrSelFrml = "0" Then
                Set frmReporte = New frmVisorReporte
    
                ReDim aReportParamS(3)
                ReDim aReportParamFn(1)
                ReDim aReportParamF(1)
                            
                aReportParamFn(0) = "Fondo"
                aReportParamFn(1) = "NombreEmpresa"
                            
                aReportParamF(0) = Trim(cboFondo.Text)
                aReportParamF(1) = gstrNombreEmpresa & Space(1)
                            
                aReportParamS(0) = strCodFondo
                aReportParamS(1) = gstrCodAdministradora
                aReportParamS(2) = Convertyyyymmdd(CVDate(gstrFchDel)) 'Convertyyyymmdd(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10))
                aReportParamS(3) = Convertyyyymmdd(DateAdd("d", 1, CVDate(gstrFchAl))) 'Convertyyyymmdd(DateAdd("d", 1, CVDate(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10))))
                
                
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

Public Sub Buscar()
        
   Dim strSQL As String
    
    Set adoConsulta = Nothing
    
    Set adoConsulta = New ADODB.Recordset
    
    If Trim(strCodContraparte) = Valor_Caracter Then
        strSQL = "SELECT " & _
                 "NumOperacion, CodFile, CodAnalitica, DescripOperacion, " & _
                 "TOPE.CodParticipe, PC.DescripPersona DescripParticipe, FechaOperacion, FechaObligacion, " & _
                 "TOPE.CodMoneda, MO.CodSigno as CodSignoMoneda, MontoOperacion, " & _
                 "TipoDocumento, NumDocumento " & _
                 "FROM TesoreriaOperacion TOPE " & _
                 "JOIN Moneda MO ON (MO.CodMoneda = TOPE.CodMoneda) " & _
                 "JOIN InstitucionPersona PC ON (PC.CodPersona = TOPE.CodPersonaContraparte " & _
                 "AND PC.TipoPersona=TOPE.TipoPersonaContraparte) " & _
                 "WHERE " & _
                 "TOPE.CodFondo = '" & strCodFondo & "' AND " & _
                 "TOPE.CodAdministradora = '" & gstrCodAdministradora & "' AND " & _
                  "TOPE.TipoOperacion in ('36','35','39') AND " & _
                 "TOPE.EstadoOperacion = '" & strEstadoOperacionBusqueda & "' AND " & _
                 "dbo.uf_ACObtenerFechaCorta(TOPE.FechaOperacion)>='" & Convertyyyymmdd(dtpFechaOperacionDesde.Value) & "' AND " & _
                 "dbo.uf_ACObtenerFechaCorta(TOPE.FechaOperacion)<='" & Convertyyyymmdd(dtpFechaOperacionHasta.Value) & "' " & _
                 "ORDER BY NumOperacion"
    Else
        strSQL = "SELECT " & _
                 "NumOperacion, CodFile, CodAnalitica, DescripOperacion, " & _
                 "TOPE.CodParticipe, PC.DescripPersona DescripParticipe, FechaOperacion, FechaObligacion, " & _
                 "TOPE.CodMoneda, MO.CodSigno as CodSignoMoneda, MontoOperacion, " & _
                 "TipoDocumento, NumDocumento " & _
                 "FROM TesoreriaOperacion TOPE " & _
                 "JOIN Moneda MO ON (MO.CodMoneda = TOPE.CodMoneda) " & _
                 "JOIN InstitucionPersona PC ON (PC.CodPersona = TOPE.CodPersonaContraparte " & _
                 "AND PC.TipoPersona=TOPE.TipoPersonaContraparte) " & _
                 "WHERE " & _
                 "TOPE.CodFondo = '" & strCodFondo & "' AND " & _
                 "TOPE.CodAdministradora = '" & gstrCodAdministradora & "' AND " & _
                 "TOPE.TipoOperacion in ('36','35','39') AND " & _
                 "TOPE.EstadoOperacion = '" & strEstadoOperacionBusqueda & "' AND " & _
                 "dbo.uf_ACObtenerFechaCorta(TOPE.FechaOperacion)>='" & Convertyyyymmdd(dtpFechaOperacionDesde.Value) & "' AND " & _
                 "dbo.uf_ACObtenerFechaCorta(TOPE.FechaOperacion)<='" & Convertyyyymmdd(dtpFechaOperacionHasta.Value) & "' AND " & _
                 "TOPE.CodPersonaContraparte='" & strCodContraparte & "' AND TipoPersonaContraparte='" & strTipoContraparte & "' " & _
                 "ORDER BY NumOperacion"
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
    
    Call AutoAjustarGrilla
    
    tdgConsulta.Refresh
    
    If adoConsulta.RecordCount > 0 Then strEstado = Reg_Consulta
   
    
End Sub

Public Sub Modificar()
    
    If strEstadoOperacionBusqueda <> Estado_Activo Then
    
        MsgBox "No se puede modificar un registro con estado diferente a Ingresado", vbOKOnly + vbCritical, Me.Caption
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
    End If
    
End Sub

Private Sub LlenarFormulario(strModo As String)

    Dim intNumSecuencial    As Integer, intRegistro As Integer
    Dim adoRegistro As New ADODB.Recordset
    Dim strSQL As String
    
    Select Case strModo
        Case Reg_Edicion
        
'            strSQL = "SELECT CodFile, CodAnalitica, CodDetalleFile, CodSubDetalleFile, " & _
'                     "TipoMovimiento, TOPE.CodParticipe, PC.DescripParticipe, DescripOperacion, FechaOperacion, " & _
'                     "FechaObligacion, FechaLiquidacion, TOPE.CodMoneda, MontoOperacion, " & _
'                     "TipoDocumento, NumDocumento, DescripObservacion, EstadoOperacion " & _
'                     "FROM TesoreriaOperacion TOPE " & _
'                     " JOIN ParticipeContrato PC ON (PC.CodParticipe = TOPE.CodParticipe) " & _
'                     " WHERE " & _
'                     "TOPE.CodFondo = '" & strCodFondo & "' AND " & _
'                     "TOPE.CodAdministradora = '" & gstrCodAdministradora & "' AND " & _
'                     "TOPE.NumOperacion = '" & tdgConsulta.Columns("NumOperacion") & "'"
'
                     
            strSQL = "SELECT CodFile, CodAnalitica, CodDetalleFile, CodSubDetalleFile, TipoMovimiento, " & _
                    "IP.CodPersona,IP.TipoPersona, IP.DescripPersona DescripProveedor, AP.DescripParametro TipoIdentidad, IP.NumIdentidad, " & _
                    "DescripOperacion, FechaOperacion,FechaObligacion, FechaLiquidacion, TOPE.CodMoneda, MontoOperacion, TipoDocumento, " & _
                    "NumDocumento , DescripObservacion, EstadoOperacion " & _
                    "FROM TesoreriaOperacion TOPE  JOIN InstitucionPersona IP ON (IP.CodPersona = TOPE.CodPersonaContraparte AND TipoPersona='04' " & _
                    "AND IP.IndVigente='X') JOIN AuxiliarParametro AP ON(AP.CodParametro=IP.TipoIdentidad AND AP.CodTipoParametro='TIPIDE') " & _
                    "WHERE TOPE.CodFondo = '" & strCodFondo & "' AND TOPE.CodAdministradora = '" & gstrCodAdministradora & "' AND " & _
                    "TOPE.NumOperacion = '" & tdgConsulta.Columns("NumOperacion") & "'"
                     
    
            adoComm.CommandText = strSQL

            Set adoRegistro = adoComm.Execute
            
            If Not adoRegistro.EOF Then

                lblDescripFondo.Caption = Trim(cboFondo.Text)
                
                lblNumOperacion.Caption = Trim(tdgConsulta.Columns("NumOperacion"))
                
                lblDescripFondo.Caption = cboFondo.Text
                
                'txtCodParticipe.Text = adoRegistro.Fields("CodParticipe")
                
                intRegistro = ObtenerItemLista(arrTipMov(), adoRegistro.Fields("TipoMovimiento"))
                If intRegistro >= 0 Then cboTipoMov.ListIndex = intRegistro

                dtpFechaObligacion.Value = Trim(adoRegistro.Fields("FechaObligacion"))
                
                lblProveedor.Caption = Trim(adoRegistro.Fields("DescripProveedor"))
                
                lblTipoDocID.Caption = Trim(adoRegistro.Fields("TipoIdentidad"))
                
                lblNumDocID.Caption = Trim(adoRegistro.Fields("NumIdentidad"))
                
                intRegistro = ObtenerItemLista(arrMoneda(), adoRegistro.Fields("CodMoneda"))
                If intRegistro >= 0 Then cboMoneda.ListIndex = intRegistro
                
                txtMontoPago.Text = Trim(adoRegistro.Fields("MontoOperacion"))
                txtNroVoucher.Text = Trim(adoRegistro.Fields("NumDocumento"))
                
                strCodContraparte = Trim(adoRegistro.Fields("CodPersona"))
                strTipoContraparte = Trim(adoRegistro.Fields("TipoPersona"))
                
                txtDescripObservaciones.Text = Trim(adoRegistro.Fields("DescripObservacion"))
                'lblDescripParticipe.Caption = adoRegistro.Fields("DescripParticipe")
            
            End If
            
            adoRegistro.Close
            
            Set adoRegistro = Nothing
        
        Case Reg_Adicion
            ''Llenar los combos del formulario
                  
            lblProveedor.Caption = Valor_Caracter
            lblTipoDocID.Caption = Valor_Caracter
            lblNumDocID.Caption = Valor_Caracter
            strCodContraparte = Valor_Caracter
            strTipoContraparte = Valor_Caracter
                  
            lblDescripFondo.Caption = Trim(cboFondo.Text)
                        
            lblNumOperacion.Caption = "GENERADO"
            
            strNumOperacion = Valor_Caracter
            
            cboMoneda.ListIndex = 0
            'txtCodParticipe.Text = Valor_Caracter
            
            'lblDescripParticipe.Caption = Valor_Caracter
            
            dtpFechaObligacion.Value = dtpFechaActual.Value
            
            cboTipoMov.ListIndex = -1
            
            txtMontoPago.Text = "0.00"
            
            txtDescripObservaciones.Text = Valor_Caracter
            
            txtNroVoucher.Text = Valor_Caracter
            
    End Select
    
End Sub

Public Sub SubImprimir2(index As Integer)
    
    
    
    Dim frmReporte    As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    
        If strCodContraparte = Valor_Caracter Then
            strCodContraparte = Valor_Comodin
        End If
        
        If strTipoContraparte = Valor_Caracter Then
            strTipoContraparte = Valor_Comodin
        End If
    
        gstrNameRepo = "AbonoRetiroCtaProveedorGrilla"
            
                Set frmReporte = New frmVisorReporte
    
                ReDim aReportParamS(7)
                ReDim aReportParamFn(3)
                ReDim aReportParamF(3)
                            
                aReportParamFn(0) = "Fondo"
                aReportParamFn(1) = "NombreEmpresa"
                aReportParamFn(2) = "Hora"
                aReportParamFn(3) = "Usuario"
                            
                aReportParamF(0) = Trim(cboFondo.Text)
                aReportParamF(1) = gstrNombreEmpresa & Space(1)
                aReportParamF(2) = Format(Time(), "hh:mm:ss")
                aReportParamF(3) = gstrLogin
                            
                aReportParamS(0) = strCodFondo
                aReportParamS(1) = gstrCodAdministradora
                aReportParamS(2) = Convertyyyymmdd(dtpFechaOperacionDesde.Value)
                aReportParamS(3) = Convertyyyymmdd(dtpFechaOperacionHasta.Value)
                aReportParamS(4) = strEstadoOperacionBusqueda
                aReportParamS(5) = strCodContraparte
                aReportParamS(6) = strTipoContraparte
                aReportParamS(7) = cboEstadoOperacionBusqueda.Text


     
    gstrSelFrml = Valor_Caracter
    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"

    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())

    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
    frmReporte.Show vbModal

    Set frmReporte = Nothing

    Screen.MousePointer = vbNormal
    
    strCodContraparte = Valor_Caracter

End Sub

Public Sub Adicionar()
    
    strEstado = Reg_Adicion
    LlenarFormulario strEstado
    cmdOpcion.Visible = False
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
    Call CargarReportes
    
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
    
    strSQL = "select CodParametro CODIGO,DescripParametro DESCRIP from AuxiliarParametro Where CodTipoParametro = 'MOVCTA' Order By DescripParametro"
    CargarControlLista strSQL, cboTipoMov, arrTipMov(), Valor_Caracter
    If cboTipoMov.ListCount > 0 Then cboTipoMov.ListIndex = 0
    
    '*** Estados ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP from AuxiliarParametro Where CodTipoParametro = 'ESACU2' Order By CodParametro"
    CargarControlLista strSQL, cboEstadoOperacionBusqueda, arrEstadoOperacionBusqueda(), Valor_Caracter
    If cboEstadoOperacionBusqueda.ListCount > 0 Then cboEstadoOperacionBusqueda.ListIndex = 0
   
    '*** Moneda ***
    strSQL = "{ call up_ACSelDatos(2) }"
    CargarControlLista strSQL, cboMoneda, arrMoneda(), Sel_Defecto
    
    If cboMoneda.ListCount > 0 Then cboMoneda.ListIndex = 0
        
End Sub
Private Sub InicializarValores()

    '*** Valores Iniciales ***
    strEstado = Reg_Defecto
    tabPagos.Tab = 0
    tabPagos.TabEnabled(1) = False
    
    '*** Ancho por defecto de las columnas de la grilla ***
'    tdgConsulta.Columns(0).Width = tdgConsulta.Width * 0.01 * 16
'    tdgConsulta.Columns(1).Width = tdgConsulta.Width * 0.01 * 9
'    tdgConsulta.Columns(2).Width = tdgConsulta.Width * 0.01 * 16
'    tdgConsulta.Columns(3).Width = tdgConsulta.Width * 0.01 * 34
    dtpFechaOperacionDesde.Value = gdatFechaActual
    dtpFechaOperacionHasta.Value = gdatFechaActual
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    
End Sub

Private Sub Form_Resize()
    Call AutoAjustarGrilla
End Sub

Private Sub Form_Unload(Cancel As Integer)
        
    Call OcultarReportes
    gstrCodParticipe = Valor_Caracter
    
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
    Call FormatoGotFocus(txtMontoPago)

End Sub

Private Sub txtMontoPago_LostFocus()
    Call FormatoCajaTexto(txtMontoPago, Decimales_Monto)
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
       .Fields.Append "NumOperacion", adVarChar, 10
       .LockType = adLockBatchOptimistic
    End With

    adoRegistroAux.Open

End Sub

Private Sub AutoAjustarGrilla()

   Dim i As Integer

  If Not adoConsulta.EOF Then
        If adoConsulta.RecordCount > 0 Then
            For i = 0 To tdgConsulta.Columns.Count - 1
            tdgConsulta.Columns(i).AutoSize
            Next
            
            tdgConsulta.Columns(0).AutoSize
            tdgConsulta.Columns(1).AutoSize
            tdgConsulta.Columns(11).AutoSize
        End If
       End If

End Sub

