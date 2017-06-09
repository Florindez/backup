VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{830D5A06-9B70-4F8C-98B6-7A19C4E7760B}#1.0#0"; "TAMControls.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmComprobanteCobro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Comprobantes de Ventas"
   ClientHeight    =   10800
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10800
   ScaleWidth      =   10785
   ShowInTaskbar   =   0   'False
   Begin TAMControls.ucBotonEdicion cmdAccion 
      Height          =   390
      Left            =   4500
      TabIndex        =   0
      Top             =   8190
      Visible         =   0   'False
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   688
      Buttons         =   4
      Caption0        =   "&Guardar"
      Tag0            =   "2"
      ToolTipText0    =   "Guardar"
      Caption1        =   "&Imprimir"
      Tag1            =   "6"
      ToolTipText1    =   "Imprimir"
      Caption2        =   "&Anular"
      Tag2            =   "7"
      ToolTipText2    =   "Anular"
      Caption3        =   "&Cancelar"
      Tag3            =   "8"
      ToolTipText3    =   "Cancelar"
      UserControlHeight=   390
      UserControlWidth=   5700
   End
   Begin TAMControls.ucBotonEdicion cmdSalir 
      Height          =   390
      Left            =   8940
      TabIndex        =   1
      Top             =   8940
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   688
      Caption0        =   "&Salir"
      Tag0            =   "9"
      ToolTipText0    =   "Salir"
      UserControlHeight=   390
      UserControlWidth=   1200
   End
   Begin TabDlg.SSTab tabRegistroCompras 
      Height          =   8520
      Left            =   120
      TabIndex        =   2
      Top             =   210
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   15028
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
      TabCaption(0)   =   "Lista"
      TabPicture(0)   =   "frmComprobanteCobro.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "gLista"
      Tab(0).Control(1)=   "ucBotonNavegacion1"
      Tab(0).Control(2)=   "fraCompras(0)"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Datos Generales"
      TabPicture(1)   =   "frmComprobanteCobro.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "ucBotonEdicion21"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraCompras(3)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "fraCompras(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Ordenes de Cobro"
      TabPicture(2)   =   "frmComprobanteCobro.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame fraCompras 
         Caption         =   "Criterios de búsqueda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Index           =   0
         Left            =   -74640
         TabIndex        =   48
         Top             =   420
         Width           =   9705
         Begin VB.ComboBox cboEstado 
            Height          =   315
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   67
            Top             =   740
            Width           =   2055
         End
         Begin VB.ComboBox cboFondo 
            Height          =   315
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   50
            Top             =   360
            Width           =   6255
         End
         Begin MSComCtl2.DTPicker dtpFechaDesde 
            Height          =   345
            Left            =   3600
            TabIndex        =   49
            Top             =   1200
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   609
            _Version        =   393216
            Format          =   179896321
            CurrentDate     =   39042
         End
         Begin MSComCtl2.DTPicker dtpFechaHasta 
            Height          =   345
            Left            =   7200
            TabIndex        =   51
            Top             =   1200
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   609
            _Version        =   393216
            Format          =   179896321
            CurrentDate     =   39042
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
            Index           =   9
            Left            =   840
            TabIndex        =   66
            Top             =   840
            Width           =   600
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
            Left            =   840
            TabIndex        =   55
            Top             =   360
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
            Index           =   1
            Left            =   2520
            TabIndex        =   54
            Top             =   1260
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
            Index           =   2
            Left            =   6000
            TabIndex        =   53
            Top             =   1260
            Width           =   510
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Registro"
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
            Index           =   15
            Left            =   840
            TabIndex        =   52
            Top             =   1290
            Width           =   1575
         End
      End
      Begin VB.Frame fraCompras 
         Caption         =   "Definición del Comprobante"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4395
         Index           =   1
         Left            =   240
         TabIndex        =   23
         Top             =   420
         Width           =   9885
         Begin VB.ComboBox cboIngreso 
            Height          =   315
            Left            =   2280
            Style           =   2  'Dropdown List
            TabIndex        =   63
            Top             =   3450
            Width           =   7215
         End
         Begin VB.TextBox txtNumComprobante 
            Height          =   315
            Left            =   7680
            MaxLength       =   10
            TabIndex        =   29
            Top             =   1260
            Width           =   1815
         End
         Begin VB.ComboBox cboTipoComprobante 
            Height          =   315
            Left            =   2280
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   825
            Width           =   7215
         End
         Begin VB.TextBox txtDescripcion 
            Height          =   315
            Left            =   2280
            MaxLength       =   800
            TabIndex        =   27
            Top             =   3870
            Width           =   7215
         End
         Begin VB.ComboBox cboMoneda 
            Height          =   315
            Left            =   2280
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   1695
            Width           =   2295
         End
         Begin VB.CommandButton cmdContratante 
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
            TabIndex        =   25
            ToolTipText     =   "Buscar Contratante"
            Top             =   2145
            Width           =   375
         End
         Begin VB.TextBox txtSerieComprobante 
            Height          =   315
            Left            =   6960
            MaxLength       =   3
            TabIndex        =   24
            Top             =   1260
            Width           =   615
         End
         Begin MSComCtl2.DTPicker dtpFechaRegistro 
            Height          =   315
            Left            =   6960
            TabIndex        =   30
            Top             =   390
            Width           =   2535
            _ExtentX        =   4471
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
            Format          =   179896321
            CurrentDate     =   39042
         End
         Begin MSComCtl2.DTPicker dtpFechaComprobante 
            Height          =   315
            Left            =   2280
            TabIndex        =   31
            Top             =   1260
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
            Format          =   179896321
            CurrentDate     =   39042
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
            Index           =   7
            Left            =   360
            TabIndex        =   64
            Top             =   3510
            Width           =   825
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Num. Comprobante"
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
            Left            =   5160
            TabIndex        =   47
            Top             =   1290
            Width           =   1620
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Num. Registro"
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
            TabIndex        =   46
            Top             =   465
            Width           =   1215
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
            Index           =   5
            Left            =   360
            TabIndex        =   45
            Top             =   2220
            Width           =   1005
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
            Index           =   6
            Left            =   360
            TabIndex        =   44
            Top             =   3930
            Width           =   1005
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Dirección"
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
            Left            =   360
            TabIndex        =   43
            Top             =   3105
            Width           =   810
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Comprobante"
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
            Left            =   360
            TabIndex        =   42
            Top             =   900
            Width           =   1560
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
            Index           =   12
            Left            =   360
            TabIndex        =   41
            Top             =   1755
            Width           =   690
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
            Index           =   13
            Left            =   360
            TabIndex        =   40
            Top             =   2670
            Width           =   1230
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
            Index           =   17
            Left            =   5160
            TabIndex        =   39
            Top             =   465
            Width           =   540
         End
         Begin VB.Label lblNumSecuencial 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2280
            TabIndex        =   38
            Top             =   390
            Width           =   2295
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Comprobante"
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
            TabIndex        =   37
            Top             =   1320
            Width           =   1710
         End
         Begin VB.Label lblContratante 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2280
            TabIndex        =   36
            Top             =   2145
            Width           =   6600
         End
         Begin VB.Label lblDireccion 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2280
            TabIndex        =   35
            Top             =   3030
            Width           =   7200
         End
         Begin VB.Label lblCodContratante 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   5550
            TabIndex        =   34
            Top             =   1830
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Label lblNumDocID 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   4170
            TabIndex        =   33
            Top             =   2580
            Width           =   2655
         End
         Begin VB.Label lblTipoDocID 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2280
            TabIndex        =   32
            Top             =   2580
            Width           =   1815
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Detalle del Cobro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6735
         Left            =   -74640
         TabIndex        =   10
         Top             =   630
         Width           =   9735
         Begin VB.CommandButton cmdCobro 
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
            Left            =   9120
            TabIndex        =   62
            ToolTipText     =   "Buscar Proveedor"
            Top             =   510
            Width           =   315
         End
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   7140
            TabIndex        =   16
            Top             =   5910
            Width           =   2295
         End
         Begin VB.TextBox txtIgv 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   7140
            TabIndex        =   15
            Top             =   5520
            Width           =   2295
         End
         Begin VB.TextBox txtSubTotal 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   7140
            TabIndex        =   14
            Top             =   5145
            Width           =   2295
         End
         Begin VB.ComboBox cboCobro 
            Height          =   315
            Left            =   1830
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   510
            Width           =   7215
         End
         Begin VB.CommandButton cmdAdicionarCobro 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   12
            Top             =   1140
            Width           =   435
         End
         Begin VB.CommandButton cmdEliminarCobro 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   11
            Top             =   1710
            Width           =   435
         End
         Begin DXDBGRIDLibCtl.dxDBGrid gIngresos 
            Height          =   3795
            Left            =   810
            OleObjectBlob   =   "frmComprobanteCobro.frx":0054
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   1110
            Width           =   8640
         End
         Begin VB.Label lblTasaIGV 
            AutoSize        =   -1  'True
            Caption         =   "-"
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
            Left            =   5760
            TabIndex        =   65
            Top             =   5550
            Width           =   75
         End
         Begin VB.Label lblSignoMonedaVV 
            AutoSize        =   -1  'True
            Caption         =   "-"
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
            Left            =   6690
            TabIndex        =   61
            Top             =   5190
            Width           =   75
         End
         Begin VB.Label lblSignoMonedaIGV 
            AutoSize        =   -1  'True
            Caption         =   "-"
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
            Left            =   6690
            TabIndex        =   60
            Top             =   5580
            Width           =   75
         End
         Begin VB.Label lblSignoMonedaPV 
            AutoSize        =   -1  'True
            Caption         =   "-"
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
            Left            =   6690
            TabIndex        =   59
            Top             =   5940
            Width           =   75
         End
         Begin VB.Label lblTotalLetras 
            Height          =   255
            Left            =   150
            TabIndex        =   58
            Top             =   6300
            Width           =   9255
         End
         Begin VB.Label lblMontoIngreso 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1560
            TabIndex        =   57
            Top             =   5700
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Label lblPV 
            AutoSize        =   -1  'True
            Caption         =   "Precio de Venta"
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
            Left            =   5160
            TabIndex        =   21
            Top             =   5925
            Width           =   1380
         End
         Begin VB.Label lblIGV 
            AutoSize        =   -1  'True
            Caption         =   "IGV"
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
            Left            =   5160
            TabIndex        =   20
            Top             =   5550
            Width           =   330
         End
         Begin VB.Label lblVV 
            AutoSize        =   -1  'True
            Caption         =   "Valor de Venta"
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
            Left            =   5160
            TabIndex        =   19
            Top             =   5175
            Width           =   1275
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Orden de Cobro"
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
            Left            =   240
            TabIndex        =   18
            Top             =   570
            Width           =   1350
         End
      End
      Begin VB.Frame fraCompras 
         Caption         =   "Definición de Pagos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   3
         Left            =   240
         TabIndex        =   3
         Top             =   4830
         Width           =   9885
         Begin VB.ComboBox cboAfectacion 
            Height          =   315
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   300
            Width           =   2295
         End
         Begin VB.ComboBox cboCreditoFiscal 
            Height          =   315
            ItemData        =   "frmComprobanteCobro.frx":46D0
            Left            =   2520
            List            =   "frmComprobanteCobro.frx":46D7
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   705
            Width           =   2295
         End
         Begin VB.TextBox txtPeriodoFiscal 
            Height          =   315
            Left            =   7080
            TabIndex        =   4
            Top             =   300
            Width           =   2295
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Impuesto"
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
            Left            =   360
            TabIndex        =   9
            Top             =   405
            Width           =   780
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Crédito Fiscal"
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
            Left            =   360
            TabIndex        =   8
            Top             =   750
            Width           =   1155
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Periodo Registro Crédito Fiscal"
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
            Left            =   5400
            TabIndex        =   7
            Top             =   330
            Width           =   1455
         End
      End
      Begin TAMControls.ucBotonNavegacion ucBotonNavegacion1 
         Height          =   30
         Left            =   -69450
         TabIndex        =   22
         Top             =   4920
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   53
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gLista 
         Height          =   4785
         Left            =   -74640
         OleObjectBlob   =   "frmComprobanteCobro.frx":46ED
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   2340
         Width           =   9690
      End
      Begin TAMControls2.ucBotonEdicion2 ucBotonEdicion21 
         Height          =   735
         Left            =   4380
         TabIndex        =   72
         Top             =   7230
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   1296
         Buttons         =   4
         Caption0        =   "&Guardar"
         Tag0            =   "2"
         ToolTipText0    =   "Guardar"
         Caption1        =   "&Imprimir"
         Tag1            =   "6"
         ToolTipText1    =   "Imprimir"
         Caption2        =   "Anular"
         Tag2            =   "7"
         ToolTipText2    =   "Anular"
         Caption3        =   "Cancelar"
         Tag3            =   "8"
         ToolTipText3    =   "Cancelar"
         UserControlWidth=   5700
      End
   End
   Begin TAMControls2.ucBotonEdicion2 cmdOpcion 
      Height          =   735
      Left            =   330
      TabIndex        =   70
      Top             =   9510
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   1296
      Buttons         =   4
      Caption0        =   "&Nuevo"
      Tag0            =   "0"
      ToolTipText0    =   "Nuevo"
      Caption1        =   "&Modificar"
      Tag1            =   "1"
      ToolTipText1    =   "Modificar"
      Caption2        =   "&Eliminar"
      Tag2            =   "4"
      ToolTipText2    =   "Eliminar"
      Caption3        =   "&Buscar"
      Tag3            =   "5"
      ToolTipText3    =   "Buscar"
      UserControlWidth=   5700
   End
End
Attribute VB_Name = "frmComprobanteCobro"
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
Dim arrFormaPagoDetraccion()    As String, arrCobro()                   As String
Dim arrDetraccion()             As String, arrTipoValorCambio()         As String
Dim arrEstado()                 As String

Dim strCodFondo                 As String, strCodMoneda                 As String
Dim strCodTipoComprobante       As String, strCodMonedaUnico            As String
Dim strCodMonedaDetraccion      As String, strCodCuentaFondoUnico       As String
Dim strCodCuentaFondoDetraccion As String, strCodAfectacion             As String
Dim strCodCreditoFiscal         As String, strCodFormaPagoUnico         As String
Dim strCodFormaPagoDetraccion   As String, strCodFileUnico              As String
Dim strCodAnaliticaUnico        As String, strCodBancoUnico             As String
Dim strCodCuentaUnico           As String, strCodFileDetraccion         As String
Dim strCodAnaliticaDetraccion   As String, strCodBancoDetraccion        As String
Dim strCodCuentaDetraccion      As String, strCodIngreso                  As String
Dim strIndDetraccion            As String, strCodAnalitica              As String
Dim strCodDetalleGasto          As String, strDetraccionSiNo            As String
Dim strIndImpuesto              As String, strIndRetencion              As String
Dim strCodValorTipoCambio       As String, strCodTipoGasto              As String
Dim strCodFile                  As String, strCodAplicacionDevengo      As String
Dim strEstado                   As String, strSQL                       As String
Dim strCodEstado                As String
Dim adoRegistro     As ADODB.Recordset
Dim adoRegistroAux              As ADODB.Recordset

Dim arrConcepto()       As String
Dim strCodConcepto      As String
Dim strCodIngresoLista  As String


Private Sub Calculos()
'Dim X As New clsNumSpanishWord
Dim intRegistro As Integer

    If Trim(txtSubTotal.Text) = Valor_Caracter Or Trim(txtIgv.Text) = Valor_Caracter Or Trim(txtTotal.Text) = Valor_Caracter Then Exit Sub

    lblIGV.Caption = "IGV"

    If strCodAfectacion = Codigo_Afecto Then
        If strIndImpuesto = Valor_Indicador Then
            txtSubTotal.Text = lblMontoIngreso.Caption
            txtIgv.Text = CStr(CCur(txtSubTotal.Text) * gdblTasaIgv)
            txtTotal.Text = CStr(CCur(txtSubTotal.Text) + CCur(txtIgv.Text))
            
'            lblIGV.Caption = "IGV"
            lblTasaIGV = gdblTasaIgv * 100
'            txtTotal.Text = lblMontoIngreso.Caption
'            txtIgv.Text = CStr((CCur(txtTotal.Text) * gdblTasaIgv) / (1 + gdblTasaIgv))
'            txtSubTotal.Text = CStr(CCur(txtTotal.Text) - CCur(txtIgv.Text))
        Else
            txtSubTotal.Text = lblMontoIngreso.Caption
            txtIgv.Text = "0"
            lblTasaIGV = "0"
            txtTotal.Text = txtSubTotal.Text
        End If
    Else
'        If strIndImpuesto = Valor_Indicador Then
            txtTotal.Text = lblMontoIngreso.Caption
            txtSubTotal.Text = txtTotal.Text
'        ElseIf strIndRetencion = Valor_Indicador Then
'            txtSubTotal.Text = lblMontoIngreso.Caption
'            txtTotal.Text = txtSubTotal.Text
'        Else
'            txtSubTotal.Text = lblMontoIngreso.Caption
'            txtTotal.Text = txtSubTotal.Text
'        End If
        txtIgv.Text = "0"
        lblTasaIGV = "0"
    End If

    'lblTotalLetras.Caption = X.ConvertCurrencyToSpanish(CDec(txtTotal.Text), cboMoneda.Text)
       
End Sub


Private Sub Deshabilita()

    strIndDetraccion = Valor_Caracter
    
    Call Calculos
    
End Sub

Private Sub Habilita()

    strIndDetraccion = Valor_Indicador
    
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
            
            Call CargarOrdenesCobro
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
End Sub

Private Sub CargarOrdenesCobro()
'*** Ordenes de pago del Fondo ***
    strSQL = "SELECT oc.NumOrdenCobro CODIGO, (RTRIM(fi.DescripIngreso)) DESCRIP " & _
            "FROM OrdenCobro oc INNER JOIN FondoIngreso fi ON oc.CodFondo = fi.CodFondo AND oc.CodAdministradora = fi.CodAdministradora AND oc.NumIngreso = fi.NumIngreso " & _
            "WHERE oc.CodFondo='" & strCodFondo & "' " & _
            "AND oc.CodAdministradora='" & gstrCodAdministradora & "' " & _
            "AND fi.CodContratante = '" & lblCodContratante.Caption & "' " & _
            "AND oc.CodMoneda = '" & strCodMoneda & "' " & _
            "AND oc.Estado = '01' " & _
            "AND oc.NumOrdenCobro NOT IN (" & strCodIngresoLista & ")" '(SELECT RTRIM(LTRIM(item)) FROM dbo.fnSplit('" & strCodGastoLista & "',','))"

    If strCodTipoComprobante <> "07" Then  ' NO ES nota de crédito
        strSQL = strSQL & " AND MontoIngreso > 0"
    Else
        'SI es nota de credito
        strSQL = strSQL & " AND MontoIngreso < 0"
    End If
CargarControlLista strSQL, cboCobro, arrCobro(), Sel_Defecto
End Sub

Public Sub Accion(nAccion As ButtonAction)
    
    Select Case nAccion
        
        Case vNew
            Call Adicionar
        Case vQuery
            Call Modificar
        Case vDelete
            Call Anular   'Eliminar
        Case vSearch
            Call Buscar
        Case vReport
            Call Anular
        Case vSave
            Call Grabar
        Case vCancel
            Call Cancelar
        Case vPrint
            Call Imprimir
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
    Call Buscar
    
End Sub

Public Sub Contabilizar()
'Dim strEstadoRegCompra As String
'Dim strMsgError As String
'
'On Error GoTo err
'
''Validamos si el registro de compra ya fue enviado a comtabilidad
'If strEstado = Reg_Edicion Then
'    strEstadoRegCompra = traerCampo("RegistroCompra", "Estado", "NumRegistro", gLista.Columns.ColumnByFieldName("NumRegistro").Value, " CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' ")
'
'    If strEstadoRegCompra = "04" Then
'        strMsgError = "El Registro de Compras ya fue enviado a Contabilidad"
'        GoTo err
'    End If
'
'    If MsgBox("¿Seguro de contabilizar el Registro de Compras?", vbQuestion + vbYesNo, App.Title) = vbYes Then
'        '*** Generar Orden si no está generada o actualizar ***
'        Call ContabilizarRegistroCompra(CInt(gLista.Columns.ColumnByFieldName("NumRegistro").Value), strCodFondo, Trim(lblCodContratante.Caption), strMsgError)
'        If strMsgError <> "" Then GoTo err
'    End If
'
'    MsgBox "Registro de Compras contabilizado con exito", vbInformation, App.Title
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
    
    Dim xmlDocIngresos As DOMDocument60 'JCB
    Dim strNumDocumentoFisico As String
    Dim strMsgError As String 'JCB
    
    'MEVH --06/03/2012
    'Dim objMontoTexto As New TNumToWords.clsNumSpanishWord
    Dim montoEnTexto As String
    Dim montoNumero As Double
    
    montoNumero = CDbl(txtTotal.Text)
'    If strCodMoneda = "02" Then
'        montoEnTexto = objMontoTexto.ConvertCurrencyToSpanish(montoNumero, "") & " Dolares Americanos. "
'    Else
'        montoEnTexto = objMontoTexto.ConvertCurrencyToSpanish(montoNumero, "") & " Nuevos Soles. "
'    End If
'
'
    'FIN - MEVH --06/03/2012
    
    Dim strNumDocumento As String
    Dim strFechaDesembolso As Date
    Dim strIndFactura As String, strIndBoleta As String
    Dim strdiaCancelacion As Integer, stranioCancelación As Integer
    Dim strmesCancelacion As String
    'FIN - MEVH --04/05/2012
    
    
    If strEstado = Reg_Consulta Then Exit Sub
    
    If Not TodoOK() Then Exit Sub
    
    XMLDetalleGrid xmlDocIngresos, "DetIngresos", gIngresos, "Item,DescripIngreso,NumOrdenCobro,CodMoneda,MontoIngreso,CodFile,CodDetalleFile,CodAnalitica", strMsgError 'JCB
'''    If strMsgError <> "" Then GoTo CtrlError 'JCB
    
    strCodFile = "000"
    strCodAnalitica = "0000000000"
        
        Me.MousePointer = vbHourglass
        
        strNuevoMod = "I"
        If strEstado = Reg_Edicion Then strNuevoMod = "U"
        
        '*** Guardar ***
        With adoComm
            
            If strCodTipoComprobante = "07" Then
                ''JAFR 03/05/12: Si es Nota de crédito, obedece a una operación de prepago/quiebre o cancelación
                ''***Obteniendo datos de la operacion original:
                .CommandText = "Select NumIngreso from OrdenCobro where CodFondo = '" & strCodFondo & "' and NumOrdenCobro = '" & _
                                gIngresos.Columns.ColumnByFieldName("NumOrdenCobro").Value & "'"
                Set adoAuxiliar = .Execute
    
                .CommandText = "Select NumOperacion from FondoIngreso where CodFondo = '" & strCodFondo & "' and NumIngreso = " & _
                                adoAuxiliar("NumIngreso")
                Set adoAuxiliar = .Execute
    
                'si son intereses adelantados los que se devuelve
                'inicio comentarios ACR: 16/05/13
                .CommandText = "Select CodFile, CodAnalitica from InversionOperacion where CodFondo = '" & strCodFondo & "' and NumOperacion = '" & _
                                adoAuxiliar("NumOperacion") & "'"
                Set adoAuxiliar = .Execute
    
'                .CommandText = "Select NumIngreso from FondoIngreso where CodFondo = '" & strCodFondo & "' and CodFile = '" & _
'                                adoAuxiliar("CodFile") & "' AND CodAnalitica = '" & adoAuxiliar("CodAnalitica") & "' --AND CodCuenta LIKE '70411[1-5]%'"
'                Set adoAuxiliar = .Execute
'                'fin comentarios ACR: 16/05/13
'
''                .CommandText = "Select NumIngreso from FondoIngreso where CodFondo = '" & strCodFondo & "' and NumOperacion = '" & _
''                                adoAuxiliar("NumOperacion") & "'" ' AND CodCuenta LIKE '70411[1-5]%'"
''                Set adoAuxiliar = .Execute
'
'
'                .CommandText = "Select NumOrdenCobro from OrdenCobro where CodFondo = '" & strCodFondo & "' and NumIngreso = " & _
'                               adoAuxiliar("NumIngreso")
'                Set adoAuxiliar = .Execute
'
'                .CommandText = "Select NumRegistro, CodFile, CodDetalleFile, CodAnalitica from RegistroVentaDetalle where CodFondo = '" & strCodFondo & "' and NumOrdenCobro = '" & _
'                               adoAuxiliar("NumOrdenCobro") & "'"
'                Set adoAuxiliar = .Execute
       
                strCodFile = adoAuxiliar("CodFile")
                'strCodDetalleFile = adoAuxiliar("CodDetalleFile")
                strCodAnalitica = adoAuxiliar("CodAnalitica")
                
'                .CommandText = "Select SerieComprobante, NumComprobante, FechaComprobante from RegistroVenta where CodFondo = '" & strCodFondo & "' and NumRegistro = " & _
                               adoAuxiliar("NumRegistro")
'                Set adoAuxiliar = .Execute
                
'                strNumDocumento = Trim(adoAuxiliar("SerieComprobante") + "-" + adoAuxiliar("NumComprobante"))

                'strNumDocumento = Trim(adoAuxiliar("NumComprobante"))
'                strFechaDesembolso = adoAuxiliar("FechaComprobante")
                
                If strCodFile = "014" Then
                    strIndFactura = "X"
                    strIndBoleta = ""
                Else
                    strIndFactura = ""
                    strIndBoleta = "X"
                End If
                
                strdiaCancelacion = Format(dtpFechaComprobante.Value, "dd")
                stranioCancelación = Format(dtpFechaComprobante.Value, "yy")
                strmesCancelacion = Format(dtpFechaComprobante.Value, "MMMM")

            End If
        
            
            .CommandText = "{ call up_CNManRegistroVenta('" & _
                strCodFondo & "','" & gstrCodAdministradora & "','" & lblNumSecuencial.Caption & "','" & Convertyyyymmdd(dtpFechaRegistro.Value) & "','','" & _
                strCodTipoComprobante & "','" & Convertyyyymmdd(dtpFechaComprobante.Value) & "','" & Trim(txtSerieComprobante.Text) & "','" & _
                Trim(txtNumComprobante.Text) & "','" & strCodConcepto & "','" & Trim(lblCodContratante.Caption) & "','" & Codigo_Tipo_Persona_Emisor & _
                "','" & Trim(lblCodContratante.Caption) & "','" & Trim(lblContratante.Caption) & "','" & Trim(lblNumDocID.Caption) & "','" & _
                Trim(lblDireccion.Caption) & "','" & txtDescripcion.Text & "','" & strCodAfectacion & "','" & strCodCreditoFiscal & "','" & _
                Trim(txtPeriodoFiscal.Text) & "','" & strCodMoneda & "'," & CDec(txtSubTotal.Text) & "," & CDec(lblTasaIGV) & "," & _
                CDec(txtIgv.Text) & "," & CDec(txtTotal.Text) & ",'','" & strCodFile & "','" & _
                strCodAnalitica & "','" & Convertyyyymmdd(Valor_Fecha) & "',0,'" & Estado_Activo & "','','','" & CrearXMLDetalle(xmlDocIngresos) & _
                "','" & strNuevoMod & "') }"

                
'            If strCodTipoComprobante = "07" Then
'                .CommandText = .CommandText & ",'" & strNumDocumento & "','" & Convertyyyymmdd(strFechaDesembolso) & "','" & strIndFactura & "','" & strIndBoleta & "'," & _
'                                strdiaCancelacion & ",'" & strmesCancelacion & "'," & stranioCancelación & ") }"
'            Else
'                .CommandText = .CommandText & ",'','','','',0,'',0) }"
'            End If
            
            adoConn.Execute .CommandText
        
        End With
                                    
        Me.MousePointer = vbDefault
                    
        If strNuevoMod = "I" Then
            MsgBox Mensaje_Adicion_Exitosa, vbExclamation
        Else
            MsgBox Mensaje_Edicion_Exitosa, vbExclamation
        End If
        
        frmMainMdi.stbMdi.Panels(3).Text = "Acción"
        
        cmdOpcion.Visible = True
        With tabRegistroCompras
            .TabEnabled(0) = True
            .Tab = 0
        End With

        Call Buscar
        
        Call CargarOrdenesCobro
    
End Sub

Private Function TodoOK() As Boolean
        
    TodoOK = False
            
    If cboTipoComprobante.ListIndex <= 0 Then
        MsgBox "Seleccione el tipo de comprobante", vbCritical, Me.Caption
        If cboTipoComprobante.Enabled Then cboTipoComprobante.SetFocus
        Exit Function
    End If
        
    If Trim(txtSerieComprobante.Text) = Valor_Caracter Then
        MsgBox "Ingrese el número de serie", vbCritical, Me.Caption
        If txtSerieComprobante.Enabled Then txtSerieComprobante.SetFocus
        Exit Function
    End If
    
    If Trim(txtNumComprobante.Text) = Valor_Caracter Then
        MsgBox "Ingrese el número de comprobante", vbCritical, Me.Caption
        If txtNumComprobante.Enabled Then txtNumComprobante.SetFocus
        Exit Function
    End If
    
    If Trim(lblContratante.Caption) = Valor_Caracter Then
        MsgBox "Seleccione el Contratante", vbCritical, Me.Caption
        If cmdContratante.Enabled Then cmdContratante.SetFocus
        Exit Function
    End If
    
    If gIngresos.Count = 0 Or (gIngresos.Count = 1 And Trim(gIngresos.Columns.ColumnByFieldName("NumOrdenCobro").Value) = "") Then
        MsgBox "Debe seleccionar las Ordenes de Cobro que conformarán el detalle del Documento.", vbCritical, Me.Caption
        If gIngresos.Enabled Then gIngresos.SetFocus
        Exit Function
    End If
    
    '*** Si todo paso OK ***
    TodoOK = True
  
End Function

Public Sub Imprimir()
Dim strMsgError As String
Dim adoRegistro As ADODB.Recordset
Dim strIndTotalizado As String

On Error GoTo err

    If MsgBox("¿Desea Imprimir el(la) " & Trim(cboTipoComprobante.List(cboTipoComprobante.ListIndex)) & " en forma Resumida?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        strIndTotalizado = Valor_Indicador
    Else
        strIndTotalizado = Valor_Caracter
    End If
    
    Call ImprimeComprobanteCobro(strCodFondo, lblNumSecuencial.Caption, strCodTipoComprobante, txtNumComprobante.Text, txtSerieComprobante.Text, strMsgError, strIndTotalizado)
    
    If strMsgError <> "" Then
        GoTo err
    Else
        'Actualizar el indicador de "Impresión" del documento
        Set adoRegistro = New ADODB.Recordset
        With adoComm
            .CommandText = "UPDATE RegistroVenta SET IndImpresion = 'X' " & _
                "WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND NumRegistro = " & gLista.Columns.ColumnByFieldName("NumRegistro").Value
            Set adoRegistro = .Execute
  
            Set adoRegistro = Nothing
        
        End With
        
        MsgBox "Se realizó la impresión del Documento.", vbInformation, Me.Caption
        
        frmMainMdi.stbMdi.Panels(3).Text = "Acción"
        
        cmdOpcion.Visible = True
        With tabRegistroCompras
            .TabEnabled(0) = True
            .Tab = 0
        End With

        Call Buscar
        
        Call CargarOrdenesCobro

    End If

Exit Sub
err:
If strMsgError = "" Then strMsgError = err.Description
MsgBox strMsgError, vbInformation, App.Title
End Sub

Public Sub Eliminar()

End Sub

Public Sub Modificar()

    If strEstado = Reg_Consulta Then
    
        'Validar si fue anulado el documento de cobro.
        If strCodEstado = "03" Then
            MsgBox "Documento de cobro está anulado. No es posible la modificación.", vbInformation, Me.Caption
        'Validar si ya fue impreso el documento de cobro. De ser así no permitir la modificación
        ElseIf DocumentoFueImpreso(gLista.Columns.ColumnByFieldName("NumRegistro").Value) = True Then
            BloquearControles (False)
            strEstado = Reg_Edicion
            LlenarFormulario strEstado
            cmdOpcion.Visible = False
            With tabRegistroCompras
                .TabEnabled(0) = False
                .Tab = 1
            End With
                        
            'MsgBox "Documento de cobro ya fue impreso. No es posible la modificación.", vbInformation, Me.Caption
        Else
            BloquearControles (True)
            strEstado = Reg_Edicion
            LlenarFormulario strEstado
            cmdOpcion.Visible = False
            With tabRegistroCompras
                .TabEnabled(0) = False
                .Tab = 1
            End With
            
        End If
        
    End If
    
End Sub

Private Function DocumentoFueImpreso(intNumRegistro As Integer) As Boolean

    Dim adoRegistro As ADODB.Recordset
    
    DocumentoFueImpreso = False
    
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        .CommandText = "SELECT IndImpresion FROM RegistroVenta " & _
            "WHERE NumRegistro=" & intNumRegistro & " AND CodFondo='" & _
            strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
        Set adoRegistro = .Execute
    
        If Not adoRegistro.EOF Then
            If Trim(adoRegistro("IndImpresion")) = "X" Then
                DocumentoFueImpreso = True
            End If
        End If
        
        adoRegistro.Close: Set adoRegistro = Nothing
        
   End With

End Function
Private Sub LlenarFormulario(strModo As String)

    Dim adoRegistro     As ADODB.Recordset, intRegistro       As Integer
    Dim adoAuxiliar     As ADODB.Recordset
    Dim strMsgError     As String
    
    Select Case strModo
        Case Reg_Adicion
            fraCompras(1).Caption = "Definición del Registro - Fondo : " & Trim(cboFondo.Text)
            
            Set adoRegistro = New ADODB.Recordset
            With adoComm
            
                .CommandText = "SELECT MAX(NumRegistro) NumRegistro FROM RegistroVenta " & _
                    "WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
                Set adoRegistro = .Execute
                
                If Not adoRegistro.EOF Then
                    If IsNull(adoRegistro("NumRegistro")) Then
                        lblNumSecuencial.Caption = "1"
                    Else
                        lblNumSecuencial.Caption = CStr(adoRegistro("NumRegistro") + 1)
                    End If
                Else
                    lblNumSecuencial.Caption = "1"
                End If
                adoRegistro.Close
                
                strCodIngresoLista = "''"
                
                dtpFechaRegistro.Value = gdatFechaActual
                If cboTipoComprobante.ListCount > 0 Then cboTipoComprobante.ListIndex = 0
                dtpFechaComprobante.Value = gdatFechaActual
                txtSerieComprobante.Text = Valor_Caracter
                txtNumComprobante.Text = Valor_Caracter
                If cboCobro.ListCount > 0 Then cboCobro.ListIndex = 0
                lblCodContratante.Caption = Valor_Caracter
                lblContratante.Caption = Valor_Caracter
                lblDireccion.Caption = Valor_Caracter
                lblTipoDocID.Caption = Valor_Caracter
                lblNumDocID.Caption = Valor_Caracter
                txtDescripcion.Text = Valor_Caracter
                
                If cboAfectacion.ListCount > 0 Then cboAfectacion.ListIndex = 0
                
                intRegistro = ObtenerItemLista(arrAfectacion(), Codigo_Afecto)
                If intRegistro >= 0 Then cboAfectacion.ListIndex = intRegistro
                
                If cboCreditoFiscal.ListCount > 0 Then cboCreditoFiscal.ListIndex = 0
                
                If cboIngreso.ListCount > 0 Then cboIngreso.ListIndex = 0
                                        
                txtPeriodoFiscal.Text = Valor_Caracter
                txtSubTotal.Text = "0": txtIgv.Text = "0"
                txtTotal.Text = "0"

                strCodFile = "000"
                
                Me.Refresh
     
            End With
                        
            cboTipoComprobante.SetFocus
        
        Case Reg_Edicion
            
            Set adoRegistro = New ADODB.Recordset

            With adoComm
                .CommandText = "SELECT * FROM RegistroVenta " & _
                    "WHERE NumRegistro=" & gLista.Columns.ColumnByFieldName("NumRegistro").Value & " AND CodFondo='" & _
                    strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
                Set adoRegistro = .Execute

                If Not adoRegistro.EOF Then
                    fraCompras(1).Caption = "Definición del Registro - Fondo : " & Trim(cboFondo.Text)

                    lblNumSecuencial.Caption = CStr(adoRegistro("NumRegistro"))
                    dtpFechaRegistro.Value = adoRegistro("FechaRegistro")

                    intRegistro = ObtenerItemLista(arrTipoComprobante(), adoRegistro("CodTipoComprobante"))
                    If intRegistro >= 0 Then cboTipoComprobante.ListIndex = intRegistro

                    dtpFechaComprobante.Value = adoRegistro("FechaComprobante")
                    
                    txtSerieComprobante.Text = adoRegistro("SerieComprobante")
                    txtNumComprobante.Text = adoRegistro("NumComprobante")

                    lblContratante.Caption = Valor_Caracter
                    lblDireccion.Caption = Valor_Caracter
                    lblCodContratante.Caption = adoRegistro("CodContratante")

                    Set adoAuxiliar = New ADODB.Recordset
                    .CommandText = "SELECT IP.NumIdentidad, IP.DescripPersona, IP.Direccion1 + IP.Direccion2 Direccion, AP.DescripParametro TipoIdentidad " & _
                        "FROM InstitucionPersona IP " & _
                        "JOIN AuxiliarParametro AP ON (AP.CodParametro = IP.TipoIdentidad AND AP.CodTipoParametro = 'TIPIDE')" & _
                        "WHERE CodPersona='" & lblCodContratante.Caption & "' AND TipoPersona='" & Codigo_Tipo_Persona_Emisor & "'"
                    Set adoAuxiliar = .Execute

                    If Not adoAuxiliar.EOF Then
                        lblTipoDocID.Caption = Trim(adoAuxiliar("TipoIdentidad"))
                        lblContratante.Caption = Trim(adoAuxiliar("DescripPersona"))
                        lblNumDocID.Caption = Trim(adoAuxiliar("NumIdentidad"))
                        lblDireccion.Caption = Trim(adoAuxiliar("Direccion"))
                    End If
                    adoAuxiliar.Close: Set adoAuxiliar = Nothing
                    
                    intRegistro = ObtenerItemLista(arrConcepto(), adoRegistro("CodCuenta"))
                    If intRegistro >= 0 Then cboIngreso.ListIndex = intRegistro

                    txtDescripcion.Text = Trim(adoRegistro("DescripRegistro"))

                    intRegistro = ObtenerItemLista(arrAfectacion(), adoRegistro("CodAfectacion"))
                    If intRegistro >= 0 Then cboAfectacion.ListIndex = intRegistro

                    intRegistro = ObtenerItemLista(arrCreditoFiscal(), adoRegistro("CodCreditoFiscal"))
                    If intRegistro >= 0 Then cboCreditoFiscal.ListIndex = intRegistro
                    
                    intRegistro = ObtenerItemLista(arrMoneda(), adoRegistro("CodMoneda"))
                    If intRegistro >= 0 Then cboMoneda.ListIndex = intRegistro

                    txtPeriodoFiscal.Text = adoRegistro("DescripPeriodoCredito")
                    txtSubTotal.Text = CStr(adoRegistro("Importe"))
                    lblTasaIGV.Caption = CStr(adoRegistro("PorcenIgv"))
                    txtIgv.Text = CStr(adoRegistro("ValorImpuesto"))
                    txtTotal.Text = CStr(adoRegistro("ValorTotal"))
                    
                    strCodFile = adoRegistro("CodFile") 'Trim(tdgPendientes.Columns(9).Value)
                    
                    If cboTipoComprobante.Enabled = True Then
                        cboTipoComprobante.SetFocus
                    End If
                    
                    'Muestro el detalle
                    .CommandText = "SELECT SecRegistroDetalle AS Item, NumOrdenCobro, CodFile, CodDetalleFile, CodAnalitica,DescripRegistroDetalle AS DescripIngreso, " & _
                                          "CodMoneda, MontoTotal AS MontoIngreso " & _
                                   "FROM RegistroVentaDetalle " & _
                                   "WHERE NumRegistro=" & gLista.Columns.ColumnByFieldName("NumRegistro").Value & " " & _
                                     "AND CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
                                     
                    Set adoRegistro = .Execute
                    
                    FormatoGrillaIngresos strMsgError
                    mostrarDatosGridRS gIngresos, adoRegistro, strMsgError
                                    
                End If
                adoRegistro.Close: Set adoRegistro = Nothing
            End With
    End Select
    
End Sub

Public Sub Adicionar()
Dim strMsgError As String

On Error GoTo err
    
    BloquearControles (True)
    
    frmMainMdi.stbMdi.Panels(3).Text = "Adicionar Registro..."
    
    FormatoGrillaIngresos strMsgError
    If strMsgError <> "" Then GoTo err
    
    strEstado = Reg_Adicion
    LlenarFormulario strEstado
    cmdOpcion.Visible = False
    With tabRegistroCompras
        .TabEnabled(0) = False
        .Tab = 1
    End With
        
Exit Sub
err:
If strMsgError = "" Then strMsgError = err.Description
MsgBox strMsgError, vbInformation, App.Title
End Sub

Private Sub cboCobro_Click()

    Dim adoRegistro         As ADODB.Recordset
    Dim curDiferencia       As Currency
    Dim intRegistro         As Integer
        
    strCodIngreso = Valor_Caracter: strCodAnalitica = Valor_Caracter
    strCodDetalleGasto = Valor_Caracter
    If cboCobro.ListIndex < 0 Then Exit Sub
    
    strCodIngreso = Trim(Left(arrCobro(cboCobro.ListIndex), 10))
    strCodAnalitica = Right(arrCobro(cboCobro.ListIndex), 8)
    
'    If strCodTipoComprobante <> "07" Then 'nota de credito
'        txtDescripcion.Text = GenerarGlosaComprobante()
'    End If
    
    
'''    lblAnalitica.Caption = Trim(tdgPendientes.Columns(9).Value) & " - " & strCodAnalitica
        
'    Set adoRegistro = New ADODB.Recordset
'
'    With adoComm

'        txtDescripcion.Text = Valor_Caracter
'        lblMontoIngreso.Caption = "0"
'        txtSubTotal.ToolTipText = Valor_Caracter
'        txtSubTotal.Text = "0"
        
'        .CommandText = "SELECT MontoGasto,MontoDevengo,DescripGasto,FechaFinal,CodCreditoFiscal,CodMoneda,CodTipoGasto " & _
'            "FROM FondoGasto " & _
'            "WHERE NumGasto=" & CInt(tdgPendientes.Columns(2).Value) & " AND CodCuenta='" & strCodIngreso & "' AND " & _
'            "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
'        Set adoRegistro = .Execute
'
'        If Not adoRegistro.EOF Then
        
'            txtDescripcion.Text = Trim(adoRegistro("DescripGasto"))
'            lblMontoIngreso.Caption = CStr(adoRegistro("MontoGasto"))
            
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
'                "WHERE CodFile='" & Trim(tdgPendientes.Columns(9).Value) & "' AND DescripDetalleFile='" & strCodIngreso & "'"
'        Else
'            '.CommandText = "SELECT CodDetalleFile FROM InversionDetalleFile " & _
'            '    "WHERE CodFile='" & Trim(tdgPendientes.Columns(8).Value) & "' AND CodDetalleFile='" & strCodIngreso & "'"
'            .CommandText = "SELECT CodDetalleFile FROM DinamicaContable " & _
'                "WHERE CodFile='" & Trim(tdgPendientes.Columns(9).Value) & "' AND CodCuenta='" & strCodIngreso & "'"
'        End If
'        Set adoRegistro = .Execute
'
'        If Not adoRegistro.EOF Then
'            strCodDetalleGasto = adoRegistro("CodDetalleFile")
'        End If
'        adoRegistro.Close: Set adoRegistro = Nothing
'    End With
    
End Sub

Private Sub cboIngreso_Click()
    strCodConcepto = Valor_Caracter
    If cboIngreso.ListIndex <= 0 Then Exit Sub
    
    strCodConcepto = Trim(arrConcepto(cboIngreso.ListIndex))
    txtDescripcion.Text = GenerarGlosaComprobante
    
    
End Sub

Private Sub cboMoneda_Click()

    strCodMoneda = Valor_Caracter
    If cboMoneda.ListIndex < 0 Then Exit Sub
    
    strCodMoneda = arrMoneda(cboMoneda.ListIndex)
    
    lblSignoMonedaVV.Caption = ObtenerCodSignoMoneda(strCodMoneda)
    lblSignoMonedaIGV.Caption = lblSignoMonedaVV.Caption
    lblSignoMonedaPV.Caption = lblSignoMonedaVV.Caption
    
    Call CargarOrdenesCobro
    
End Sub

Private Sub cboTipoComprobante_Click()

    Dim adoRegistro     As ADODB.Recordset
    
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
            lblVV.Caption = Trim(adoRegistro("DescripCampo1"))
            lblIGV.Caption = Trim(adoRegistro("DescripCampo2"))
            lblPV.Caption = Trim(adoRegistro("DescripCampo3"))
            
'''            strCtaImpuesto = ObtenerCuentaAdministracion("025", "R")
'''            If strIndRetencion = Valor_Indicador Then strCtaImpuesto = ObtenerCuentaAdministracion("036", "R")
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    'Call cboCobro_Click
    
'    If strCodTipoComprobante = "07" Then 'nota de credito
'
'        strSQL = "SELECT FCI.CodCuenta CODIGO, " & _
'                "RTRIM(DescripCuenta) + ' (DEVOLUCION DE INTERESES)' AS DESCRIP " & _
'                "FROM FondoConceptoIngreso FCI JOIN PlanContable PCG ON(PCG.CodCuenta=FCI.CodCuenta AND PCG.CodAdministradora=FCI.CodAdministradora) " & _
'                "WHERE CodFondo='" & strCodFondo & "' AND FCI.CodAdministradora='" & gstrCodAdministradora & "' AND FCI.CodCuenta LIKE '496%' " & _
'                "ORDER BY DescripCuenta"
'        CargarControlLista strSQL, cboIngreso, arrConcepto(), Sel_Defecto
'
'        cboIngreso.Enabled = True 'False
'        GenerarGlosaComprobante
'    ElseIf strCodTipoComprobante = "08" Then
'
'       strSQL = "SELECT FCI.CodCuenta CODIGO,(RTRIM(DescripCuenta)) DESCRIP " & _
'                "FROM FondoConceptoIngreso FCI JOIN PlanContable PCG ON(PCG.CodCuenta=FCI.CodCuenta AND PCG.CodAdministradora=FCI.CodAdministradora) " & _
'                "WHERE CodFondo='" & strCodFondo & "' AND FCI.CodAdministradora='" & gstrCodAdministradora & "' AND FCI.CodCuenta LIKE '759%' " & _
'                "ORDER BY DescripCuenta"
'        CargarControlLista strSQL, cboIngreso, arrConcepto(), Sel_Defecto
'        cboIngreso.Enabled = True
'    Else
        
        strSQL = "SELECT FCI.CodCuenta CODIGO,(RTRIM(DescripCuenta)) DESCRIP " & _
                "FROM FondoConceptoIngreso FCI JOIN PlanContable PCG ON(PCG.CodCuenta=FCI.CodCuenta AND PCG.CodAdministradora=FCI.CodAdministradora) " & _
                "WHERE CodFondo='" & strCodFondo & "' AND FCI.CodAdministradora='" & gstrCodAdministradora & "' AND FCI.CodCuenta LIKE '704%' " & _
                "ORDER BY DescripCuenta"
        CargarControlLista strSQL, cboIngreso, arrConcepto(), Sel_Defecto
        
        cboIngreso.Enabled = True
'    End If
    
    'Call cboDetraccion_Click
    Call CargarOrdenesCobro
    Call Calculos
   
End Sub

Private Sub cmdAdicionarCobro_Click()
Dim strMsgError As String

On Error GoTo err

    If cboCobro.ListIndex <= 0 Then
        strMsgError = "Debe seleccionar un Cobro."
        GoTo err
    End If
    
    If gIngresos.Columns.ColumnByFieldName("DescripIngreso").Value <> "" Or gIngresos.Count = 0 Then
        gIngresos.Dataset.Insert
    End If
    
    gIngresos.Dataset.Edit
    
    gIngresos.Columns.ColumnByFieldName("item").Value = gIngresos.Count
    gIngresos.Columns.ColumnByFieldName("NumOrdenCobro").Value = strCodIngreso
    gIngresos.Columns.ColumnByFieldName("CodFile").Value = ""
    gIngresos.Columns.ColumnByFieldName("CodAnalitica").Value = ""
    gIngresos.Columns.ColumnByFieldName("DescripIngreso").Value = cboCobro.Text
    gIngresos.Columns.ColumnByFieldName("CodMoneda").Value = ""
    gIngresos.Columns.ColumnByFieldName("MontoIngreso").Value = 0
    gIngresos.Columns.ColumnByFieldName("CodDetalleFile").Value = ""
    
    Set adoRegistro = New ADODB.Recordset

    With adoComm

        'el CodCreditoFiscal lo jalamos de la tabla fondo ingreso o del form JCB?
        .CommandText = "SELECT op.MontoOrdenCobro, op.CodMoneda, fi.CodFile, fi.CodCuenta, fi.CodAnalitica " & _
            "FROM OrdenCobro op INNER JOIN FondoIngreso fi ON op.CodFondo = fi.CodFondo AND op.CodAdministradora = fi.CodAdministradora AND op.NumIngreso = fi.NumIngreso " & _
            "WHERE op.NumOrdenCobro = " & strCodIngreso & " " & _
              "AND op.CodFondo='" & strCodFondo & "' AND op.CodAdministradora='" & gstrCodAdministradora & "'"
        Set adoRegistro = .Execute

        If Not adoRegistro.EOF Then
            If strCodTipoComprobante <> "07" Then
                gIngresos.Columns.ColumnByFieldName("MontoIngreso").Value = adoRegistro("MontoOrdenCobro")
            Else
                gIngresos.Columns.ColumnByFieldName("MontoIngreso").Value = adoRegistro("MontoOrdenCobro") * (-1)
            End If
            
            If Trim(strCodIngresoLista) = "''" Then
                strCodIngresoLista = "'" & strCodIngreso & "'" 'adoRegistro("NumGasto")
            Else
                strCodIngresoLista = strCodIngresoLista & ",'" & strCodIngreso & "'" 'adoRegistro("NumGasto")
            End If
            
            gIngresos.Columns.ColumnByFieldName("CodMoneda").Value = adoRegistro("CodMoneda")
            gIngresos.Columns.ColumnByFieldName("CodAnalitica").Value = adoRegistro("CodAnalitica")
            gIngresos.Columns.ColumnByFieldName("CodFile").Value = Trim(adoRegistro("CodFile"))
        End If
        
        'Como sacamos el codido detalle file? JCB?
        If Trim(adoRegistro("CodFile")) = "090" Then
            .CommandText = "SELECT CodDetalleFile FROM InversionDetalleFile " & _
                "WHERE CodFile='" & Trim(adoRegistro("CodFile")) & "' AND DescripDetalleFile='" & adoRegistro("CodCuenta") & "'"
        Else
            .CommandText = "SELECT CodDetalleFile FROM DinamicaContable " & _
                "WHERE CodFile='" & Trim(adoRegistro("CodFile")) & "' AND CodCuenta='" & adoRegistro("CodCuenta") & "'"
        End If
        
        Set adoRegistro = .Execute

        If Not adoRegistro.EOF Then
            gIngresos.Columns.ColumnByFieldName("CodDetalleFile").Value = adoRegistro("CodDetalleFile")
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With

    gIngresos.Dataset.Post
            
    lblMontoIngreso.Caption = gIngresos.Columns.ColumnByFieldName("MontoIngreso").SummaryFooterValue
    
    Call Calculos
    
    Call CargarOrdenesCobro
    
    cboCobro.ListIndex = 0
    
Exit Sub
err:
If strMsgError = "" Then strMsgError = err.Description
MsgBox strMsgError, vbInformation, App.Title
End Sub

Private Sub cmdCobro_Click()
Dim rsCobrosTemp As ADODB.Recordset
Dim strMsgError  As String
Dim indAceptar   As Boolean

On Error GoTo err

    If strCodMoneda = "" Then
        strMsgError = "Seleccione una Moneda"
        GoTo err
    End If
    
    If lblCodContratante.Caption = "" Then
        strMsgError = "Seleccione un Contratante"
        GoTo err
    End If

    'frmComprobanteCobroAyuda.mostrarForm strCodFondo, strCodMoneda, lblCodContratante.Caption, strCodConcepto, rsCobrosTemp, indAceptar
    
    If indAceptar = False Then Exit Sub
    
    'FormatoGrillaIngresos strMsgError
    If strMsgError <> "" Then GoTo err
    
  
    Do While Not rsCobrosTemp.EOF
        
        If gIngresos.Columns.ColumnByFieldName("DescripIngreso").Value <> "" Or gIngresos.Count = 0 Then
            gIngresos.Dataset.Insert
        End If
        
        gIngresos.Dataset.Edit
        
        gIngresos.Columns.ColumnByFieldName("item").Value = gIngresos.Count
        gIngresos.Columns.ColumnByFieldName("NumOrdenCobro").Value = "" & rsCobrosTemp.Fields("NumOrdenCobro")
        gIngresos.Columns.ColumnByFieldName("CodFile").Value = "" & rsCobrosTemp.Fields("CodFile")
        gIngresos.Columns.ColumnByFieldName("CodAnalitica").Value = "" & rsCobrosTemp.Fields("CodAnalitica")
        gIngresos.Columns.ColumnByFieldName("DescripIngreso").Value = "" & rsCobrosTemp.Fields("DescripIngreso")
        gIngresos.Columns.ColumnByFieldName("CodMoneda").Value = "" & rsCobrosTemp.Fields("CodMoneda")
        gIngresos.Columns.ColumnByFieldName("MontoIngreso").Value = "" & rsCobrosTemp.Fields("MontoIngreso")
        gIngresos.Columns.ColumnByFieldName("CodDetalleFile").Value = "" & rsCobrosTemp.Fields("CodDetalleFile")
        
        gIngresos.Columns.ColumnByFieldName("NumAnexo").Value = "" & rsCobrosTemp.Fields("NumAnexo")
        gIngresos.Columns.ColumnByFieldName("NumContrato").Value = "" & rsCobrosTemp.Fields("NumContrato")
        gIngresos.Columns.ColumnByFieldName("NumDocumentoFisico").Value = "" & rsCobrosTemp.Fields("NumDocumentoFisico")
        gIngresos.Columns.ColumnByFieldName("FechaDefinicion").Value = "" & rsCobrosTemp.Fields("FechaDefinicion")
    
        gIngresos.Dataset.Post
    
        rsCobrosTemp.MoveNext
    Loop
    
    lblMontoIngreso.Caption = gIngresos.Columns.ColumnByFieldName("MontoIngreso").SummaryFooterValue
    
    'ACC 03/08/2010
    'If Trim(txtDescripcion.Text) = "" Then txtDescripcion.Text = GenerarGlosaComprobante()
    txtDescripcion.Text = GenerarGlosaComprobante()
    
    Call Calculos
    
Exit Sub
err:
If strMsgError = "" Then strMsgError = err.Description
MsgBox strMsgError, vbInformation, App.Title
End Sub

Private Sub cmdEliminarCobro_Click()
Dim strMsgError As String
Dim i As Integer

On Error GoTo err

    If gIngresos.Count = 1 Then
    
    
        'Elimina de la lista de elementos seleccionados (strCodGastoLista) el elemento que se esta sacando de la grilla
        If InStr(1, strCodIngresoLista, gIngresos.Columns.ColumnByFieldName("NumOrdenCobro").Value) > 0 Then
            'Es el ultimo elemento
            If InStr(1, strCodIngresoLista, "'" & gIngresos.Columns.ColumnByFieldName("NumOrdenCobro").Value & "'") + Len(gIngresos.Columns.ColumnByFieldName("NumOrdenCobro").Value & "'") = Len(strCodIngresoLista) Then
                If Len(gIngresos.Columns.ColumnByFieldName("NumOrdenCobro").Value) = Len(strCodIngresoLista) - 2 Then 'hay solo un elemento
                    strCodIngresoLista = "''" 'Replace(strCodGastoLista, gGastos.Columns.ColumnByFieldName("NumOrdenPago").Value, Valor_Caracter)
                Else
                    strCodIngresoLista = Replace(strCodIngresoLista, ",'" & gIngresos.Columns.ColumnByFieldName("NumOrdenCobro").Value & "'", Valor_Caracter)
                End If
            Else 'no es el ultimo elemento
                strCodIngresoLista = Replace(strCodIngresoLista, "'" & gIngresos.Columns.ColumnByFieldName("NumOrdenCobro").Value & "',", Valor_Caracter)
            End If
        End If
    
        gIngresos.Dataset.Edit
        
        gIngresos.Columns.ColumnByFieldName("Item").Value = 1
        gIngresos.Columns.ColumnByFieldName("NumOrdenCobro").Value = ""
        gIngresos.Columns.ColumnByFieldName("CodFile").Value = ""
        gIngresos.Columns.ColumnByFieldName("CodAnalitica").Value = ""
        gIngresos.Columns.ColumnByFieldName("DescripIngreso").Value = ""
        gIngresos.Columns.ColumnByFieldName("CodMoneda").Value = ""
        gIngresos.Columns.ColumnByFieldName("MontoIngreso").Value = 0
        
        gIngresos.Dataset.Post
        
    Else

        'Elimina de la lista de elementos seleccionados (strCodGastoLista) el elemento que se esta sacando de la grilla
        If InStr(1, strCodIngresoLista, gIngresos.Columns.ColumnByFieldName("NumOrdenCobro").Value) > 0 Then
            'Es el ultimo elemento
            If InStr(1, strCodIngresoLista, "'" & gIngresos.Columns.ColumnByFieldName("NumOrdenCobro").Value & "'") + Len(gIngresos.Columns.ColumnByFieldName("NumOrdenCobro").Value & "'") = Len(strCodIngresoLista) Then
                If "'" & Len(gIngresos.Columns.ColumnByFieldName("NumOrdenCobro").Value) & "'" = Len(strCodIngresoLista) Then 'hay solo un elemento
                    strCodIngresoLista = Replace(strCodIngresoLista, "'" & gIngresos.Columns.ColumnByFieldName("NumOrdenCobro").Value & "'", Valor_Caracter)
                Else
                    strCodIngresoLista = Replace(strCodIngresoLista, ",'" & gIngresos.Columns.ColumnByFieldName("NumOrdenCobro").Value & "'", Valor_Caracter)
                End If
            Else 'no es el ultimo elemento
                strCodIngresoLista = Replace(strCodIngresoLista, "'" & gIngresos.Columns.ColumnByFieldName("NumOrdenCobro").Value & "',", Valor_Caracter)
            End If
        End If

        gIngresos.Dataset.Delete
                    
        gIngresos.Dataset.First
        Do While Not gIngresos.Dataset.EOF
            
            If gIngresos.Columns.ColumnByFieldName("Item").Value > 0 Then
                i = i + 1
                gIngresos.Dataset.Edit
                gIngresos.Columns.ColumnByFieldName("Item").Value = i
                gIngresos.Dataset.Post
            End If
            
            gIngresos.Dataset.Next
        Loop
        If gIngresos.Dataset.State = dsEdit Or gIngresos.Dataset.State = dsInsert Then
            gIngresos.Dataset.Post
        End If
    
    End If
    
    lblMontoIngreso.Caption = gIngresos.Columns.ColumnByFieldName("MontoGasto").SummaryFooterValue
    
    'txtDescripcion.Text = GenerarGlosaComprobante()
    
    cboCobro.ListIndex = 0

    Call Calculos
        
    Call CargarOrdenesCobro
    
Exit Sub
err:
If strMsgError = "" Then strMsgError = err.Description
MsgBox strMsgError, vbInformation, App.Title
End Sub

Private Sub cmdContratante_Click()

    'gstrFormulario = "frmComprobantePago"
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
        
        frmBus.Caption = " Relación de Contratantes"
        .sSql = "{ call up_ACSelDatos(32) }"
        
        .OutputColumns = "1,2,3,4,5,6"
        .HiddenColumns = "1,2,3,6"
        
        .BuscarTabla
        
        Screen.MousePointer = vbNormal
        frmBus.Show 1
       
        If .iParams.Count = 0 Then Exit Sub
        
        If .iParams(1).Valor <> "" Then
            lblContratante.Caption = .iParams(5).Valor
            lblTipoDocID.Caption = .iParams(3).Valor
            lblNumDocID.Caption = .iParams(4).Valor
            lblDireccion.Caption = .iParams(6).Valor
            lblCodContratante.Caption = .iParams(1).Valor
        End If
            
       
    End With
    
    Set frmBus = Nothing
    
    Call CargarOrdenesCobro
    
End Sub

Private Sub dtpFechaComprobante_Change()

    If dtpFechaComprobante.Value > gdatFechaActual Then
        MsgBox "La Fecha de comprobante debe ser igual o anterior a la fecha actual...se cambiará por la fecha actual !", vbInformation, Me.Caption
        dtpFechaComprobante.Value = gdatFechaActual
    End If

End Sub

Private Sub Form_Activate()

    Call CargarReportes
    
End Sub

Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Registro de Ventas"
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
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

    Dim c As Object

    For Each c In Me.Controls
        
        If TypeOf c Is Label Then
            Call FormatoEtiqueta(c, vbLeftJustify)
        End If
        
        If TypeOf c Is Frame Then
            Call FormatoMarco(c)
        End If
    Next

            
End Sub

Public Sub Buscar()
            
    strSQL = "SELECT NumRegistro,CodTipoComprobante,CodContratante,DescripRegistro,RV.CodMoneda,ValorTotal, " & _
        "TCP.DescripTipoComprobantePago DescripTipoComprobante, CodSigno,FechaRegistro,DescripPersona DescripContratante,RV.NumIngreso " & _
        "FROM RegistroVenta RV JOIN TipoComprobantePago TCP ON(TCP.CodTipoComprobantePago=RV.CodTipoComprobante) " & _
        "JOIN Moneda MON ON(MON.CodMoneda=RV.CodMoneda) " & _
        "LEFT JOIN InstitucionPersona IP ON(IP.CodPersona=RV.CodContratante AND IP.TipoPersona=RV.TipoContratante) " & _
        "WHERE (FechaRegistro>='" & Convertyyyymmdd(dtpFechaDesde.Value) & "' AND FechaRegistro<'" & Convertyyyymmdd(DateAdd("d", 1, dtpFechaHasta.Value)) & "') AND " & _
        "CodAdministradora='" & gstrCodAdministradora & "' AND CodFondo='" & strCodFondo & "' "
        
        
        
    If strCodEstado <> Valor_Caracter Then
        strSQL = strSQL & "AND RV.Estado='" & strCodEstado & "' ORDER BY NumRegistro"
    Else
        strSQL = strSQL & "ORDER BY NumRegistro"
    End If
        

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
            
End Sub
Private Sub CargarListas()
            
    '*** Fondos ***
    strSQL = "SELECT CodFondo CODIGO,DescripFondo DESCRIP FROM Fondo WHERE CodAdministradora='" & gstrCodAdministradora & "' AND Estado='01' and CodFondo = '" & gstrCodFondoContable & "' ORDER BY DescripFondo"
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
    
    '*** Estados del Registro ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP From AuxiliarParametro WHERE CodTipoParametro='ESTREG' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboEstado, arrEstado(), Sel_Todos
    If cboEstado.ListCount >= 0 Then cboEstado.ListIndex = 1

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
    
End Sub

Private Sub InicializarValores()

    strEstado = Reg_Defecto
    tabRegistroCompras.Tab = 0
    
    
    strCodIngresoLista = "''"
      
    dtpFechaDesde.Value = gdatFechaActual
    dtpFechaHasta.Value = gdatFechaActual
    
    ConfGrid gLista, False, False, False, False
    ConfGrid gIngresos, True, False, False, False
    
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    Set cmdOpcion.FormularioActivo = Me
    
End Sub
Private Sub Form_Unload(Cancel As Integer)

    Call OcultarReportes
    frmMainMdi.stbMdi.Panels(3).Text = "Acción"
    Set frmComprobanteCobro = Nothing
    
End Sub

Private Sub lblContratante_Change()
'*** Ordenes de pago del Fondo ***MEVH
    strSQL = "SELECT oc.NumOrdenCobro CODIGO, (RTRIM(fi.DescripIngreso)) DESCRIP " & _
             "FROM OrdenCobro oc INNER JOIN FondoIngreso fi ON oc.CodFondo = fi.CodFondo AND oc.CodAdministradora = fi.CodAdministradora AND oc.NumIngreso = fi.NumIngreso " & _
             "WHERE oc.CodFondo='" & strCodFondo & "' " & _
               "AND oc.CodAdministradora='" & gstrCodAdministradora & "' " & _
               "AND fi.CodContratante = '" & lblCodContratante.Caption & "' " & _
               "AND oc.CodMoneda = '" & strCodMoneda & "' " & _
               "AND oc.Estado = '01'"
               
    If strCodTipoComprobante <> "07" Then  ' NO ES nota de crédito
    strSQL = strSQL & " AND MontoIngreso > 0"
    Else
        'SI es nota de credito
        strSQL = strSQL & " AND MontoIngreso < 0"
    End If
    
    While gIngresos.Dataset.RecordCount > 1
        gIngresos.Dataset.Delete
    Wend
    strCodIngresoLista = "''"
    gIngresos.Dataset.Edit
    
    gIngresos.Columns.ColumnByFieldName("Item").Value = 1
    gIngresos.Columns.ColumnByFieldName("NumOrdenCobro").Value = ""
    gIngresos.Columns.ColumnByFieldName("CodFile").Value = ""
    gIngresos.Columns.ColumnByFieldName("CodAnalitica").Value = ""
    gIngresos.Columns.ColumnByFieldName("DescripIngreso").Value = ""
    gIngresos.Columns.ColumnByFieldName("CodMoneda").Value = ""
    gIngresos.Columns.ColumnByFieldName("MontoIngreso").Value = 0
        
    gIngresos.Dataset.Post
    
    CargarControlLista strSQL, cboCobro, arrCobro(), Sel_Defecto
End Sub

Private Sub tabRegistroCompras_Click(PreviousTab As Integer)
    cmdAccion.Visible = False
    Select Case tabRegistroCompras.Tab
        Case 1, 2, 3
            cmdAccion.Visible = True
            If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vQuery)
            If strEstado = Reg_Defecto Then tabRegistroCompras.Tab = 0
    End Select
    
End Sub

Private Sub txtIgv_Change()

    Call FormatoCajaTexto(txtIgv, Decimales_Monto)
    
End Sub

Private Sub txtNumComprobante_LostFocus()
    txtNumComprobante.Text = Format(txtNumComprobante.Text, "000000")
End Sub

Private Sub txtSerieComprobante_LostFocus()
Dim rst As New ADODB.Recordset

    txtSerieComprobante.Text = Format(txtSerieComprobante.Text, "000")
    
    If Trim(txtNumComprobante.Text) = "" Then
        strSQL = "SELECT ISNULL(MAX(NumComprobante),0) AS NumComprobante FROM RegistroVenta " & _
                 "WHERE CodAdministradora = '" & gstrCodAdministradora & "' " & _
                   "AND CodFondo = '" & strCodFondo & "' " & _
                   "AND CodTipoComprobante = '" & strCodTipoComprobante & "' " & _
                   "AND SerieComprobante = '" & txtSerieComprobante.Text & "' "
                   
        rst.Open strSQL, gstrConnectConsulta, adOpenForwardOnly, adLockReadOnly
        If Not rst.EOF Then
            txtNumComprobante.Text = Format(rst.Fields("NumComprobante") + 1, "000000")
        End If
                       
        rst.Close
        Set rst = Nothing
    End If
End Sub

Private Sub txtSubTotal_Change()

    Call FormatoCajaTexto(txtSubTotal, Decimales_Monto)
                
End Sub

Private Sub txtSubTotal_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtSubTotal, Decimales_Monto)
    If KeyAscii = vbKeyReturn Then Call Calculos
    
End Sub

Private Sub txtTotal_Change()

    Call FormatoCajaTexto(txtTotal, Decimales_Monto)
    
End Sub

Private Sub txtTotal_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtTotal, Decimales_Monto)
    If KeyAscii = vbKeyReturn Then Call Calculos
    
End Sub

Private Sub FormatoGrillaIngresos(ByRef strMsgError As String) 'JCB
Dim rsGastos As New ADODB.Recordset
On Error GoTo err
    '********FORMATO GRILLA DE GASTOS
    rsGastos.Fields.Append "Item", adInteger, , adFldRowID
    rsGastos.Fields.Append "NumOrdenCobro", adVarChar, 10, adFldIsNullable
    rsGastos.Fields.Append "CodFile", adVarChar, 10, adFldIsNullable
    rsGastos.Fields.Append "CodAnalitica", adVarChar, 10, adFldIsNullable
    rsGastos.Fields.Append "DescripIngreso", adVarChar, 80, adFldIsNullable
    rsGastos.Fields.Append "CodMoneda", adVarChar, 2, adFldIsNullable
    rsGastos.Fields.Append "MontoIngreso", adDouble, , adFldIsNullable
    rsGastos.Fields.Append "CodDetalleFile", adVarChar, 10, adFldIsNullable
    
    rsGastos.Fields.Append "NumAnexo", adVarChar, 10, adFldIsNullable
    rsGastos.Fields.Append "NumContrato", adVarChar, 15, adFldIsNullable
    rsGastos.Fields.Append "NumDocumentoFisico", adVarChar, 15, adFldIsNullable
    rsGastos.Fields.Append "FechaDefinicion", adVarChar, 10, adFldIsNullable
    
    rsGastos.Open
    rsGastos.AddNew

    rsGastos.Fields("Item") = 1
    rsGastos.Fields("NumOrdenCobro") = ""
    rsGastos.Fields("CodFile") = ""
    rsGastos.Fields("CodAnalitica") = ""
    rsGastos.Fields("DescripIngreso") = ""
    rsGastos.Fields("CodMoneda") = ""
    rsGastos.Fields("MontoIngreso") = 0
    rsGastos.Fields("CodDetalleFile") = ""
    rsGastos.Fields("NumAnexo") = ""
    rsGastos.Fields("NumContrato") = ""
    rsGastos.Fields("NumDocumentoFisico") = ""
    rsGastos.Fields("FechaDefinicion") = ""
    
    Set gIngresos.DataSource = Nothing
    mostrarDatosGridSQL gIngresos, rsGastos, strMsgError
    If strMsgError <> "" Then GoTo err

Exit Sub
err:
If strMsgError = "" Then strMsgError = err.Description
End Sub

Public Sub Anular()
 
    With adoComm
        '*** Anula registro existente***
        
        If strEstado <> Reg_Adicion Then

            If MsgBox("¿Seguro que desea anular el Documento?", vbQuestion + vbYesNo, App.Title) = vbNo Then
                Exit Sub
            End If
            .CommandText = "{ call up_CNManAnulaRegistroVenta('" & _
                strCodFondo & "','" & gstrCodAdministradora & "'," & _
                CInt(lblNumSecuencial.Caption) & ") }"
            .Execute .CommandText
            MsgBox "El comprobante fue anulado con éxito.", vbExclamation
            
            Call Buscar
    
            cmdOpcion.Visible = True
            With tabRegistroCompras
                .TabEnabled(0) = True
                .Tab = 0
            End With

        Else
            If MsgBox("¿Seguro que desea anular un documento en blanco?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                If cboTipoComprobante.ListIndex >= 0 And txtSerieComprobante.Text <> "" And txtNumComprobante <> "" And cboMoneda.ListIndex >= 0 Then
                    'Se inserta un documento de cobro con estado anulado.
                    .CommandText = "Insert into RegistroVenta values ('" & strCodFondo & "','" & gstrCodAdministradora & "'," & CInt(lblNumSecuencial.Caption) & _
                    ",'" & Convertyyyymmdd(dtpFechaComprobante.Value) & "','" & arrTipoComprobante(cboTipoComprobante.ListIndex) & "','" & Convertyyyymmdd(dtpFechaComprobante.Value) & _
                    "','" & txtSerieComprobante & "','" & txtNumComprobante & "','','02','','','','','ANULADO','01','01','','" & arrMoneda(cboMoneda.ListIndex) & _
                    "','',0.0,0.0,0.0,'ANULADO','000','00000000','" & Convertyyyymmdd(dtpFechaComprobante.Value) & "',0,'03',0.0,'X','','19000101','','',0,'',0,'','')"
                    .Execute .CommandText
                    
                    MsgBox "El comprobante fue anulado con éxito.", vbExclamation
                    
                    Call Buscar
    
                    cmdOpcion.Visible = True
                    With tabRegistroCompras
                        .TabEnabled(0) = True
                        .Tab = 0
                    End With

                Else
                     MsgBox "Se necesita el Tipo de Documento, Moneda, Número de serie y Número de Documento para anular documento en blanco.", vbExclamation
                End If
            End If
        End If
    End With
        
End Sub

Private Function GenerarGlosaComprobante() As String

Dim adoAuxiliar As ADODB.Recordset
Dim strNumAnexo As String, strNumContrato As String, strNumDocumentoFisico As String, strFechaDefinicion As String
Dim strGlosaComprobante As String
Dim FechaOperacion, FechaVencimiento As Date

    gIngresos.Dataset.First
    strNumAnexo = Trim("" & gIngresos.Columns.ColumnByFieldName("NumAnexo").Value)
    strNumContrato = Trim("" & gIngresos.Columns.ColumnByFieldName("NumContrato").Value)
    strFechaDefinicion = Trim("" & gIngresos.Columns.ColumnByFieldName("FechaDefinicion").Value)
    strGlosaComprobante = ""
    
    Do While gIngresos.Dataset.EOF = False And gIngresos.Dataset.RecordCount > 0  'Or gIngresos.Dataset.BOF = False
    
        If strCodTipoComprobante <> "07" Then
            If InStr(strNumDocumentoFisico, Trim("" & gIngresos.Columns.ColumnByFieldName("NumDocumentoFisico").Value)) = 0 Then
                strNumDocumentoFisico = strNumDocumentoFisico & Trim("" & gIngresos.Columns.ColumnByFieldName("NumDocumentoFisico").Value) & ", "
            End If
            
        Else
            If gIngresos.Columns.ColumnByFieldName("NumOrdenCobro").Value <> Valor_Caracter Then
            adoComm.CommandText = "Select NumIngreso from OrdenCobro where CodFondo = '" & strCodFondo & "' and NumOrdenCobro = '" & _
                                    gIngresos.Columns.ColumnByFieldName("NumOrdenCobro").Value & "'"
            Set adoAuxiliar = adoComm.Execute
            
            adoComm.CommandText = "Select NumOperacion from FondoIngreso where CodFondo = '" & strCodFondo & "' and NumIngreso = " & _
                                    adoAuxiliar("NumIngreso")
            Set adoAuxiliar = adoComm.Execute
            
            adoComm.CommandText = "Select FechaOperacion, CodFile, CodAnalitica from InversionOperacion where CodFondo = '" & strCodFondo & "' and NumOperacion = '" & _
                                    adoAuxiliar("NumOperacion") & "'"
            Set adoAuxiliar = adoComm.Execute

            FechaOperacion = adoAuxiliar("FechaOperacion")
            
            adoComm.CommandText = "Select FechaVencimiento from InversionOperacion where CodFondo = '" & strCodFondo & "' and CodFile = '" & _
                                    adoAuxiliar("CodFile") & "' and CodAnalitica = '" & adoAuxiliar("CodAnalitica") & "' and TipoOperacion = '" & Codigo_Orden_Compra & "'"
            Set adoAuxiliar = adoComm.Execute
            
            FechaVencimiento = adoAuxiliar("FechaVencimiento")
            
            strGlosaComprobante = strGlosaComprobante & "INTERESES CORRIDOS DEL " & FechaOperacion & " AL " & FechaVencimiento & vbNewLine
            End If
        End If
        
        gIngresos.Dataset.Next

    Loop
    
    If strCodTipoComprobante <> "07" Then

        If Len(strNumDocumentoFisico) > 0 Then strNumDocumentoFisico = Left(strNumDocumentoFisico, Len(strNumDocumentoFisico) - 2)
        
        strGlosaComprobante = "POR " & cboIngreso.Text & ", SEGUN ANEXO " & strNumAnexo & " DEL CONTRATO " & strNumContrato & " CON FECHA DE OPERACIÓN " & strFechaDefinicion & " DE LOS DOCUMENTOS " & strNumDocumentoFisico
    End If
    
    GenerarGlosaComprobante = strGlosaComprobante

End Function

'MEVH 07/06/2012
Public Sub SubImprimir(Index As Integer)

    Dim strSeleccionRegistro    As String
    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()

    Select Case Index
        Case 1
            If Index = 1 Then gstrNameRepo = "RegistroVenta"
'           If index = 2 Then gstrNameRepo = "RegistroComprasParte2"
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

Private Sub BloquearControles(flag As Boolean)
    
    dtpFechaComprobante.Enabled = flag
    txtSerieComprobante.Enabled = flag
    txtNumComprobante.Enabled = flag
    cboMoneda.Enabled = flag
    cmdContratante.Enabled = flag
    cboIngreso.Enabled = flag
    txtDescripcion.Enabled = flag
    cboAfectacion.Enabled = flag
    cboCreditoFiscal.Enabled = flag
    cmdAccion.Button(0).Enabled = flag
    cmdAccion.Button(1).Enabled = flag
    cboCobro.Enabled = flag
    cmdCobro.Enabled = flag
    cboTipoComprobante.Enabled = flag
    cmdAdicionarCobro.Enabled = flag
    cmdEliminarCobro.Enabled = flag
    
    If flag Then
        Call ValidarPermisoUsoControl(Trim(gstrLogin), Me, Trim(App.Title) + Separador_Codigo_Objeto + _
        gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)

    End If
    
    
End Sub



Private Sub ucBotonEdicion21_GotFocus()

End Sub
