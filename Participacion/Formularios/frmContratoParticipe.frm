VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmContratoParticipe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contrato Partícipe"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8175
   ScaleWidth      =   15315
   Visible         =   0   'False
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   13080
      TabIndex        =   7
      Top             =   7320
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
      Caption0        =   "&Salir"
      Tag0            =   "9"
      Visible0        =   0   'False
      ToolTipText0    =   "Salir"
      UserControlWidth=   1200
   End
   Begin TAMControls2.ucBotonEdicion2 cmdOpcion 
      Height          =   735
      Left            =   720
      TabIndex        =   6
      Top             =   7320
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   1296
      Buttons         =   4
      Caption0        =   "&Nuevo"
      Tag0            =   "0"
      Visible0        =   0   'False
      ToolTipText0    =   "Nuevo"
      Caption1        =   "&Modificar"
      Tag1            =   "3"
      Visible1        =   0   'False
      ToolTipText1    =   "Modificar"
      Caption2        =   "&Eliminar"
      Tag2            =   "4"
      Visible2        =   0   'False
      ToolTipText2    =   "Eliminar"
      Caption3        =   "&Buscar"
      Tag3            =   "5"
      Visible3        =   0   'False
      ToolTipText3    =   "Buscar"
      UserControlWidth=   5700
   End
   Begin TabDlg.SSTab tabContrato 
      Height          =   7080
      Left            =   30
      TabIndex        =   30
      Top             =   60
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   12488
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
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
      TabPicture(0)   =   "frmContratoParticipe.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraContrato(0)"
      Tab(0).Control(1)=   "tdgConsulta"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "frmContratoParticipe.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraContrato(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraContrato(2)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame fraContrato 
         Caption         =   "Dirección Envío Información"
         Height          =   4155
         Index           =   2
         Left            =   8790
         TabIndex        =   44
         Top             =   1710
         Width           =   5880
         Begin VB.ComboBox cboPais 
            Height          =   315
            ItemData        =   "frmContratoParticipe.frx":0038
            Left            =   2265
            List            =   "frmContratoParticipe.frx":003F
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   1695
            Width           =   3315
         End
         Begin VB.ComboBox cboDireccionPostal 
            ForeColor       =   &H80000006&
            Height          =   315
            Left            =   2265
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   360
            Width           =   3315
         End
         Begin VB.ComboBox cboProvincia 
            Height          =   315
            Left            =   2265
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   2435
            Width           =   3315
         End
         Begin VB.ComboBox cboDistrito 
            Height          =   315
            Left            =   2265
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   2805
            Width           =   3315
         End
         Begin VB.ComboBox cboDepartamento 
            Height          =   315
            Left            =   2265
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   2065
            Width           =   3315
         End
         Begin VB.ComboBox cboAgencia 
            Height          =   315
            Left            =   2265
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   3545
            Width           =   3315
         End
         Begin VB.TextBox txtDireccion1 
            Height          =   315
            Left            =   300
            MaxLength       =   45
            MultiLine       =   -1  'True
            TabIndex        =   22
            Top             =   840
            Width           =   5265
         End
         Begin VB.TextBox txtDireccion2 
            Height          =   315
            Left            =   300
            MaxLength       =   45
            MultiLine       =   -1  'True
            TabIndex        =   23
            Top             =   1230
            Width           =   5265
         End
         Begin VB.ComboBox cboSucursal 
            Height          =   315
            Left            =   2265
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   3175
            Width           =   3315
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Distrito"
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
            Height          =   210
            Index           =   11
            Left            =   300
            TabIndex        =   51
            Top             =   2825
            Width           =   1275
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Pais"
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
            Height          =   210
            Index           =   7
            Left            =   300
            TabIndex        =   50
            Top             =   1715
            Width           =   1275
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Provincia"
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
            Height          =   210
            Index           =   10
            Left            =   300
            TabIndex        =   49
            Top             =   2455
            Width           =   1275
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Departamento"
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
            Height          =   210
            Index           =   8
            Left            =   300
            TabIndex        =   48
            Top             =   2085
            Width           =   1275
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Dirección Postal"
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
            Index           =   6
            Left            =   300
            TabIndex        =   47
            Top             =   405
            Width           =   1395
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Agencia"
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
            Left            =   300
            TabIndex        =   46
            Top             =   3565
            Width           =   960
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Sucursal"
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
            Left            =   300
            TabIndex        =   45
            Top             =   3195
            Width           =   960
         End
      End
      Begin VB.Frame fraContrato 
         Caption         =   "Definición"
         Height          =   6195
         Index           =   1
         Left            =   330
         TabIndex        =   32
         Top             =   570
         Width           =   14505
         Begin VB.ComboBox cboComisionista 
            Height          =   315
            Left            =   2220
            Style           =   2  'Dropdown List
            TabIndex        =   61
            Top             =   2565
            Visible         =   0   'False
            Width           =   4245
         End
         Begin TAMControls2.ucBotonEdicion2 cmdAccion 
            Height          =   735
            Left            =   11160
            TabIndex        =   56
            Top             =   5370
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
         Begin VB.ComboBox cboTipoContrato 
            Height          =   315
            Left            =   2220
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   450
            Width           =   2745
         End
         Begin VB.ComboBox cboTipoMancomuno 
            Height          =   315
            Left            =   6540
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   450
            Visible         =   0   'False
            Width           =   2610
         End
         Begin VB.ComboBox cboPromotor 
            Height          =   315
            Left            =   2220
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   3990
            Width           =   5925
         End
         Begin VB.TextBox txtContrato 
            Enabled         =   0   'False
            Height          =   315
            Left            =   10770
            Locked          =   -1  'True
            TabIndex        =   53
            Top             =   450
            Width           =   3555
         End
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
            Left            =   7770
            TabIndex        =   9
            ToolTipText     =   "Búsqueda de Cliente"
            Top             =   885
            Width           =   315
         End
         Begin VB.ComboBox cboTipoDocumento 
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2220
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1350
            Width           =   4245
         End
         Begin VB.TextBox txtNumDocumentoCliente 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2220
            TabIndex        =   11
            Top             =   1755
            Width           =   4245
         End
         Begin VB.ComboBox cboCustodia 
            Height          =   315
            ItemData        =   "frmContratoParticipe.frx":004C
            Left            =   2220
            List            =   "frmContratoParticipe.frx":0053
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   4815
            Width           =   2175
         End
         Begin VB.CommandButton cmdMancomuno 
            Caption         =   "&Ver Distribución de Mancomunos"
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
            Height          =   375
            Left            =   4920
            TabIndex        =   18
            Top             =   4770
            Width           =   3195
         End
         Begin VB.CommandButton cmdCuenta 
            Caption         =   "Cuen&ta"
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
            Height          =   375
            Left            =   720
            TabIndex        =   19
            Top             =   5460
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CommandButton cmdRepresentante 
            Caption         =   "&Representante"
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
            Height          =   375
            Left            =   2700
            TabIndex        =   20
            Top             =   5460
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Label lblcomisionista 
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
            Height          =   300
            Index           =   19
            Left            =   360
            TabIndex        =   60
            Top             =   2640
            Visible         =   0   'False
            Width           =   1710
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Tipo Contrato"
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
            Height          =   270
            Index           =   18
            Left            =   360
            TabIndex        =   58
            Top             =   480
            Width           =   1710
         End
         Begin VB.Line Line1 
            X1              =   390
            X2              =   8250
            Y1              =   3000
            Y2              =   3000
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Ejecutivo Comercial"
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
            TabIndex        =   55
            Top             =   4080
            Width           =   1680
         End
         Begin VB.Label lblTipoMancomuno 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5070
            TabIndex        =   17
            Top             =   5730
            Visible         =   0   'False
            Width           =   3975
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Tipo Mancomuno"
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
            Height          =   210
            Index           =   17
            Left            =   5040
            TabIndex        =   42
            Top             =   480
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.Label lblFechaIngreso 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "dd/mm/yyyy"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2220
            TabIndex        =   15
            Top             =   4410
            Width           =   2175
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Fecha de Ingreso"
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
            Height          =   270
            Index           =   16
            Left            =   390
            TabIndex        =   41
            Top             =   4455
            Width           =   1710
         End
         Begin VB.Label lblDescripParticipe 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2220
            TabIndex        =   13
            Top             =   3180
            Width           =   5880
         End
         Begin VB.Label lblDescrip 
            Caption         =   "En Custodia ?"
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
            Height          =   210
            Index           =   14
            Left            =   390
            TabIndex        =   40
            Top             =   4860
            Width           =   1275
         End
         Begin VB.Label lblCodParticipe 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2220
            TabIndex        =   14
            Top             =   3600
            Width           =   3555
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Código Partícipe"
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
            Height          =   300
            Index           =   4
            Left            =   390
            TabIndex        =   39
            Top             =   3645
            Width           =   1710
         End
         Begin VB.Label lblCodCliente 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2220
            TabIndex        =   12
            Top             =   2160
            Width           =   4245
         End
         Begin VB.Label lblDescripTitular 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2220
            TabIndex        =   8
            Top             =   900
            Width           =   5475
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Código Cliente"
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
            Height          =   300
            Index           =   3
            Left            =   360
            TabIndex        =   38
            Top             =   2205
            Width           =   1710
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Tipo de Documento"
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
            Height          =   300
            Index           =   0
            Left            =   330
            TabIndex        =   37
            Top             =   1425
            Width           =   1710
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Num. Documento"
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
            Height          =   300
            Index           =   1
            Left            =   360
            TabIndex        =   36
            Top             =   1830
            Width           =   1725
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Cliente Titular"
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
            Height          =   270
            Index           =   2
            Left            =   360
            TabIndex        =   35
            Top             =   975
            Width           =   1710
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Num. Contrato"
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
            Height          =   210
            Index           =   9
            Left            =   9450
            TabIndex        =   34
            Top             =   480
            Width           =   1275
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Participe"
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
            Height          =   210
            Index           =   5
            Left            =   390
            TabIndex        =   33
            Top             =   3225
            Width           =   1275
         End
      End
      Begin VB.Frame fraContrato 
         Caption         =   "Criterios de Búsqueda"
         Height          =   1575
         Index           =   0
         Left            =   -74640
         TabIndex        =   31
         Top             =   540
         Width           =   14565
         Begin VB.OptionButton optParticipe 
            Caption         =   "Num. Documento"
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
            Left            =   720
            TabIndex        =   4
            Top             =   915
            Width           =   1890
         End
         Begin VB.TextBox txtCodParticipe 
            Height          =   285
            Left            =   3105
            MaxLength       =   20
            TabIndex        =   1
            Top             =   480
            Width           =   2940
         End
         Begin VB.OptionButton optParticipe 
            Caption         =   "Código Partícipe"
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
            Left            =   720
            TabIndex        =   0
            Top             =   495
            Value           =   -1  'True
            Width           =   1905
         End
         Begin VB.OptionButton optParticipe 
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
            Index           =   2
            Left            =   6720
            TabIndex        =   2
            Top             =   495
            Width           =   1905
         End
         Begin VB.TextBox txtDescripcion 
            Enabled         =   0   'False
            Height          =   315
            Left            =   8745
            MaxLength       =   50
            TabIndex        =   3
            Top             =   480
            Width           =   5130
         End
         Begin VB.TextBox txtNumDocumento 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3105
            MaxLength       =   15
            TabIndex        =   5
            Top             =   900
            Width           =   2940
         End
         Begin VB.Label lblContador 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   9105
            TabIndex        =   52
            Top             =   870
            Width           =   2940
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmContratoParticipe.frx":0064
         Height          =   4155
         Left            =   -74640
         OleObjectBlob   =   "frmContratoParticipe.frx":007E
         TabIndex        =   43
         Top             =   2340
         Width           =   14565
      End
   End
End
Attribute VB_Name = "frmContratoParticipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrSiNo()               As String, arrDireccionPostal()         As String
Dim arrPais()               As String, arrDepartamento()            As String
Dim arrProvincia()          As String, arrDistrito()                As String
Dim arrSucursal()           As String, arrAgencia()                 As String
Dim arrPromotor()           As String, arrTipoMancomuno()           As String
Dim arrTipoContrato()       As String, arrComisionista()            As String

Dim strCodSiNo              As String, strCodDireccionPostal        As String
Dim strCodPais              As String, strCodDepartamento           As String
Dim strCodProvincia         As String, strCodDistrito               As String
Dim strCodSucursal          As String, strCodAgencia                As String
Dim strCodPromotor          As String, strCodTipoMancomuno          As String
Dim strCodTipoDocumento     As String, strSql                       As String
Dim strEstado               As String, strClaseParticipe            As String
Dim strTipoContrato         As String

Dim blnCliente              As Boolean
Dim strApellidoPaterno      As String
Dim strApellidoMaterno      As String
Dim strNombres              As String
Dim strRazonSocial          As String

Dim strCodComisionista   As String

Dim adoConsulta             As ADODB.Recordset
Dim indSortAsc              As Boolean, indSortDesc                 As Boolean

Public Sub SubImprimir(Index As Integer)

    Dim strSeleccionRegistro    As String
    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()

    If tabContrato.Tab = 1 Then Exit Sub
    
    Select Case Index
        Case 1
            gstrNameRepo = "ParticipeContrato"
                        
            '*** Lista de Clientes por rango de fecha ***
            strSeleccionRegistro = "{Participe.FechaIngreso} IN 'Fch1' TO 'Fch2'"
            gstrSelFrml = strSeleccionRegistro
            frmRangoFecha.Show vbModal
                
            If gstrSelFrml <> "0" Then
            
            Set frmReporte = New frmVisorReporte

            ReDim aReportParamS(1)
            ReDim aReportParamFn(4)
            ReDim aReportParamF(4)
                        
            aReportParamFn(0) = "Usuario"
            aReportParamFn(1) = "Hora"
            aReportParamFn(2) = "NombreEmpresa"
            aReportParamFn(3) = "FechaDel"
            aReportParamFn(4) = "FechaAl"
            
            aReportParamF(0) = gstrLogin
            aReportParamF(1) = Format(Time(), "hh:mm:ss")
            aReportParamF(2) = gstrNombreEmpresa & Space(1)
            aReportParamF(3) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10)
            aReportParamF(4) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10)
                        
            aReportParamS(0) = Convertyyyymmdd(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10))
            aReportParamS(1) = Convertyyyymmdd(DateAdd("d", 1, CVDate(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10))))
            End If
            
        Case 2
            gstrNameRepo = "ContratoAdministracion"
                        
            '*** Contrato de Administración de Cuotas de Participación ***
            Set frmReporte = New frmVisorReporte

            ReDim aReportParamS(0)
            ReDim aReportParamFn(4)
            ReDim aReportParamF(2)

            aReportParamFn(0) = "Usuario"
            aReportParamFn(1) = "Hora"
            aReportParamFn(2) = "NombreEmpresa"
            aReportParamFn(3) = "FechaDel"
            aReportParamFn(4) = "FechaAl"

            aReportParamF(0) = gstrLogin
            aReportParamF(1) = Format(Time(), "hh:mm:ss")
            aReportParamF(2) = gstrNombreEmpresa & Space(1)

                        
            aReportParamS(0) = CStr(Trim(tdgConsulta.Columns(0)))

    End Select
    
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

Public Sub Adicionar()

'    If blnCliente Then
    
        frmMainMdi.stbMdi.Panels(3).Text = "Adicionar contrato..."
        
        strEstado = Reg_Adicion
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabContrato
            .TabEnabled(0) = False
            .Tab = 1
        End With
        Call Deshabilita
'    Else
'        MsgBox "Solo se puede adicionar contratos para clientes", vbCritical
'        Exit Sub
'    End If
    
End Sub

Private Sub LlenarFormulario(strModo As String)

    Dim strSql As String
    Dim intRegistro As Integer
    Dim adoRegistro As ADODB.Recordset
    
    Select Case strModo
        Case Reg_Adicion
        
            lblCodParticipe.Caption = "GENERADO"
            
            txtContrato.Text = "GENERADO"
            
            gstrCodParticipe = Trim(lblCodParticipe.Caption)
                        
            cmdBusqueda.Enabled = True
            
            cboTipoDocumento.ListIndex = -1
            If cboTipoDocumento.ListCount > 0 Then cboTipoDocumento.ListIndex = 0
            
            txtNumDocumentoCliente.Text = Valor_Caracter
            lblCodCliente.Caption = Valor_Caracter
            lblDescripTitular.Caption = Valor_Caracter
            lblDescripParticipe.Caption = Valor_Caracter
            
            'txtContrato.Enabled = True
            'txtContrato.Text = Valor_Caracter
            lblFechaIngreso.Caption = CStr(gdatFechaActual)
            
            cboCustodia.ListIndex = -1
            intRegistro = ObtenerItemLista(arrSiNo(), Codigo_Respuesta_Si)
            If intRegistro >= 0 Then cboCustodia.ListIndex = intRegistro
            cboCustodia.Enabled = False
                                    
            cboDireccionPostal.ListIndex = -1
            If cboDireccionPostal.ListCount > 0 Then cboDireccionPostal.ListIndex = 0
                                    
            cboPromotor.ListIndex = -1
            If cboPromotor.ListCount > 0 Then cboPromotor.ListIndex = 0
            
            txtDireccion1.Text = Valor_Caracter: txtDireccion1.Enabled = False
            txtDireccion2.Text = Valor_Caracter: txtDireccion2.Enabled = False

            cboPais.ListIndex = -1
            If cboPais.ListCount > 0 Then cboPais.ListIndex = 0
            
            cboPais.Enabled = False: cboDepartamento.Enabled = False
            cboProvincia.Enabled = False: cboDistrito.Enabled = False
                        
            intRegistro = ObtenerItemLista(arrSucursal(), gstrCodSucursal)
            If intRegistro >= 0 Then cboSucursal.ListIndex = intRegistro
                                    
            strCodTipoMancomuno = Codigo_Tipo_Mancomuno_Individual
            
            cmdMancomuno.Enabled = False
            cmdCuenta.Enabled = False: cmdRepresentante.Enabled = False
                                  
            cmdBusqueda.SetFocus
                        
        Case Reg_Edicion
            Dim strCodParticipe As String

            Set adoRegistro = New ADODB.Recordset

            strCodParticipe = Trim(tdgConsulta.Columns(0))
            gstrCodParticipe = strCodParticipe
            
            'Datos del participe
            adoComm.CommandText = "{ call up_ACSelDatosParametro(12,'" & strCodParticipe & "') }"
            Set adoRegistro = adoComm.Execute

            If Not adoRegistro.EOF Then
                lblCodParticipe.Caption = strCodParticipe
                lblCodCliente.Caption = Trim(adoRegistro("CodUnico"))
                
                intRegistro = ObtenerItemLista(garrTipoDocumento(), adoRegistro("TipoIdentidad"))
                If intRegistro >= 0 Then cboTipoDocumento.ListIndex = intRegistro
            
                txtNumDocumentoCliente.Text = Trim(adoRegistro("NumIdentidad"))
                                
                cmdBusqueda.Enabled = False
                
                strClaseParticipe = adoRegistro("ClaseParticipe")
                
                If adoRegistro("ClaseParticipe") = Codigo_Persona_Natural Or adoRegistro("ClaseParticipe") = Codigo_Persona_Mancomuno Then
                    lblDescripTitular.Caption = Trim(adoRegistro("ApellidoPaterno")) & Space(1) & Trim(adoRegistro("ApellidoMaterno")) & Space(1) & Trim(adoRegistro("Nombres"))
                    cmdMancomuno.Enabled = True
                    strApellidoPaterno = Trim(adoRegistro("ApellidoPaterno"))
                    strApellidoMaterno = Trim(adoRegistro("ApellidoMaterno"))
                    strNombres = Trim(adoRegistro("Nombres"))
                    strRazonSocial = ""
                End If
                
                If adoRegistro("ClaseParticipe") = Codigo_Persona_Mancomuno Then
                    lblDescripTitular.Caption = Trim(adoRegistro("ApellidoPaterno")) & Space(1) & Trim(adoRegistro("ApellidoMaterno")) & Space(1) & Trim(adoRegistro("Nombres"))
                    cmdMancomuno.Enabled = True
                    strApellidoPaterno = ""
                    strApellidoMaterno = ""
                    strNombres = ""
                    strRazonSocial = ""
                End If
                
                If adoRegistro("ClaseParticipe") = Codigo_Persona_Juridica Then
                    lblDescripTitular.Caption = Trim(adoRegistro("RazonSocial"))
                    cmdMancomuno.Enabled = False
                    strApellidoPaterno = ""
                    strApellidoMaterno = ""
                    strNombres = ""
                    strRazonSocial = Trim(adoRegistro("RazonSocial"))
                End If
                
                lblDescripParticipe.Caption = Trim(tdgConsulta.Columns(1))
                lblFechaIngreso.Caption = Trim(tdgConsulta.Columns(5))
                txtContrato.Text = Trim(adoRegistro("NumContrato"))
                txtContrato.Enabled = False

                If Trim(adoRegistro("IndCustodia")) = "" Then
                    intRegistro = ObtenerItemLista(arrSiNo(), Codigo_Respuesta_No)
                Else
                    intRegistro = ObtenerItemLista(arrSiNo(), Codigo_Respuesta_Si)
                End If
                If intRegistro >= 0 Then cboCustodia.ListIndex = intRegistro
                cboCustodia.Enabled = False

                strCodTipoMancomuno = adoRegistro("TipoMancomuno")
                lblTipoMancomuno.Caption = Trim(adoRegistro("DescripTipoMancomuno"))
                
                intRegistro = ObtenerItemLista(arrDireccionPostal(), adoRegistro("TipoCorreoPostal"))
                If intRegistro >= 0 Then cboDireccionPostal.ListIndex = intRegistro
                
                txtDireccion1.Text = Trim(adoRegistro("DireccionPostal1"))
                txtDireccion2.Text = Trim(adoRegistro("DireccionPostal2"))
                
                intRegistro = ObtenerItemLista(arrPais(), adoRegistro("CodPais"))
                If intRegistro >= 0 Then cboPais.ListIndex = intRegistro

                intRegistro = ObtenerItemLista(arrDepartamento(), adoRegistro("CodDepartamento"))
                If intRegistro >= 0 Then cboDepartamento.ListIndex = intRegistro

                intRegistro = ObtenerItemLista(arrProvincia(), adoRegistro("CodProvincia"))
                If intRegistro >= 0 Then cboProvincia.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrDistrito(), adoRegistro("CodDistrito"))
                If intRegistro >= 0 Then cboDistrito.ListIndex = intRegistro

                intRegistro = ObtenerItemLista(arrProvincia(), adoRegistro("CodProvincia"))
                If intRegistro >= 0 Then cboProvincia.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrSucursal(), adoRegistro("CodSucursal"))
                If intRegistro >= 0 Then cboSucursal.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrAgencia(), adoRegistro("CodAgenciaBancaria"))
                If intRegistro >= 0 Then cboAgencia.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrPromotor(), adoRegistro("CodPromotor"))
                If intRegistro >= 0 Then cboPromotor.ListIndex = intRegistro
            End If
            adoRegistro.Close: Set adoRegistro = Nothing
    
            cmdCuenta.Enabled = True: cmdRepresentante.Enabled = True
            
    End Select
    
End Sub

Private Sub ObtenerDatosCliente()

    Dim adoRegistro     As ADODB.Recordset
    Dim strClaseCliente  As String
    
    Set adoRegistro = New ADODB.Recordset
            
    adoComm.CommandText = "SELECT ValorParametro FROM AuxiliarParametro WHERE CodTipoParametro='TIPIDE' AND " & _
        "CodParametro='" & strCodTipoDocumento & "'"
    Set adoRegistro = adoComm.Execute
    
    If Not adoRegistro.EOF Then
        strClaseCliente = Trim(adoRegistro("ValorParametro"))
    End If
    adoRegistro.Close
    
    adoComm.CommandText = "{ call up_ACSelDatosParametro(36,'" & strCodTipoDocumento & "','" & Trim(txtNumDocumentoCliente.Text) & "','" & strClaseCliente & "') }"
    Set adoRegistro = adoComm.Execute
    
    If Not adoRegistro.EOF Then
        lblCodCliente.Caption = Trim(adoRegistro("CodUnico"))
        lblDescripParticipe.Caption = Trim(adoRegistro("DescripCliente"))
        lblDescripTitular.Caption = Trim(adoRegistro("DescripCliente"))
    End If
    adoRegistro.Close: Set adoRegistro = Nothing
    
End Sub

Public Sub Buscar()

    Set adoConsulta = New ADODB.Recordset
                    
    strEstado = Reg_Defecto
   
    strSql = "SELECT CodParticipe,Tabla1.DescripParametro TipoIdentidad,NumIdentidad,DescripParticipe,FechaIngreso,Tabla2.DescripParametro TipoMancomuno, TipoIdentidad CodTipoIdentidad, ClaseParticipe "
    strSql = strSql & "FROM ParticipeContrato JOIN AuxiliarParametro Tabla1 ON(Tabla1.CodParametro=ParticipeContrato.TipoIdentidad AND Tabla1.CodTipoParametro='TIPIDE') "
    strSql = strSql & "JOIN AuxiliarParametro Tabla2 ON(Tabla2.CodParametro=ParticipeContrato.TipoMancomuno AND Tabla2.CodTipoParametro='TIPMAN') "
    
    If Trim(txtCodParticipe.Text) <> "" And optParticipe(0).Value Then
        strSql = strSql & "WHERE CodParticipe='" & Trim(txtCodParticipe.Text) & "'"
    ElseIf Trim(txtNumDocumento.Text) <> "" And optParticipe(1).Value Then
        strSql = strSql & "WHERE NumIdentidad='" & Trim(txtNumDocumento.Text) & "'"
    ElseIf Trim(txtDescripcion.Text) <> "" And optParticipe(2).Value Then
        strSql = strSql & "WHERE DescripParticipe LIKE '%" & Trim(txtDescripcion.Text) & "%'"
    End If
    
        strSql = strSql & " AND EstadoParticipe='01'"
                            
    tdgConsulta.Columns(0).Caption = "Código"
    tdgConsulta.Columns(0).DataField = "CodParticipe"
    
    tdgConsulta.Columns(1).Caption = "Descripción"
    tdgConsulta.Columns(1).DataField = "DescripParticipe"
    
    tdgConsulta.Columns(2).Caption = "Tipo Mancomuno"
    tdgConsulta.Columns(2).DataField = "TipoMancomuno"
    
    tdgConsulta.Columns(3).Caption = "Tipo Ident."
    tdgConsulta.Columns(3).DataField = "TipoIdentidad"
    
    tdgConsulta.Columns(4).Caption = "Número"
    tdgConsulta.Columns(4).DataField = "NumIdentidad"
    
    If Not tdgConsulta.Columns(5).Visible Then tdgConsulta.Columns(5).Visible = True
        
    tdgConsulta.Columns(5).Caption = "Ingreso"
    tdgConsulta.Columns(5).DataField = "FechaIngreso"
                            
    tdgConsulta.Columns(6).Caption = "CodTipoIdentidad"
    tdgConsulta.Columns(6).DataField = "CodTipoIdentidad"
    
    tdgConsulta.Columns(7).Caption = "ClaseParticipe"
    tdgConsulta.Columns(7).DataField = "ClaseParticipe"
    

                
    If strSql = Valor_Caracter Then Exit Sub
    
    Me.MousePointer = vbHourglass
        
    With adoConsulta
        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open strSql
    End With
        
    tdgConsulta.DataSource = adoConsulta
    
    Me.lblContador.Caption = "Se encontraron " & adoConsulta.RecordCount & " registros"
    
    If adoConsulta.RecordCount > 0 Then strEstado = Reg_Consulta
                
    Me.MousePointer = vbDefault
    
End Sub

Public Sub Cancelar()

    cmdOpcion.Visible = True
    With tabContrato
        .TabEnabled(0) = True
        .TabEnabled(1) = False
        .Tab = 0
    End With
    Call Buscar
    
End Sub

Private Sub Deshabilita()

    txtDireccion1.Enabled = False
    txtDireccion2.Enabled = False

    cboPais.Enabled = False
    cboDepartamento.Enabled = False
    cboProvincia.Enabled = False
    cboDistrito.Enabled = False
    
End Sub

Public Sub Eliminar()

    Dim strSql  As String
    
    Me.MousePointer = vbHourglass
    If strEstado = Reg_Consulta Or strEstado = Reg_Edicion Then
        If MsgBox(Mensaje_Eliminacion, vbQuestion + vbYesNo, gstrNombreEmpresa) = vbYes Then
            
            adoComm.CommandText = "Update ParticipeContrato set EstadoParticipe='02' where CodParticipe = '" & tdgConsulta.Columns("CodParticipe") & "'"
        
            adoConn.Execute adoComm.CommandText
            
            Buscar

        End If
    End If
    Me.MousePointer = vbDefault

'    If tabBusqueda.Tab = 0 Then
'        MsgBox "Solo se puede eliminar un contrato", vbCritical, Me.Caption
'        Exit Sub
'    End If
'
'    If strEstado = "Consulta" Or strEstado = "Edicion" Then
'        Dim adoSolici     As New ADODB.Recordset
'    Dim adoCertif     As New ADODB.Recordset
'    Dim adoCertif2    As New ADODB.Recordset
'    Dim adoPersre     As New ADODB.Recordset
'    Dim strCodPart    As String
'    Dim bnlFlgEjecuta As Boolean
'    Dim adoRecord     As New ADODB.Recordset
'
'        bnlFlgEjecuta = False
'        If (strNivAcceso = "1") Then
'            bnlFlgEjecuta = True
'        ElseIf (strNivAcceso = "3") Then
'            adoComm.CommandText = "Sp_INF_SelectMantPart02 '43', '" & gstrParticipe & "'"
'            Set adoRecord = adoComm.Execute
'            If (adoRecord!FCH_INGR = gstrFechaAct) And (Trim$(adoRecord!USU_ACTU) = gstrLogin) Then
'                bnlFlgEjecuta = True
'            Else
'                MsgBox "Acceso Denegado....Permiso de solo Lectura", vbInformation, gstrNombreEmpresa
'                Exit Sub
'            End If
'            adoRecord.Close: Set adoRecord = Nothing
'        ElseIf (strNivAcceso = "5") Then
'            MsgBox "Acceso Denegado....Permiso de solo Lectura", vbInformation, gstrNombreEmpresa
'            Exit Sub
'        End If
'
'        If bnlFlgEjecuta = True Then
'            adoComm.CommandText = "Sp_INF_SelectMantPart02 '44', '" & gstrParticipe & "'"
'            Set adoSolici = adoComm.Execute
'
'            adoComm.CommandText = "Sp_INF_SelectMantPart02 '45', '" & gstrParticipe & "'"
'            Set adoCertif = adoComm.Execute
'
'            adoComm.CommandText = "Sp_INF_SelectMantPart02 '46', '" & gstrParticipe & "'"
'            Set adoCertif2 = adoComm.Execute
'
'            If (adoCertif2!TOTAL > 0) Then
'                MsgBox "Error >> El partícipe tiene operaciones vigentes .....!", vbCritical, gstrNombreEmpresa
'                MousePointer = vbDefault: Exit Sub
'            End If
'
'            If (adoSolici.EOF And adoCertif.EOF) Then
'                MousePointer = vbDefault
'                If MsgBox("Esta Seguro de Eliminar a :" & Chr(10) & Chr(13) _
'                    & "Partícipe          : " & Trim$(lblDatosPersona(2).Caption) & Chr(10) & Chr(13) _
'                    , vbQuestion + vbYesNo, gstrNombreEmpresa) = vbYes Then
'
'                    MousePointer = vbHourglass
'
'                    adoComm.CommandText = "Sp_INF_SelectMantPart02 '47', '" & gstrParticipe & "'"
'                    adoConn.Execute adoComm.CommandText
'
'                    adoComm.CommandText = "Sp_INF_SelectMantPart02 '48', '" & gstrParticipe & "'"
'                    adoConn.Execute adoComm.CommandText
'
'                    adoComm.CommandText = "Sp_INF_SelectMantPart02 '49', '" & gstrParticipe & "'"
'                    adoConn.Execute adoComm.CommandText
'
'                    adoComm.CommandText = "Sp_INF_SelectMantPart02 '50', '" & gstrParticipe & "'"
'                    Set adoPersre = adoComm.Execute
'
'                    Do Until adoPersre.EOF
'                        If Trim(adoPersre!COD_PAR2) <> "" Then
'                            adoComm.CommandText = "Sp_INF_SelectMantPart02 '51', '" & adoPersre!COD_PAR2 & "'"
'                            adoConn.Execute adoComm.CommandText
'                        End If
'                        adoPersre.MoveNext
'                    Loop
'
'                    adoComm.CommandText = "Sp_INF_SelectMantPart02 '52', '" & gstrParticipe & "'"
'                    adoConn.Execute adoComm.CommandText
'
'                    adoComm.CommandText = "Sp_INF_SelectMantPart02 '53', '" & gstrParticipe & "'"
'                    adoConn.Execute adoComm.CommandText
'
'                    MousePointer = vbDefault
'                    MsgBox "Registro eliminado exitosamente....!"
'                    LimpiaScr
'                    ucButEdi1.Visible = True
'                    With tabContrato
'                        .TabEnabled(0) = True
'                        .Tab = 0
'                    End With
'                    CargarGrillaParticipes
'                End If
'            Else
'                If (adoCertif2!TOTAL = 0) Then
'                    adoComm.CommandText = "Sp_INF_SelectMantPart02 '54', '" & gstrFechaAct & "', '" & gstrLogin & "', '" & gstrParticipe & "'"
'                    adoConn.Execute adoComm.CommandText
'                End If
'                MsgBox "Registro eliminado exitosamente....!"
'                LimpiaScr
'                ucButEdi1.Visible = True
'                With tabContrato
'                    .TabEnabled(0) = True
'                    .Tab = 0
'                End With
'                CargarGrillaParticipes
'            End If
'        End If
'    End If
    
End Sub

Private Sub Habilita()

    txtDireccion1.Enabled = True
    txtDireccion2.Enabled = True

    cboPais.Enabled = True
    cboDepartamento.Enabled = True
    cboProvincia.Enabled = True
    cboDistrito.Enabled = True
    
End Sub

Public Sub Imprimir()

End Sub

Public Sub Grabar()

    Dim adoRegistro As ADODB.Recordset
    Dim intRegistro As Integer
    Dim adoError As ADODB.Error
    Dim strErrMsg As String
    Dim intAccion As Long
    Dim lngNumError As Long
    Dim strNumContrato As String
                
    If strEstado = Reg_Defecto Then Exit Sub
    
    If Not TodoOK() Then Exit Sub
    
    On Error GoTo CtrlError
    
    If strEstado = Reg_Adicion Then
        Dim strSql              As String
        Dim strTipoIdentidad    As String, strNumIdentidad  As String
        Dim strClaseCliente     As String
        
        Me.MousePointer = vbHourglass
                                
        '*** Guardar Contrato ***
        With adoComm
            strSql = "{ call up_PRManContratoParticipe('" & _
                Trim(lblCodParticipe.Caption) & "','" & _
                Trim(lblCodCliente.Caption) & "','"
            
            Set adoRegistro = New ADODB.Recordset
        
            'Datos del titular
            .CommandText = "{ call up_ACSelDatosParametro(7,'" & Trim(lblCodCliente.Caption) & "') }"
            Set adoRegistro = .Execute
            
            If Not adoRegistro.EOF Then
                strClaseCliente = Trim(adoRegistro("ClaseCliente"))
                
                strSql = strSql & strClaseCliente & "','" & _
                    Trim(adoRegistro("ApellidoPaterno")) & "','" & _
                    Trim(adoRegistro("ApellidoMaterno")) & "','" & _
                    Trim(adoRegistro("Nombres")) & "','" & _
                    Trim(adoRegistro("RazonSocial")) & "','" & _
                    Trim(adoRegistro("DescripCliente")) & "','"
                    
                strTipoIdentidad = Trim(adoRegistro("TipoIdentidad"))
                
                strSql = strSql & strTipoIdentidad & "','"
                
                strNumIdentidad = Trim(adoRegistro("NumIdentidad"))
                
                strSql = strSql & strNumIdentidad & "','" & _
                    Trim(adoRegistro("SexoCliente")) & "','" & _
                    Trim(adoRegistro("EstadoCivil")) & "','" & _
                    Convertyyyymmdd(adoRegistro("FechaNacimiento")) & "','" & _
                    Trim(txtDireccion1.Text) & "','" & _
                    Trim(txtDireccion2.Text) & "','" & _
                    "','" & _
                    strCodPais & "','" & _
                    strCodDepartamento & "','" & _
                    strCodProvincia & "','" & _
                    strCodDistrito & "','" & _
                    Trim(adoRegistro("CodNacionalidad")) & "','" & _
                    Trim(adoRegistro("NumTelefono")) & "','" & _
                    Trim(adoRegistro("NumFax")) & "','" & _
                    Trim(adoRegistro("NumIdentidad")) & "','"
            End If
            adoRegistro.Close: Set adoRegistro = Nothing
            
            strSql = strSql & Convertyyyymmdd(CVDate(lblFechaIngreso.Caption)) & "','" & _
                strCodPromotor & "','" & _
                "X','" & _
                strCodSucursal & "','" & _
                strCodAgencia & "','" & _
                strCodDireccionPostal & "','" & _
                Convertyyyymmdd(CVDate(Valor_Fecha)) & "','" & _
                Convertyyyymmdd(CVDate(Valor_Fecha)) & "','" & _
                strCodTipoMancomuno & "','" & _
                Trim(txtContrato.Text) & "','" & _
                "','" & _
                Estado_Activo & "','" & _
                Codigo_FormaIngreso_Suscripcion & "','" & _
                gstrLogin & "','" & Convertyyyymmdd(gdatFechaActual) & "','" & _
                gstrLogin & "','" & Convertyyyymmdd(gdatFechaActual) & "','" & _
                "I') }"
            .CommandText = strSql
            
            'adoConn.Execute .CommandText, intRegistro
            
            Set adoRegistro = New ADODB.Recordset
            
            Set adoRegistro = adoConn.Execute(.CommandText)

            If Not adoRegistro.EOF Then
                gstrCodParticipe = adoRegistro("CodParticipe")
                strNumContrato = adoRegistro("NumContrato")
                txtContrato.Text = strNumContrato
            End If
            
            adoRegistro.Close: Set adoRegistro = Nothing
            
        End With

        
        Me.MousePointer = vbDefault
        
        'MsgBox Mensaje_Adicion_Exitosa, vbExclamation
        MsgBox "Se ha Creado Exitosamente el Contrato Nro. " & strNumContrato & ".", vbOKOnly + vbInformation, Me.Caption
       
        If strClaseCliente = Codigo_Persona_Natural Then
            If strTipoContrato = Codigo_Tipo_Contrato_Mancomuno Then
                MsgBox "A Continuación se deberán Ingresar los Participes Mancómunos.", vbOKOnly + vbInformation, Me.Caption
'                intRegistro = ObtenerItemLista(arrTipoMancomuno(), Codigo_Tipo_Mancomuno_Conjunto)
'                If intRegistro >= 0 Then cboTipoMancomuno.ListIndex = intRegistro
                Call cmdMancomuno_Click
            End If
            cmdMancomuno.Enabled = True
'        Else  'Juridico
'            MsgBox "A Continuación se deberán Ingresar los Representantes para el Participe Jurídico.", vbOK + vbInformation, Me.Caption
'            cmdRepresentante_Click
'            cmdMancomuno.Enabled = False
        End If
        
        cmdCuenta.Enabled = True: cmdRepresentante.Enabled = True
        
        lblCodParticipe.Caption = gstrCodParticipe
        txtCodParticipe.Text = Trim(lblCodParticipe.Caption)
    
    End If
    
    If strEstado = Reg_Edicion Then
        If MsgBox(Mensaje_Edicion, vbQuestion + vbYesNo, gstrNombreEmpresa) = vbNo Then Exit Sub
        
        Me.MousePointer = vbHourglass
                                            
        If strCodTipoMancomuno <> Codigo_Tipo_Mancomuno_Individual Then
            strApellidoPaterno = ""
            strApellidoMaterno = ""
            strNombres = ""
            strRazonSocial = ""
            strClaseParticipe = Codigo_Persona_Mancomuno
            strTipoIdentidad = Codigo_Tipo_Numero_Participe 'Codigo de participe
            strNumIdentidad = lblCodParticipe.Caption
        End If
        
        '*** Actualizar Contrato ***
        With adoComm
            .CommandText = "{ call up_PRManContratoParticipe('" & _
                Trim(lblCodParticipe.Caption) & "','" + lblCodCliente.Caption + "'," & _
                "'" + strClaseParticipe + "','" + strApellidoPaterno + "'," & _
                "'" + strApellidoMaterno + "','" + strNombres + "'," & _
                "'" + strRazonSocial + "','" + Trim(lblDescripParticipe.Caption) + "'," & _
                "'','','','','" & _
                Convertyyyymmdd(gdatFechaActual) & "','" & _
                Trim(txtDireccion1.Text) & "','" & _
                Trim(txtDireccion2.Text) & "','','" & _
                strCodPais & "','" & strCodDepartamento & "','" & _
                strCodProvincia & "','" & strCodDistrito & "','" & _
                "','','','','" & _
                Convertyyyymmdd(CVDate(lblFechaIngreso.Caption)) & "','" & _
                strCodPromotor & "','X','" & _
                strCodSucursal & "','" & strCodAgencia & "','" & _
                strCodDireccionPostal & "','" & _
                Convertyyyymmdd(CVDate(Valor_Fecha)) & "','" & _
                Convertyyyymmdd(CVDate(Valor_Fecha)) & "','" & _
                "','" & Trim(txtContrato.Text) & "','" & _
                "','','','" & _
                gstrLogin & "','" & Convertyyyymmdd(gdatFechaActual) & "','" & _
                gstrLogin & "','" & Convertyyyymmdd(gdatFechaActual) & "','" & _
                "U') }"
            adoConn.Execute .CommandText
                                            
        End With

        Me.MousePointer = vbDefault
        
        MsgBox Mensaje_Edicion_Exitosa, vbExclamation
                                
    End If
    
    frmMainMdi.stbMdi.Panels(3).Text = "Acción"

    cmdOpcion.Visible = True
    With tabContrato
        .TabEnabled(0) = True
        .Tab = 0
    End With
    
    Call Buscar
    
    Exit Sub

CtrlError:
    If adoConn.Errors.Count > 0 Then
        For Each adoError In adoConn.Errors
            strErrMsg = strErrMsg & adoError.Description & " (" & adoError.NativeError & ") " & Chr(13)
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

Private Function TodoOK() As Boolean

    TodoOK = False
    
    If Trim(txtNumDocumentoCliente.Text) = Valor_Caracter Then
        MsgBox "Debe seleccionar el Cliente.", vbCritical
        cmdBusqueda.SetFocus
        Exit Function
    End If
    
    If Trim(txtContrato.Text) = Valor_Caracter Then
        MsgBox "El Campo Número de Contrato no es Válido.", vbCritical
        txtContrato.SetFocus
        Exit Function
    End If
    
    If strEstado = Reg_Adicion Then
        If Not ValidarNumContrato Then Exit Function
    End If
    
    'If cboDireccionPostal.ListIndex = 0 Then
    '    MsgBox "Seleccione el Tipo de Dirección Postal.", vbCritical
    '    cboDireccionPostal.SetFocus
    '    Exit Function
    'End If
    
    'If strCodDireccionPostal = Codigo_Tipo_Direccion_Otro Then
    '    If cboPais.ListIndex = 0 Then
    '        MsgBox "Seleccione el País.", vbCritical
    '        cboPais.SetFocus
    '        Exit Function
    '    End If
        
    '    If cboDepartamento.ListIndex = 0 Then
    '        If cboDepartamento.ListCount > 1 Then
    '            MsgBox "Seleccione el Departamento.", vbCritical
    '            cboDepartamento.SetFocus
    '            Exit Function
    '        End If
    '    End If
        
    '    If cboProvincia.ListIndex = 0 Then
    '        If cboProvincia.ListCount > 1 Then
    '            MsgBox "Seleccione la Provincia.", vbCritical
    '            cboProvincia.SetFocus
    '            Exit Function
    '        End If
    '    End If
        
     '   If cboDistrito.ListIndex = 0 Then
    '        If cboDistrito.ListCount > 1 Then
     '           MsgBox "Seleccione el Distrito.", vbCritical
     '           cboDistrito.SetFocus
     '           Exit Function
     '       End If
     '   End If
    'End If
            
    'If cboSucursal.ListIndex = 0 Then
    '    MsgBox "Seleccione la Sucursal.", vbCritical
    '    cboSucursal.SetFocus
    '    Exit Function
    'End If

    'If cboAgencia.ListIndex = 0 Then
    '    MsgBox "Seleccione la Agencia.", vbCritical
    '    cboAgencia.SetFocus
    '    Exit Function
    'End If
                                                  
    If cboPromotor.ListIndex = 0 Then
        MsgBox "Seleccione el Ejecutivo Comercial.", vbCritical
        cboPromotor.SetFocus
        Exit Function
    End If
    
    '*** Si todo paso OK ***
    TodoOK = True

End Function
Public Sub Modificar()
        
    If Not blnCliente Then
        If strEstado = Reg_Consulta Then
            strEstado = Reg_Edicion
            LlenarFormulario strEstado
            cmdOpcion.Visible = False
            With tabContrato
                .TabEnabled(0) = False
                .Tab = 1
            End With
            'Call Habilita
        End If
    Else
        MsgBox "Solo se pueden modificar los partícipes", vbCritical
        Exit Sub
    End If
        
End Sub

Public Sub Salir()

    Unload Me
    
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
Function ValidarNumContrato() As Boolean

    Dim adoRegistro As ADODB.Recordset
    
    Set adoRegistro = New ADODB.Recordset
    ValidarNumContrato = False
    
    With adoComm
        .CommandText = "SELECT NumContrato FROM ParticipeContrato WHERE NumContrato='" & Trim(txtContrato.Text) & "'"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            tabContrato.Tab = 1
            MsgBox "El Número de Contrato ya existe...", vbCritical, gstrNombreEmpresa
            txtContrato.SetFocus
            txtContrato.SelStart = 0
            txtContrato.SelLength = Len(txtContrato.Text)
            Exit Function
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    '*** Si No existe ***
    ValidarNumContrato = True
    
End Function

Private Sub cboAgencia_Click()

    Dim strSql As String, intRegistro   As Integer
    
    strCodAgencia = Valor_Caracter
    If cboAgencia.ListIndex < 0 Then Exit Sub
    
    strCodAgencia = Trim(arrAgencia(cboAgencia.ListIndex))
    
    'strSQL = "{ call up_ACSelDatosParametro(11,'" & strCodAgencia & "') }"
    'CargarControlLista strSQL, cboPromotor, arrPromotor(), Sel_Defecto
    
    'If cboPromotor.ListCount > 0 Then cboPromotor.ListIndex = 0
    'intRegistro = ObtenerItemLista(arrPromotor(), gstrCodPromotor)
    'If intRegistro >= 0 Then cboPromotor.ListIndex = intRegistro
    
End Sub

Private Sub cboDepartamento_Click()
    
    Dim strSql As String
    
    strCodDepartamento = ""
    If cboDepartamento.ListIndex < 0 Then Exit Sub
    
    strCodDepartamento = Trim(arrDepartamento(cboDepartamento.ListIndex))
    
    strSql = "{ call up_ACSelDatosParametro(2,'" & strCodPais & "','" & strCodDepartamento & "') }"
    CargarControlLista strSql, cboProvincia, arrProvincia(), Sel_Defecto
    
    If cboProvincia.ListCount > -1 Then cboProvincia.ListIndex = 0
    
End Sub

Private Sub cboDireccionPostal_Click()

    Dim adoRegistro As ADODB.Recordset
    Dim intRegistro As Integer
    
    strCodDireccionPostal = ""
    If cboDireccionPostal.ListIndex < 0 Then Exit Sub
    
    strCodDireccionPostal = Trim(arrDireccionPostal(cboDireccionPostal.ListIndex))
    
    Select Case strCodDireccionPostal
        Case Codigo_Tipo_Direccion_Domicilio
            Set adoRegistro = New ADODB.Recordset
            
            If strEstado = Reg_Adicion Then
                adoComm.CommandText = "{ call up_ACSelDatosParametro(7,'" & Trim(lblCodCliente.Caption) & "') }"
            Else
                adoComm.CommandText = "{ call up_ACSelDatosParametro(36,'" & Trim(tdgConsulta.Columns(6)) & "','" & Trim(txtNumDocumento.Text) & "','" & Trim(tdgConsulta.Columns(7)) & "') }"
            End If
            
            Set adoRegistro = adoComm.Execute
            
            If Not adoRegistro.EOF Then
                txtDireccion1.Text = adoRegistro("DireccionCliente1")
                txtDireccion2.Text = adoRegistro("DireccionCliente2")
                
                intRegistro = ObtenerItemLista(arrPais(), adoRegistro("CodPais"))
                If intRegistro >= 0 Then cboPais.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrDepartamento(), adoRegistro("CodDepartamento"))
                If intRegistro >= 0 Then cboDepartamento.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrProvincia(), adoRegistro("CodProvincia"))
                If intRegistro >= 0 Then cboProvincia.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrDistrito(), adoRegistro("CodDistrito"))
                If intRegistro >= 0 Then cboDistrito.ListIndex = intRegistro
            End If
            adoRegistro.Close: Set adoRegistro = Nothing
            
            Call Deshabilita
        
        Case Codigo_Tipo_Direccion_Oficina
            Set adoRegistro = New ADODB.Recordset
            
            adoComm.CommandText = "{ call up_ACSelDatosParametro(7,'" & Trim(lblCodCliente.Caption) & "') }"
            Set adoRegistro = adoComm.Execute
            
            If Not adoRegistro.EOF Then
                txtDireccion1.Text = adoRegistro("DireccionTrabajo1")
                txtDireccion2.Text = adoRegistro("DireccionTrabajo2")
                
                intRegistro = ObtenerItemLista(arrPais(), adoRegistro("CodPaisTrabajo"))
                If intRegistro >= 0 Then cboPais.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrDepartamento(), adoRegistro("CodDepartamentoTrabajo"))
                If intRegistro >= 0 Then cboDepartamento.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrProvincia(), adoRegistro("CodProvinciaTrabajo"))
                If intRegistro >= 0 Then cboProvincia.ListIndex = intRegistro
                
                intRegistro = ObtenerItemLista(arrDistrito(), adoRegistro("CodDistritoTrabajo"))
                If intRegistro >= 0 Then cboDistrito.ListIndex = intRegistro
            End If
            adoRegistro.Close: Set adoRegistro = Nothing
            
            Call Deshabilita
        
        Case Codigo_Tipo_Direccion_Retencion
            Set adoRegistro = New ADODB.Recordset
            
            '*** Cargar dirección de la administradora ***
            With adoComm
                .CommandText = "SELECT DescripDireccion1,DescripDireccion2,CodPais,CodDepartamento,CodProvincia,CodDistrito " & _
                    "FROM Administradora WHERE CodAdministradora='" & gstrCodAdministradora & "'"
                    
                Set adoRegistro = .Execute
                
                If Not adoRegistro.EOF Then
                    txtDireccion1 = Trim(adoRegistro("DescripDireccion1"))
                    txtDireccion2 = Trim(adoRegistro("DescripDireccion2"))
                    
                    intRegistro = ObtenerItemLista(arrPais(), Trim(adoRegistro("CodPais")))
                    If intRegistro >= 0 Then cboPais.ListIndex = intRegistro
                    
                    intRegistro = ObtenerItemLista(arrDepartamento(), Trim(adoRegistro("CodDepartamento")))
                    If intRegistro >= 0 Then cboDepartamento.ListIndex = intRegistro
                    
                    intRegistro = ObtenerItemLista(arrProvincia(), Trim(adoRegistro("CodProvincia")))
                    If intRegistro >= 0 Then cboProvincia.ListIndex = intRegistro
                    
                    intRegistro = ObtenerItemLista(arrDistrito(), Trim(adoRegistro("CodDistrito")))
                    If intRegistro >= 0 Then cboDistrito.ListIndex = intRegistro
                End If
                adoRegistro.Close: Set adoRegistro = Nothing
            End With
        
            Call Deshabilita
        
        Case Codigo_Tipo_Direccion_Otro
            Call Habilita
                                    
            txtDireccion1.Text = ""
            txtDireccion2.Text = ""
            
            cboPais.ListIndex = -1
            If cboPais.ListCount > 0 Then cboPais.ListIndex = 0
            
            txtDireccion1.SetFocus
    End Select
    
End Sub
Private Sub cboDistrito_Click()

    strCodDistrito = ""
    If cboDistrito.ListIndex < 0 Then Exit Sub
    
    strCodDistrito = Trim(arrDistrito(cboDistrito.ListIndex))
    
End Sub
Private Sub cboPais_Click()

    Dim strSql As String
    
    strCodPais = ""
    If cboPais.ListIndex < 0 Then Exit Sub
    
    strCodPais = Trim(arrPais(cboPais.ListIndex))
    
    strSql = "{ call up_ACSelDatosParametro(1,'" & strCodPais & "') }"
    CargarControlLista strSql, cboDepartamento, arrDepartamento(), Sel_Defecto
    
    If cboDepartamento.ListCount > -1 Then cboDepartamento.ListIndex = 0
    
End Sub
Private Sub cboPromotor_Click()

    strCodPromotor = ""
    If cboPromotor.ListIndex < 0 Then Exit Sub
    
    strCodPromotor = Trim(arrPromotor(cboPromotor.ListIndex))
    
End Sub
Private Sub cboProvincia_Click()

    Dim strSql As String
    
    strCodProvincia = ""
    If cboProvincia.ListIndex < 0 Then Exit Sub
    
    strCodProvincia = Trim(arrProvincia(cboProvincia.ListIndex))
    
    strSql = "{ call up_ACSelDatosParametro(3,'" & strCodPais & "','" & strCodDepartamento & "','" & strCodProvincia & "') }"
    CargarControlLista strSql, cboDistrito, arrDistrito(), Sel_Defecto
    
    If cboDistrito.ListCount > -1 Then cboDistrito.ListIndex = 0
    
End Sub

Private Sub cboSucursal_Click()

    Dim strSql As String, intRegistro   As Integer
    
    strCodSucursal = Valor_Caracter
    If cboSucursal.ListIndex < 0 Then Exit Sub
    
    strCodSucursal = Trim(arrSucursal(cboSucursal.ListIndex))
    
    strSql = "{ call up_ACSelDatosParametro(10,'" & strCodSucursal & "') }"
    CargarControlLista strSql, cboAgencia, arrAgencia(), Sel_Defecto
    
    If cboAgencia.ListCount > 0 Then cboAgencia.ListIndex = 0
    
    intRegistro = ObtenerItemLista(arrAgencia(), gstrCodAgencia)
    If intRegistro >= 0 Then cboAgencia.ListIndex = intRegistro
    
End Sub

Private Sub cboTipoContrato_Click()

    Dim intRegistro As Integer

    strTipoContrato = ""
    If cboTipoContrato.ListIndex < 0 Then Exit Sub
    
    strTipoContrato = Trim(arrTipoContrato(cboTipoContrato.ListIndex))

    '*** Tipo de Mancomuno ***
    
    If strTipoContrato = Codigo_Tipo_Contrato_Individual Then
        'cmdMancomuno.Enabled = False
        lblDescrip(17).Visible = False
        cboTipoMancomuno.Visible = False
        
        strCodTipoMancomuno = Codigo_Tipo_Mancomuno_Individual
        cboTipoMancomuno.Clear
    Else
        'cmdMancomuno.Enabled = True
        lblDescrip(17).Visible = True
        cboTipoMancomuno.Visible = True
        
        strSql = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPMAN' AND CodParametro <> '01' ORDER BY DescripParametro"
        CargarControlLista strSql, cboTipoMancomuno, arrTipoMancomuno(), Valor_Caracter
        
        intRegistro = ObtenerItemLista(arrTipoMancomuno(), Codigo_Tipo_Mancomuno_Conjunto)
        If intRegistro >= 0 Then cboTipoMancomuno.ListIndex = intRegistro
    End If

End Sub

Private Sub cboTipoDocumento_Click()

    strCodTipoDocumento = Valor_Caracter
    If cboTipoDocumento.ListIndex < 0 Then Exit Sub
    
    strCodTipoDocumento = Trim(garrTipoDocumento(cboTipoDocumento.ListIndex))
    
End Sub

Private Sub cboTipoMancomuno_Click()

    strCodTipoMancomuno = ""
    If cboTipoMancomuno.ListIndex < 0 Then Exit Sub
    
    strCodTipoMancomuno = Trim(arrTipoMancomuno(cboTipoMancomuno.ListIndex))

'Codigo_Tipo_Mancomuno_Individual = "01"
'Codigo_Tipo_Mancomuno_Indistinto = "03"
'Codigo_Tipo_Mancomuno_Conjunto = "02"

End Sub

Private Sub cmdBuscarComisionista_Click()

    gstrFormulario = Me.Name
    frmBuscarComisionista.Show vbModal
    'frmBusquedarepresentante.strEs
End Sub

Private Sub cmdBusqueda_Click()

    gstrFormulario = Me.Name
    frmBusquedaCliente.Show vbModal
    'frmBusquedarepresentante.strEs
End Sub

Private Sub cmdCuenta_Click()

    frmCuentaParticipe.lblParticipe = Trim(lblDescripParticipe.Caption)
    frmCuentaParticipe.Show
    
End Sub

Private Sub cmdMancomuno_Click()
    
    frmMancomunado.lblParticipe = Trim(lblDescripParticipe.Caption)
    frmMancomunado.strCodTipoMancomuno = strCodTipoMancomuno
    frmMancomunado.Show vbModal
                
End Sub

Private Sub cmdRepresentante_Click()

    frmRepresentanteParticipe.lblParticipe = Trim(lblDescripParticipe.Caption)
    frmRepresentanteParticipe.lblCodClienteParticipe = Trim(lblCodCliente.Caption)
    frmRepresentanteParticipe.Show
    
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
 
    Call ValidarPermisoUsoControl(Trim(gstrLoginUS), Me, Trim(App.Title) + Separador_Codigo_Objeto + _
    gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)
    
    CentrarForm Me

End Sub


Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Lista de Participes por Fecha de Ingreso"
    
    'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Visible = True
    'frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Text = "Contrato de Administración de Cuotas de Participación"
    
End Sub

Private Sub CargarListas()

    Dim strSql  As String, intRegistro As Long
    
    '*** Si/No ***
    strSql = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='RESPSN' ORDER BY DescripParametro"
    CargarControlLista strSql, cboCustodia, arrSiNo(), ""
        
    '*** Tipo Dirección Postal ***
    strSql = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='ENVINF'"
    CargarControlLista strSql, cboDireccionPostal, arrDireccionPostal(), Sel_Defecto
    
    '*** Pais  ***
    strSql = "{ call up_ACSelDatos(13) }"
    CargarControlLista strSql, cboPais, arrPais(), Sel_Defecto
    
    '*** Sucursal ***
    strSql = "{ call up_ACSelDatos(15) }"
    CargarControlLista strSql, cboSucursal, arrSucursal(), Sel_Defecto
    
    '*** Tipo Documento Identidad ***
    strSql = "{ call up_ACSelDatos(11) }"
    CargarControlLista strSql, cboTipoDocumento, garrTipoDocumento(), Sel_Defecto
        
    'strSQL = "{ call up_ACSelDatosParametro(11,'" & strCodAgencia & "') }"
    strSql = "{ call up_ACSelDatos(41) }"
    CargarControlLista strSql, cboPromotor, arrPromotor(), Sel_Defecto
    
    '*** Tipo de Contrato ***
    strSql = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPCON' ORDER BY DescripParametro"
    CargarControlLista strSql, cboTipoContrato, arrTipoContrato(), Valor_Caracter
    
    If cboTipoContrato.ListCount > 0 Then cboTipoContrato.ListIndex = 0
    
    'If cboPromotor.ListCount > 0 Then cboPromotor.ListIndex = 0
    'intRegistro = ObtenerItemLista(arrPromotor(), gstrCodPromotor)
    'If intRegistro >= 0 Then cboPromotor.ListIndex = intRegistro
End Sub
Private Sub InicializarValores()

    '*** Valores Iniciales ***
    strEstado = Reg_Defecto
    blnCliente = False
    tabContrato.Tab = 0
    tabContrato.TabEnabled(1) = False
    optParticipe(0).Value = vbChecked
    
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgConsulta.Columns(0).Width = tdgConsulta.Width * 0.01 * 15
    tdgConsulta.Columns(1).Width = tdgConsulta.Width * 0.01 * 36
    tdgConsulta.Columns(2).Width = tdgConsulta.Width * 0.01 * 13
    tdgConsulta.Columns(3).Width = tdgConsulta.Width * 0.01 * 13
    tdgConsulta.Columns(4).Width = tdgConsulta.Width * 0.01 * 10
    
    '*** Verificando Nivel de Acceso de Usuario ***
'    strNivAcceso = AccesoForm(gstrNomOpc, gstrNumInd)

    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmContratoParticipe = Nothing
    Call OcultarReportes
    frmMainMdi.stbMdi.Panels(3).Text = "Acción..."
    
End Sub

Private Sub optParticipe_Click(Index As Integer)

    If Index = 0 Then
        txtCodParticipe.Enabled = True
        txtNumDocumento.Enabled = False
        txtDescripcion.Enabled = False
        txtCodParticipe.Text = ""
    ElseIf Index = 1 Then
        txtCodParticipe.Enabled = False
        txtNumDocumento.Enabled = True
        txtDescripcion.Enabled = False
        txtNumDocumento.Text = ""
    Else
        txtCodParticipe.Enabled = False
        txtNumDocumento.Enabled = False
        txtDescripcion.Enabled = True
        txtDescripcion.Text = ""
    End If
    
    Call Buscar
    
End Sub

Private Sub tabContrato_Click(PreviousTab As Integer)

    Select Case tabContrato.Tab
        Case 1, 2
            If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vModify)
            If strEstado = Reg_Defecto Then tabContrato.Tab = 0
                                
    End Select

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

Private Sub txtCodParticipe_LostFocus()

    txtCodParticipe.Text = Format(txtCodParticipe.Text, "00000000000000000000")
    
End Sub


Private Sub txtContrato_LostFocus()

   txtContrato.Text = Format(txtContrato.Text, "000000000000000")
    
    If strEstado = Reg_Adicion Then
        If Not ValidarNumContrato() Then Exit Sub
    End If
    
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub

Private Sub txtDireccion1_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub

Private Sub txtNumDocumentoCliente_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Call ObtenerDatosCliente
    End If
    
End Sub
