VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{830D5A06-9B70-4F8C-98B6-7A19C4E7760B}#1.0#0"; "TAMControls.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmSolicitudParticipe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Solicitud de Operación"
   ClientHeight    =   9450
   ClientLeft      =   1425
   ClientTop       =   1605
   ClientWidth     =   14595
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9450
   ScaleWidth      =   14595
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   12360
      TabIndex        =   10
      Top             =   8600
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
      Left            =   810
      TabIndex        =   9
      Top             =   8600
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
   Begin TabDlg.SSTab tabSolicitud 
      Height          =   8385
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   14565
      _ExtentX        =   25691
      _ExtentY        =   14790
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      ForeColor       =   8388608
      TabCaption(0)   =   "Lista"
      TabPicture(0)   =   "frmSolicitudParticipe.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraCriterios"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraSolicitud"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Solicitud"
      TabPicture(1)   =   "frmSolicitudParticipe.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraDatosParticipe"
      Tab(1).Control(1)=   "cmdAccion"
      Tab(1).Control(2)=   "fraTipoAporte"
      Tab(1).Control(3)=   "fraFormaPago"
      Tab(1).Control(4)=   "fraDatosGenerales"
      Tab(1).Control(5)=   "fraDatosSolicitud"
      Tab(1).Control(6)=   "fraDatosSolicitudDesconocida"
      Tab(1).Control(7)=   "Label2"
      Tab(1).ControlCount=   8
      Begin VB.Frame fraDatosParticipe 
         Caption         =   "Datos Contrato"
         Height          =   2085
         Left            =   -74640
         TabIndex        =   49
         Top             =   2100
         Width           =   13755
         Begin VB.ComboBox cboComisionista 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   104
            Top             =   1500
            Width           =   6045
         End
         Begin VB.TextBox txtTitularSolicitante 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2040
            TabIndex        =   18
            Top             =   720
            Width           =   6015
         End
         Begin VB.TextBox txtEjecutivoComercial 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2040
            TabIndex        =   38
            Top             =   2250
            Visible         =   0   'False
            Width           =   4365
         End
         Begin VB.TextBox txtTipoDocumento 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2040
            TabIndex        =   19
            Top             =   1110
            Width           =   3675
         End
         Begin VB.ComboBox cboCertificado 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   1860
            Width           =   6045
         End
         Begin VB.TextBox txtNumDocumento 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5790
            TabIndex        =   20
            Top             =   1105
            Width           =   2265
         End
         Begin VB.CommandButton cmdBusqueda 
            Caption         =   "..."
            Height          =   315
            Left            =   13050
            TabIndex        =   17
            ToolTipText     =   "Búsqueda de Partícipe"
            Top             =   315
            Width           =   315
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Titular/Solicitante"
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   38
            Left            =   360
            TabIndex        =   99
            Top             =   780
            Width           =   1635
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Certificado"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   5
            Left            =   360
            TabIndex        =   98
            Top             =   1920
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lblComisionista 
            Caption         =   "Comisionista"
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   8
            Left            =   360
            TabIndex        =   85
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Partícipe"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   17
            Left            =   360
            TabIndex        =   73
            Top             =   345
            Width           =   1455
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Tipo Contrato"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   16
            Left            =   8340
            TabIndex        =   72
            Top             =   750
            Width           =   1275
         End
         Begin VB.Label lblDescripTipoParticipe 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   10140
            TabIndex        =   71
            Top             =   720
            Width           =   2805
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Cuotas Bloquedas"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   29
            Left            =   8340
            TabIndex        =   55
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label lblCuotasBloqueadas 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   10140
            TabIndex        =   54
            Top             =   1500
            Width           =   2820
         End
         Begin VB.Label lblCuotasDisponibles 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   10140
            TabIndex        =   53
            Top             =   1105
            Width           =   2805
         End
         Begin VB.Label lblDescripParticipe 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2040
            TabIndex        =   52
            Top             =   330
            Width           =   10905
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Tipo/Num Doc ID."
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   27
            Left            =   360
            TabIndex        =   51
            Top             =   1155
            Width           =   1575
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Cuotas Disponibles"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   9
            Left            =   8340
            TabIndex        =   50
            Top             =   1185
            Width           =   1695
         End
      End
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   -64410
         TabIndex        =   39
         Top             =   7530
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
      Begin VB.Frame fraTipoAporte 
         Caption         =   "Tipo Aporte"
         Height          =   735
         Left            =   -74640
         TabIndex        =   100
         Top             =   4260
         Width           =   13755
         Begin VB.ComboBox cboActivoAporte 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   7440
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   270
            Width           =   5625
         End
         Begin VB.ComboBox cboTipoAporte 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   240
            Width           =   2220
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Activo"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   40
            Left            =   6240
            TabIndex        =   102
            Top             =   270
            Width           =   975
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Tipo de Aporte"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   39
            Left            =   360
            TabIndex        =   101
            Top             =   270
            Width           =   1575
         End
      End
      Begin VB.Frame fraFormaPago 
         Caption         =   "Forma de Pago"
         Height          =   1155
         Left            =   -74640
         TabIndex        =   65
         Top             =   6240
         Width           =   13755
         Begin VB.ComboBox cboCuenta 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   5640
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   250
            Visible         =   0   'False
            Width           =   7380
         End
         Begin VB.CheckBox chkPagoParcial 
            Caption         =   "Pago Parcial"
            Height          =   195
            Left            =   11040
            TabIndex        =   35
            Top             =   750
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.ComboBox cboCuentaFondo 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2490
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   720
            Width           =   8370
         End
         Begin VB.ComboBox cboTipoFormaPago 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   270
            Width           =   2220
         End
         Begin VB.ComboBox cboBanco 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   5640
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   250
            Visible         =   0   'False
            Width           =   2370
         End
         Begin VB.ComboBox cboNumCuenta 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   10215
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   250
            Visible         =   0   'False
            Width           =   2220
         End
         Begin VB.TextBox txtNumCheque 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   10215
            MaxLength       =   15
            TabIndex        =   32
            Top             =   250
            Visible         =   0   'False
            Width           =   2220
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   3
            X1              =   330
            X2              =   13200
            Y1              =   630
            Y2              =   630
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Cuenta Fondo Destino"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   20
            Left            =   420
            TabIndex        =   69
            Top             =   800
            Width           =   1935
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Tipo"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   28
            Left            =   420
            TabIndex        =   68
            Top             =   300
            Width           =   735
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Núm.Cheque"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   19
            Left            =   8325
            TabIndex        =   67
            Top             =   250
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Cuenta"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   18
            Left            =   4590
            TabIndex        =   66
            Top             =   300
            Visible         =   0   'False
            Width           =   735
         End
      End
      Begin VB.Frame fraDatosGenerales 
         Caption         =   "Datos Generales"
         Height          =   1515
         Left            =   -74640
         TabIndex        =   43
         Top             =   420
         Width           =   13755
         Begin VB.Timer tmrHora 
            Enabled         =   0   'False
            Interval        =   60000
            Left            =   6810
            Top             =   1770
         End
         Begin VB.TextBox txtNumPapeleta 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   1000
            Width           =   4425
         End
         Begin VB.ComboBox cboTipoOperacion 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   300
            Width           =   6045
         End
         Begin VB.ComboBox cboFondo 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   650
            Width           =   6045
         End
         Begin VB.ComboBox cboEjecutivo 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   10140
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   1000
            Visible         =   0   'False
            Width           =   2805
         End
         Begin MSComCtl2.DTPicker dtpHoraSolicitud 
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
            Left            =   11970
            TabIndex        =   14
            Top             =   300
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
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
            CustomFormat    =   "HH:mm"
            Format          =   51511299
            UpDown          =   -1  'True
            CurrentDate     =   38831
         End
         Begin MSComCtl2.DTPicker dtpFechaValorCuota 
            Height          =   315
            Left            =   10140
            TabIndex        =   15
            Top             =   650
            Width           =   1335
            _ExtentX        =   2355
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
            Format          =   51511297
            CurrentDate     =   38068
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Operador"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   42
            Left            =   8280
            TabIndex        =   103
            Top             =   1050
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Valor de Redención"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   31
            Left            =   8280
            TabIndex        =   76
            Top             =   700
            Width           =   1695
         End
         Begin VB.Label lblFechaSolicitud 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "dd/mm/yyyy"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   10140
            TabIndex        =   70
            Top             =   300
            Width           =   1305
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Operación"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   26
            Left            =   360
            TabIndex        =   48
            Top             =   350
            Width           =   975
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Num.Papeleta"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   3
            Left            =   360
            TabIndex        =   47
            Top             =   1050
            Width           =   1215
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Fondo"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   4
            Left            =   360
            TabIndex        =   46
            Top             =   700
            Width           =   975
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Solicitud"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   6
            Left            =   8280
            TabIndex        =   45
            Top             =   350
            Width           =   1335
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Hora"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   7
            Left            =   11490
            TabIndex        =   44
            Top             =   350
            Width           =   420
         End
      End
      Begin VB.Frame fraSolicitud 
         Height          =   5145
         Left            =   360
         TabIndex        =   42
         Top             =   3030
         Width           =   13815
         Begin VB.ListBox lstLeyenda 
            Height          =   255
            Left            =   9120
            TabIndex        =   34
            Top             =   240
            Visible         =   0   'False
            Width           =   1200
         End
         Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
            Bindings        =   "frmSolicitudParticipe.frx":0038
            Height          =   4035
            Left            =   360
            OleObjectBlob   =   "frmSolicitudParticipe.frx":0052
            TabIndex        =   8
            Top             =   750
            Width           =   13005
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Suscripciones (0)"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   21
            Left            =   390
            TabIndex        =   75
            Top             =   300
            Width           =   1815
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Rescates (0)"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   30
            Left            =   2550
            TabIndex        =   74
            Top             =   300
            Width           =   1815
         End
      End
      Begin VB.Frame fraCriterios 
         Caption         =   "Criterios de búsqueda"
         Height          =   2310
         Left            =   360
         TabIndex        =   77
         Top             =   600
         Width           =   13815
         Begin VB.ComboBox cboPromotor 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   9300
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1710
            Width           =   4215
         End
         Begin VB.ComboBox cboAgencia 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4830
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   1710
            Width           =   4185
         End
         Begin VB.ComboBox cboSucursal 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   360
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   1710
            Width           =   4185
         End
         Begin VB.ComboBox cboFondoSolicitud 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   450
            Width           =   6255
         End
         Begin VB.ComboBox cboTipoSolicitud 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   885
            Width           =   6255
         End
         Begin MSComCtl2.DTPicker dtpFechaDesde 
            Height          =   315
            Left            =   10080
            TabIndex        =   3
            Top             =   450
            Width           =   1545
            _ExtentX        =   2725
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
            Format          =   51511297
            CurrentDate     =   38068
         End
         Begin MSComCtl2.DTPicker dtpFechaHasta 
            Height          =   315
            Left            =   10080
            TabIndex        =   4
            Top             =   885
            Width           =   1545
            _ExtentX        =   2725
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
            Format          =   51511297
            CurrentDate     =   38068
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Hasta"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   25
            Left            =   9330
            TabIndex        =   84
            Top             =   900
            Width           =   615
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Desde"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   24
            Left            =   9330
            TabIndex        =   83
            Top             =   465
            Width           =   615
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Fondo"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   22
            Left            =   360
            TabIndex        =   82
            Top             =   465
            Width           =   1335
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Sucursal"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   360
            TabIndex        =   81
            Top             =   1425
            Width           =   1215
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Agencia"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   1
            Left            =   4860
            TabIndex        =   80
            Top             =   1425
            Width           =   1215
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Operador"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   2
            Left            =   9300
            TabIndex        =   79
            Top             =   1425
            Width           =   1095
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Tipo Operación"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   23
            Left            =   360
            TabIndex        =   78
            Top             =   900
            Width           =   1335
         End
      End
      Begin VB.Frame fraDatosSolicitud 
         Caption         =   "Datos Solicitud Conocida"
         Height          =   1155
         Left            =   -74640
         TabIndex        =   56
         Top             =   5040
         Width           =   13755
         Begin VB.TextBox txtMontoIgv 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5640
            TabIndex        =   27
            Top             =   675
            Width           =   2370
         End
         Begin VB.TextBox txtMontoNetoSolicitud 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   10140
            TabIndex        =   28
            Top             =   675
            Width           =   2250
         End
         Begin TAMControls.TAMTextBox txtValorCuota 
            Height          =   315
            Left            =   10140
            TabIndex        =   26
            Top             =   300
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
            Container       =   "frmSolicitudParticipe.frx":7082
            Text            =   "0.00000"
            Decimales       =   5
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            AceptaNegativos =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   5
            MaximoValor     =   999999999
         End
         Begin TAMControls.TAMTextBox txtCantCuotasSolicitud 
            Height          =   315
            Left            =   2040
            TabIndex        =   24
            Top             =   300
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
            Container       =   "frmSolicitudParticipe.frx":709E
            Text            =   "0.0000"
            Decimales       =   4
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            AceptaNegativos =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   4
            MaximoValor     =   999999999
         End
         Begin TAMControls.TAMTextBox txtMontoComision 
            Height          =   315
            Left            =   5640
            TabIndex        =   25
            Top             =   300
            Width           =   2385
            _ExtentX        =   4207
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
            Container       =   "frmSolicitudParticipe.frx":70BA
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
         Begin VB.Label lblValorCuota 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6450
            TabIndex        =   64
            Top             =   1590
            Visible         =   0   'False
            Width           =   2220
         End
         Begin VB.Label lblMontoSolicitud 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2040
            TabIndex        =   63
            Top             =   705
            Width           =   2220
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Cuotas"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   10
            Left            =   480
            TabIndex        =   62
            Top             =   315
            Width           =   855
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Monto"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   11
            Left            =   480
            TabIndex        =   61
            Top             =   720
            Width           =   855
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Valor Cuota"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   12
            Left            =   8325
            TabIndex        =   60
            Top             =   315
            Width           =   1095
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Monto Neto"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   13
            Left            =   8325
            TabIndex        =   59
            Top             =   705
            Width           =   1095
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Comisión"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   14
            Left            =   4590
            TabIndex        =   58
            Top             =   315
            Width           =   855
         End
         Begin VB.Label lblDescrip 
            Caption         =   "I.G.V."
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   15
            Left            =   4590
            TabIndex        =   57
            Top             =   735
            Width           =   855
         End
      End
      Begin VB.Frame fraDatosSolicitudDesconocida 
         Caption         =   "Datos Solicitud Desconocida"
         Height          =   1155
         Left            =   -74640
         TabIndex        =   86
         Top             =   5040
         Visible         =   0   'False
         Width           =   13755
         Begin TAMControls.TAMTextBox txtCantCuotasSolicitudD 
            Height          =   315
            Left            =   2040
            TabIndex        =   36
            Top             =   420
            Width           =   2235
            _ExtentX        =   3942
            _ExtentY        =   556
            BackColor       =   -2147483633
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
            Container       =   "frmSolicitudParticipe.frx":70D6
            Text            =   "0.00000"
            Decimales       =   5
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            AceptaNegativos =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   5
            MaximoValor     =   999999999
         End
         Begin TAMControls.TAMTextBox txtMontoSolicitud 
            Height          =   315
            Left            =   2040
            TabIndex        =   37
            Top             =   810
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
            Container       =   "frmSolicitudParticipe.frx":70F2
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
         Begin VB.Label lblMontoNetoSolicitud 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   10260
            TabIndex        =   41
            Top             =   810
            Width           =   2220
         End
         Begin VB.Label lblValorCuotaD 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   10260
            TabIndex        =   96
            Top             =   420
            Width           =   2220
         End
         Begin VB.Label lblMontoIgv 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   5760
            TabIndex        =   95
            Top             =   810
            Width           =   2220
         End
         Begin VB.Label lblDescrip 
            Caption         =   "I.G.V."
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   37
            Left            =   4560
            TabIndex        =   94
            Top             =   855
            Width           =   855
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Comisión"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   36
            Left            =   4560
            TabIndex        =   93
            Top             =   435
            Width           =   855
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Monto Neto"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   35
            Left            =   8325
            TabIndex        =   92
            Top             =   825
            Width           =   1095
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Valor Cuota"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   34
            Left            =   8325
            TabIndex        =   91
            Top             =   435
            Width           =   1095
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Monto"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   33
            Left            =   480
            TabIndex        =   90
            Top             =   840
            Width           =   855
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Cuotas"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   32
            Left            =   480
            TabIndex        =   89
            Top             =   435
            Width           =   855
         End
         Begin VB.Label lblMontoComision 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   5760
            TabIndex        =   88
            Top             =   420
            Width           =   2220
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6450
            TabIndex        =   87
            Top             =   1590
            Visible         =   0   'False
            Width           =   2220
         End
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   285
         Left            =   -72210
         TabIndex        =   97
         Top             =   8400
         Visible         =   0   'False
         Width           =   2955
      End
   End
End
Attribute VB_Name = "frmSolicitudParticipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrFondoSolicitud()                 As String, arrTipoSolicitud()               As String
Dim arrSucursal()                       As String, arrAgencia()                     As String
Dim arrPromotor()                       As String, arrTipoOperacion()               As String
Dim arrFondo()                          As String, arrEjecutivo()                   As String
Dim arrTipoFormaPago()                  As String, arrCuenta()                      As String
Dim arrCuentaFondo()                    As String, arrNumCuenta()                   As String
Dim arrBanco()                          As String, arrComisionista()                As String
Dim arrLeyendaSuscripcion()             As String, arrLeyendaRescate()              As String
Dim arrTipoAporte()                     As String, arrActivoAporte()                As String

Dim strTipoAporte                       As String, strCodActivoAporte               As String
Dim strCodFondoSolicitud                As String, strCodTipoSolicitud              As String
Dim strCodSucursal                      As String, strCodAgencia                    As String
Dim strCodSucursalSolicitud             As String, strCodAgenciaSolicitud           As String
Dim strCodSucursalDestino               As String, strCodAgenciaDestino             As String
Dim strCodPromotor                      As String, strCodTipoOperacion              As String
Dim strCodClaseOperacion                As String, strCodTipoValuacion              As String
Dim strCodFondo                         As String, strCodEjecutivo                  As String
Dim strCodEjecutivoDestino              As String, strCodBancoDestino               As String
Dim strCodTipoDocumento                 As String, strCodTipoFormaPago              As String
Dim strNumCuentaDestino                 As String, strNumCuenta                     As String
Dim strCodBanco                         As String, strCodMonedaFondo                As String
Dim strCodComision                      As String, strIndPagoParcial                As String
Dim strEstado                           As String, strHoraCorte                     As String
Dim strSql                              As String, strNumPapeleta                   As String
Dim strTipoCuenta                       As String, strTipoCuentaDestino             As String
Dim indExisteCuentaParticipe            As String, strCodComisionista               As String
Dim numSecCondicion                     As Integer

Dim datFechaEtapaOperativa              As Date, datFechaEtapaPreOperativa          As Date

Dim dblTasaSuscripcion                  As Double, dblTasaRescate                   As Double
Dim dblCantCuotaMinSuscripcionInicial   As Double, dblMontoMinSuscripcionInicial    As Double
Dim dblCantMinCuotaSuscripcion          As Double, dblMontoMinSuscripcion           As Double
Dim dblPorcenMaxParticipe               As Double, dblValorCuota                    As Double
Dim dblCantCuotaInicio                  As Double, dblCantMaxCuotaFondo             As Double
Dim dblPorcenPago                       As Double
Dim curMontoEmitido                     As Currency, curMontoPago                   As Currency

Dim blnCuota                            As Boolean, blnMonto                        As Boolean
Dim blnValorConocido                    As Boolean
Dim arrCertificado()                    As String, strNumCertificado                As String
Dim strFechaProceso                     As String
Dim strTDBCodParticipe                  As String, strTDBNumCertificado             As String
Dim strCodCliente                       As String
Dim adoConsulta                         As ADODB.Recordset
Dim indSortAsc                          As Boolean, indSortDesc                     As Boolean

Private Sub ObtenerFechaValorCuota()

    Dim adoRegistro     As ADODB.Recordset, adoRegistroTmp  As ADODB.Recordset
    Dim datFechaInicial As Date, datFechaFinal              As Date
    Dim intmes          As Integer, intPeriodo              As Integer
    Dim intCantDias     As Integer, datFechaTemporal        As Date
    
    Set adoRegistro = New ADODB.Recordset
    With adoComm
        .CommandText = "SELECT FechaInicioEtapaPreOperativa, FrecuenciaValorizacion, IndTipoValuacionModificable FROM Fondo WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            intmes = Month(CVDate(lblFechaSolicitud.Caption))
            intPeriodo = Year(CVDate(lblFechaSolicitud.Caption))
            
            If Trim(adoRegistro("IndTipoValuacionModificable")) = Valor_Indicador Then
                dtpFechaValorCuota.Enabled = True
            Else
                dtpFechaValorCuota.Value = False
            End If
            
            If Trim(adoRegistro("FrecuenciaValorizacion")) = Codigo_Tipo_Frecuencia_Diaria Then
                dtpFechaValorCuota.Value = CVDate(lblFechaSolicitud.Caption)
            End If
            
            Set adoRegistroTmp = New ADODB.Recordset
            .CommandText = "SELECT ValorParametro FROM AuxiliarParametro " & _
                "WHERE CodParametro='" & Trim(adoRegistro("FrecuenciaValorizacion")) & "' AND CodTipoParametro='TIPFRE'"
            Set adoRegistroTmp = .Execute
            If Not adoRegistroTmp.EOF Then
                intCantDias = CInt(adoRegistroTmp("ValorParametro"))
            End If
            adoRegistroTmp.Close: Set adoRegistroTmp = Nothing
            
            If Trim(adoRegistro("FrecuenciaValorizacion")) = Codigo_Tipo_Frecuencia_Mensual Then
                datFechaTemporal = DateAdd("m", -1, CVDate(lblFechaSolicitud.Caption))
                intmes = Month(datFechaTemporal)
                intPeriodo = Year(datFechaTemporal)
                datFechaInicial = DateAdd("d", 1, UltimaFechaMes(intmes, intPeriodo))
                intmes = Month(datFechaInicial)
                intPeriodo = Year(datFechaInicial)
                datFechaFinal = UltimaFechaMes(intmes, intPeriodo)
                
                If datFechaInicial > CVDate(adoRegistro("FechaInicioEtapaPreOperativa")) Then
                    dtpFechaValorCuota.Value = datFechaInicial
                Else
                    dtpFechaValorCuota.Value = CVDate(adoRegistro("FechaInicioEtapaPreOperativa"))
                End If
                dtpFechaValorCuota_Change
            End If
            
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With

End Sub

Public Sub SubImprimir(Index As Integer)

    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    Dim strFechaDesde           As String, strFechaHasta        As String
    Dim proceder As Integer
    proceder = 1
    
    If tabSolicitud.Tab = 1 Then Exit Sub
    
    Select Case Index
        Case 1
            'If adoConsulta.RecordCount > 0 Then
                gstrNameRepo = "ParticipeSolicitud"
                            
                Set frmReporte = New frmVisorReporte
    
                ReDim aReportParamS(7)
                ReDim aReportParamFn(5)
                ReDim aReportParamF(5)
    
                strFechaDesde = Convertyyyymmdd(dtpFechaDesde.Value)
                strFechaHasta = Convertyyyymmdd(DateAdd("d", 1, dtpFechaHasta.Value))
                
                aReportParamFn(0) = "Usuario"
                aReportParamFn(1) = "Hora"
                aReportParamFn(2) = "NombreEmpresa"
                aReportParamFn(3) = "Fondo"
                aReportParamFn(4) = "FechaDesde"
                aReportParamFn(5) = "FechaHasta"
                
                aReportParamF(0) = gstrLogin
                aReportParamF(1) = Format(Time(), "hh:mm:ss")
                aReportParamF(2) = gstrNombreEmpresa & Space(1)
                aReportParamF(3) = Trim(cboFondoSolicitud.Text)
                aReportParamF(4) = CStr(dtpFechaDesde.Value)
                aReportParamF(5) = CStr(dtpFechaHasta.Value)
                            
                aReportParamS(0) = strCodFondoSolicitud
                aReportParamS(1) = gstrCodAdministradora
                aReportParamS(2) = strCodTipoSolicitud
                aReportParamS(3) = strFechaDesde
                aReportParamS(4) = strFechaHasta
                aReportParamS(5) = strCodSucursal
                aReportParamS(6) = strCodAgencia
                aReportParamS(7) = strCodPromotor
                
                proceder = 1
'            Else
'                MsgBox "No existe información para los criterios seleccionados", vbCritical
'                proceder = 0
'            End If
        Case 2
            'MsgBox adoConsulta.Fields("NumSolicitud"), vbCritical
            If Not adoConsulta.EOF Then
                'If adoConsulta.Fields("NumSolicitud") <> Null Then
                    gstrNameRepo = "SolicitudSuscripcion"
                                
                    Set frmReporte = New frmVisorReporte
        
                    ReDim aReportParamS(3)
                    ReDim aReportParamFn(3)
                    ReDim aReportParamF(3)
        
                    aReportParamFn(0) = "Usuario"
                    aReportParamFn(1) = "Hora"
                    aReportParamFn(2) = "NombreEmpresa"
                    aReportParamFn(3) = "Fondo"
                    
                    aReportParamF(0) = gstrLogin
                    aReportParamF(1) = Format(Time(), "hh:mm:ss")
                    aReportParamF(2) = gstrNombreEmpresa & Space(1)
                    aReportParamF(3) = Trim(cboFondoSolicitud.Text)
                
                    aReportParamS(0) = strCodFondoSolicitud
                    aReportParamS(1) = gstrCodAdministradora
                    aReportParamS(2) = adoConsulta.Fields("CodParticipe")
                    aReportParamS(3) = adoConsulta.Fields("NumSolicitud")
                'End If
            Else
                MsgBox "Debe Seleccionar una Solicitud de Suscripción para ver el Reporte", vbCritical
                proceder = 0
            End If
            
        Case 3
            'MsgBox adoConsulta.Fields("NumSolicitud"), vbCritical
            If Not adoConsulta.EOF Then
                'If adoConsulta.Fields("NumSolicitud") = Null Then
                    gstrNameRepo = "SolicitudRescate"
                                
                    Set frmReporte = New frmVisorReporte
        
                    ReDim aReportParamS(3)
                    ReDim aReportParamFn(3)
                    ReDim aReportParamF(3)
        
                    aReportParamFn(0) = "Usuario"
                    aReportParamFn(1) = "Hora"
                    aReportParamFn(2) = "NombreEmpresa"
                    aReportParamFn(3) = "Fondo"
                    
                    aReportParamF(0) = gstrLogin
                    aReportParamF(1) = Format(Time(), "hh:mm:ss")
                    aReportParamF(2) = gstrNombreEmpresa & Space(1)
                    aReportParamF(3) = Trim(cboFondoSolicitud.Text)
                
                    aReportParamS(0) = strCodFondoSolicitud
                    aReportParamS(1) = gstrCodAdministradora
                    aReportParamS(2) = adoConsulta.Fields("CodParticipe")
                    aReportParamS(3) = adoConsulta.Fields("NumSolicitud")
                'End If
            Else
                MsgBox "Debe Seleccionar una Solicitud de Rescate para ver el Reporte", vbCritical
                proceder = 0
            End If
    End Select
    
    If proceder = 1 Then
        gstrSelFrml = ""
        frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"
    
        Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())
    
        frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
        frmReporte.Show vbModal
    
        Set frmReporte = Nothing
    
        Screen.MousePointer = vbNormal
    End If
    
End Sub
Public Sub Adicionar()

    frmMainMdi.stbMdi.Panels(3).Text = "Adicionar Solicitud..."
    
    strEstado = Reg_Adicion
    LlenarFormulario strEstado
    cmdOpcion.Visible = False
    With tabSolicitud
        .TabEnabled(0) = False
        .Tab = 1
    End With
    Call Deshabilita
    
End Sub

Private Sub LlenarFormulario(strModo As String)

    Dim strSql As String
    Dim intRegistro As Integer
    Dim adoRegistro As ADODB.Recordset
    
    Select Case strModo
        Case Reg_Adicion
            cboTipoOperacion.ListIndex = -1
            intRegistro = ObtenerItemLista(arrTipoOperacion(), Codigo_Operacion_Suscripcion)
            If intRegistro >= 0 Then cboTipoOperacion.ListIndex = intRegistro
            
            cboFondo.ListIndex = -1
            If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
                        
            txtNumPapeleta.Text = Valor_Caracter
            lblFechaSolicitud.Caption = CStr(gdatFechaActual)
            dtpFechaValorCuota.Value = gdatFechaActual
            dtpFechaValorCuota.Enabled = False
            dtpHoraSolicitud.Value = ObtenerHoraServidor
            
            cboEjecutivo.ListIndex = -1
            intRegistro = ObtenerItemLista(arrEjecutivo(), gstrCodPromotor)
            If intRegistro >= 0 Then cboEjecutivo.ListIndex = intRegistro
            
'            cboTipoDocumento.ListIndex = -1
'            If cboTipoDocumento.ListCount > 0 Then cboTipoDocumento.ListIndex = 0
            cboCertificado.ListIndex = -1
                        
            txtNumDocumento.Text = Valor_Caracter
            lblDescripTipoParticipe.Caption = Valor_Caracter
            lblDescripParticipe.Caption = Valor_Caracter
            txtTitularSolicitante.Text = Valor_Caracter
            txtTipoDocumento.Text = Valor_Caracter
            lblCuotasDisponibles.Caption = "0"
            lblCuotasBloqueadas.Caption = "0"
            
            txtCantCuotasSolicitud.Text = "0"
            txtMontoNetoSolicitud.Text = "0"
            txtMontoComision.Text = "0"
            txtMontoIgv.Text = "0"
            Call ColorControlDeshabilitado(txtMontoComision)
            Call ColorControlDeshabilitado(txtMontoIgv)
            'lblValorCuota.Caption = "0"
            txtValorCuota.Text = "0"
            
            
            adoComm.CommandText = "{ call up_ACSelDatosParametro(64,'" & strCodFondo & "') }"
            Set adoRegistro = adoComm.Execute

            If Not adoRegistro.EOF Then
                If IsNull(Mid(adoRegistro("NumPapeleta"), 1, 15)) Then
                    txtNumPapeleta.Text = Format(1, "000000000000000")
                Else
                    'txtNumPapeleta.Text = "P" & Format(CInt(Mid(adoRegistro("NumPapeleta"), 2, Len(adoRegistro("NumPapeleta")))) + 1, "00000000000000")
                    'txtNumPapeleta.Text = Format(Val(Mid(adoRegistro("NumPapeleta"), 2, Len(adoRegistro("NumPapeleta")))) + 1, "0000000000000")
                    txtNumPapeleta.Text = adoRegistro("NumPapeleta")
                End If
            Else
                txtNumPapeleta.Text = Format(1, "000000000000000")
            End If
            
            adoRegistro.Close: Set adoRegistro = Nothing
            
            
            lblMontoSolicitud.Caption = "0"
            
            cboTipoFormaPago.ListIndex = -1
            If cboTipoFormaPago.ListCount > 0 Then cboTipoFormaPago.ListIndex = 0
            intRegistro = ObtenerItemLista(arrTipoFormaPago(), Codigo_FormaPago_Efectivo)
            If intRegistro >= 0 Then cboTipoFormaPago.ListIndex = intRegistro
            
            cboCuentaFondo.ListIndex = -1
            If cboCuentaFondo.ListCount > 0 Then cboCuentaFondo.ListIndex = 0
            
            cboNumCuenta.ListIndex = -1
            If cboNumCuenta.ListCount > 0 Then cboNumCuenta.ListIndex = 0
                                    
            cboBanco.ListIndex = -1
            If cboBanco.ListCount > 0 Then cboBanco.ListIndex = 0
                        
            txtNumCheque.Text = Valor_Caracter
            chkPagoParcial.Value = vbUnchecked
            chkPagoParcial.Enabled = False
            
            strSql = "{ call up_ACSelDatos(41) }"
            CargarControlLista strSql, cboEjecutivo, arrEjecutivo(), Sel_Defecto
    
            If cboEjecutivo.ListCount > 0 Then cboEjecutivo.ListIndex = 0
            intRegistro = ObtenerItemLista(arrEjecutivo(), strCodEjecutivo)
            If intRegistro >= 0 Then cboEjecutivo.ListIndex = intRegistro
             
            cboTipoOperacion.SetFocus
    End Select
    
    tmrHora.Enabled = True
    
End Sub
Public Sub Buscar()

    Dim strFechaDesde       As String, strFechaHasta        As String
    Dim datFechaSiguiente   As Date
    Dim strSql              As String
                                                                                    
    Me.MousePointer = vbHourglass
    
    strFechaDesde = Convertyyyymmdd(dtpFechaDesde.Value)
    datFechaSiguiente = DateAdd("d", 1, dtpFechaHasta.Value)
    strFechaHasta = Convertyyyymmdd(datFechaSiguiente)
                
    If cboTipoSolicitud.ListIndex > -1 And cboSucursal.ListIndex > 0 And cboAgencia.ListIndex > 0 And cboPromotor.ListIndex > 0 Then
    
        strSql = "{ call up_ACSelDatosParametro(34,'" & strCodFondoSolicitud & "','" & _
            gstrCodAdministradora & "','" & strFechaDesde & "','" & strFechaHasta & "','" & _
            strCodTipoSolicitud & "','" & strCodSucursal & "','" & strCodAgencia & "','" & _
            strCodPromotor & "') }"
            
    ElseIf cboTipoSolicitud.ListIndex > -1 And cboSucursal.ListIndex > 0 And cboAgencia.ListIndex > 0 Then
    
        strSql = "{ call up_ACSelDatosParametro(33,'" & strCodFondoSolicitud & "','" & _
            gstrCodAdministradora & "','" & strFechaDesde & "','" & strFechaHasta & "','" & _
            strCodTipoSolicitud & "','" & strCodSucursal & "','" & strCodAgencia & "') }"
    
    ElseIf cboTipoSolicitud.ListIndex > -1 And cboSucursal.ListIndex > 0 Then
    
        strSql = "{ call up_ACSelDatosParametro(32,'" & strCodFondoSolicitud & "','" & _
            gstrCodAdministradora & "','" & strFechaDesde & "','" & strFechaHasta & "','" & _
            strCodTipoSolicitud & "','" & strCodSucursal & "') }"
    
    Else
    
        strSql = "{ call up_ACSelDatosParametro(31,'" & strCodFondoSolicitud & "','" & _
            gstrCodAdministradora & "','" & strFechaDesde & "','" & strFechaHasta & "','" & _
            strCodTipoSolicitud & "') }"
            
    End If
    
    Set adoConsulta = New ADODB.Recordset
    
    strEstado = Reg_Defecto
    With adoConsulta
        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open strSql
    End With
        
    tdgConsulta.DataSource = adoConsulta
    
    If adoConsulta.RecordCount > 0 Then
        strEstado = Reg_Consulta
        
        Dim intNumSuscripciones As Integer, intNumRescates  As Integer

        intNumSuscripciones = 0: intNumRescates = 0
        With adoConsulta
            .MoveFirst
            
            Do While Not .EOF
                If Left(.Fields("TipoSolicitud"), 1) = "S" Then intNumSuscripciones = intNumSuscripciones + 1
                If Left(.Fields("TipoSolicitud"), 1) = "R" Then intNumRescates = intNumRescates + 1

                .MoveNext
            Loop
        End With
        lblDescrip(21).Caption = "Suscripciones (" & CStr(intNumSuscripciones) & ")"
        lblDescrip(30).Caption = "Rescates (" & CStr(intNumRescates) & ")"
    Else
        lblDescrip(21).Caption = "Suscripciones (0)"
        lblDescrip(30).Caption = "Rescates (0)"
    End If
            
    Me.MousePointer = vbDefault
    
End Sub

Private Sub Calcular()

    Dim dblMonto            As Double, dblCuota     As Double
    Dim dblMontoComision    As Double, dblMontoIgv  As Double
    
    If blnCuota Then '*** Calcular Monto ***
        If blnValorConocido Then
            dblMonto = Round(CDbl(txtCantCuotasSolicitud.Text) * txtValorCuota.Value, 2)
            lblMontoSolicitud.Caption = CStr(dblMonto)
                                                
            If strCodTipoOperacion = Codigo_Operacion_Suscripcion Then
                dblMontoComision = Round(dblMonto * dblTasaSuscripcion, 2)
            Else
                dblMontoComision = Round(dblMonto * dblTasaRescate, 2)
            End If
                        
            txtMontoComision.Text = CStr(dblMontoComision)
            dblMontoIgv = Round(dblMontoComision * gdblTasaIgv, 2)
            txtMontoIgv.Text = CStr(dblMontoIgv)
            
            If strCodTipoOperacion = Codigo_Operacion_Suscripcion Then
                txtMontoNetoSolicitud.Text = CStr(dblMonto + (dblMontoComision + dblMontoIgv))
            Else
                txtMontoNetoSolicitud.Text = CStr(dblMonto - (dblMontoComision + dblMontoIgv))
            End If
            
        Else
            txtCantCuotasSolicitud.Text = "0"
        End If
    Else '*** Calcular Cantidad de Cuotas ***
        If blnValorConocido Then
            If strCodTipoOperacion = Codigo_Operacion_Suscripcion Then
                dblMonto = Round(CDbl(txtMontoNetoSolicitud.Text) / (1 + (dblTasaSuscripcion + (dblTasaSuscripcion * gdblTasaIgv))), 2)
                dblMontoComision = Round(dblMonto * dblTasaSuscripcion, 2)
            Else
                dblMonto = Round(CDbl(txtMontoNetoSolicitud.Text) / (1 - (dblTasaRescate + (dblTasaRescate * gdblTasaIgv))), 2)
                dblMontoComision = Round(dblMonto * dblTasaRescate, 2)
            End If
            
            txtMontoComision.Text = CStr(dblMontoComision)
            dblMontoIgv = Round(dblMontoComision * gdblTasaIgv, 2)
            txtMontoIgv.Text = CStr(dblMontoIgv)
            
            '*** Recalculando Monto ***
            If strCodTipoOperacion = Codigo_Operacion_Suscripcion Then
                dblMonto = CDbl(txtMontoNetoSolicitud.Text) - dblMontoComision - dblMontoIgv
            Else
                dblMonto = CDbl(txtMontoNetoSolicitud.Text) + dblMontoComision + dblMontoIgv
            End If
            
            lblMontoSolicitud.Caption = CStr(dblMonto)
            dblCuota = Round(dblMonto / txtValorCuota.Value, 4) 'dblValorCuota
            txtCantCuotasSolicitud.Text = CStr(dblCuota)
            If strCodTipoValuacion <> Codigo_Asignacion_TMenos1 Then txtCantCuotasSolicitud.Text = "0"
        Else
            If strCodTipoOperacion = Codigo_Operacion_Suscripcion Then
                dblMonto = Round(CDbl(txtMontoNetoSolicitud.Text) / (1 + (dblTasaSuscripcion + (dblTasaSuscripcion * gdblTasaIgv))), 2)
                dblMontoComision = Round(dblMonto * dblTasaSuscripcion, 2)
            Else
                dblMonto = Round(CDbl(txtMontoNetoSolicitud.Text) / (1 - (dblTasaRescate + (dblTasaRescate * gdblTasaIgv))), 2)
                dblMontoComision = Round(dblMonto * dblTasaRescate, 2)
            End If
            
            txtMontoComision.Text = CStr(dblMontoComision)
            dblMontoIgv = Round(dblMontoComision * gdblTasaIgv, 2)
            txtMontoIgv.Text = CStr(dblMontoIgv)
            
            '*** Recalculando Monto ***
            If strCodTipoOperacion = Codigo_Operacion_Suscripcion Then
                dblMonto = CDbl(txtMontoNetoSolicitud.Text) - dblMontoComision - dblMontoIgv
            Else
                dblMonto = CDbl(txtMontoNetoSolicitud.Text) + dblMontoComision + dblMontoIgv
            End If
            
            lblMontoSolicitud.Caption = CStr(dblMonto)
            txtCantCuotasSolicitud.Text = "0"
        End If
    End If

End Sub

Public Sub Cancelar()

    cmdOpcion.Visible = True
    With tabSolicitud
        .TabEnabled(0) = True
        .Tab = 0
    End With
    
    gstrCodParticipe = Valor_Caracter
    tmrHora.Enabled = False
    Call Buscar
    
End Sub

Private Sub Deshabilita()

    fraDatosParticipe.Enabled = False
    fraDatosSolicitud.Enabled = False
    fraFormaPago.Enabled = False
    fraTipoAporte.Enabled = False
    
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
        Case vExit
            Call Salir
        
    End Select
    
End Sub

Public Sub Eliminar()

    If strEstado = Reg_Consulta Or strEstado = Reg_Edicion Then
        If MsgBox(Mensaje_Eliminacion, vbQuestion + vbYesNo, gstrNombreEmpresa) = vbYes Then
            adoComm.CommandText = "UPDATE ParticipeSolicitud SET EstadoSolicitud='" & Estado_Solicitud_Anulada & "' " & _
                "WHERE NumSolicitud='" & tdgConsulta.Columns(0) & "' AND CodParticipe='" & tdgConsulta.Columns(8) & "' AND " & _
                "CodFondo='" & strCodFondoSolicitud & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
            adoConn.Execute adoComm.CommandText

            If chkPagoParcial.Value Then
                adoComm.CommandText = "DELETE ParticipePagoSuscripcionTmp " & _
                    "WHERE NumSolicitud='" & tdgConsulta.Columns(0) & "' AND CodParticipe='" & tdgConsulta.Columns(8) & "' AND " & _
                    "CodFondo='" & strCodFondoSolicitud & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
                adoConn.Execute adoComm.CommandText
            End If
            
            Call Buscar
        End If
    End If

End Sub

Public Sub Grabar()

    Dim adoRegistro As ADODB.Recordset
    Dim intRegistro As Integer
    Dim strMensaje As String
    Dim arrCodActivo() As String
    Dim strCodFileActivo As String
    Dim strCodAnaliticaActivo As String
    Dim strCodCuentaActivo As String
    
                
    If strEstado = Reg_Consulta Then Exit Sub
        
    If strEstado = Reg_Adicion Then
        If TodoOK() Then
            Dim strTipoIdentidad    As String, strNumIdentidad  As String
            Dim strClaseCliente     As String, strNumSolicitud  As String
            
            Dim dblCantCuotasSolicitud As Double, dblMontoSolicitud As Double
            Dim dblMontoComision As Double, dblMontoIgv As Double
            Dim dblMontoNetoSolicitud As Double
            
            If strTipoAporte = Tipo_Aporte_Dinerario Then
                strCodFileActivo = Valor_Caracter
                strCodAnaliticaActivo = Valor_Caracter
            ElseIf strTipoAporte = Tipo_Aporte_NoDinerario Then
                arrCodActivo = Split(strCodActivoAporte, "|")
                strCodFileActivo = arrCodActivo(0)
                strCodAnaliticaActivo = arrCodActivo(1)
                strCodCuentaActivo = arrCodActivo(2)
            End If
            
'            adoComm.CommandText = "{ call up_ACSelDatosParametro(64,'" & strCodFondo & "') }"
'            Set adoRegistro = adoComm.Execute
'
'            If Not adoRegistro.EOF Then
'                If IsNull(Mid(adoRegistro("NumPapeleta"), 1, 15)) Then
'                    txtNumPapeleta.Text = Format(1, "000000000000000")
'                Else
'                    txtNumPapeleta.Text = adoRegistro("NumPapeleta")
'                End If
'            Else
'                txtNumPapeleta.Text = Format(1, "000000000000000")
'            End If
            
            'adoRegistro.Close: Set adoRegistro = Nothing
            
            
            strNumSolicitud = Valor_Caracter
            
            dblCantCuotasSolicitud = 0: dblMontoSolicitud = 0
            dblMontoComision = 0: dblMontoIgv = 0
            dblMontoNetoSolicitud = 0
            
            If blnValorConocido Then
                dblCantCuotasSolicitud = CDec(txtCantCuotasSolicitud.Text)
                dblMontoSolicitud = CDec(lblMontoSolicitud.Caption)
                dblMontoComision = CDec(txtMontoComision.Text)
                dblMontoIgv = CDec(txtMontoIgv.Text)
                dblMontoNetoSolicitud = CDec(txtMontoNetoSolicitud.Text)
            Else
                dblCantCuotasSolicitud = CDec(txtCantCuotasSolicitudD.Text)
                dblMontoSolicitud = CDec(txtMontoSolicitud.Text)
                dblMontoComision = 0 'CDec(lblMontoComision.Caption)
                dblMontoIgv = 0 'CDec(lblMontoIgv.Caption)
                dblMontoNetoSolicitud = CDec(lblMontoNetoSolicitud.Caption)
            End If


            If strCodTipoOperacion = Codigo_Operacion_Suscripcion Then
                If blnValorConocido Then
                    strCodClaseOperacion = Codigo_Clase_SuscripcionConocida
                Else
                    strCodClaseOperacion = Codigo_Clase_SuscripcionDesconocida
                End If
            Else
                If blnValorConocido Then
                    If CDbl(lblCuotasDisponibles.Caption) - CDbl(txtCantCuotasSolicitud.Text) = 0 Then
                        strCodClaseOperacion = Codigo_Clase_RescateTotalConocido
                    Else
                        strCodClaseOperacion = Codigo_Clase_RescateParcialConocido
                    End If
                Else
                    If CDbl(lblCuotasDisponibles.Caption) - CDbl(txtCantCuotasSolicitudD.Text) = 0 Then
                        strCodClaseOperacion = Codigo_Clase_RescateTotalDesconocido
                    Else
                        strCodClaseOperacion = Codigo_Clase_RescateParcialDesconocido
                    End If
                End If
            End If
                                                
                                                
            strMensaje = "Para proceder al Registro de la Solicitud de " & IIf(strCodTipoOperacion = Codigo_Operacion_Suscripcion, "Suscripción", "Rescate") & " confirme los siguientes datos : " & Chr(vbKeyReturn) & Chr(vbKeyReturn) & _
                        "Fecha de Solicitud" & vbTab & vbTab & ">" & Space(2) & CStr(lblFechaSolicitud.Caption) & Chr(vbKeyReturn) & _
                        "Hora de Solicitud" & vbTab & vbTab & ">" & Space(2) & CStr(Format(dtpHoraSolicitud.Value, "hh:mm")) & Chr(vbKeyReturn) & _
                        "Fecha de Liquidación" & vbTab & ">" & Space(2) & CStr(Convertddmmyyyy(strFechaProceso)) & Chr(vbKeyReturn) & _
                        "Cantidad de Cuotas" & vbTab & vbTab & ">" & Space(2) & IIf(blnValorConocido, txtCantCuotasSolicitud.Text, IIf(Not blnValorConocido And strCodTipoOperacion = Codigo_Operacion_Rescate, txtCantCuotasSolicitudD.Text, "DESCONOCIDO")) & Chr(vbKeyReturn) & _
                        "Valor de Cuota" & vbTab & vbTab & ">" & Space(2) & IIf(blnValorConocido, txtValorCuota.Text, "DESCONOCIDO") & Chr(vbKeyReturn) & _
                        "Monto Total" & vbTab & vbTab & ">" & Space(2) & IIf(Not blnValorConocido And strCodTipoOperacion = Codigo_Operacion_Rescate, "DESCONOCIDO", IIf(Not blnValorConocido And strCodTipoOperacion = Codigo_Operacion_Suscripcion, CStr(Format(lblMontoNetoSolicitud.Caption, "###,###,###,###,##0.00")), CStr(Format(txtMontoNetoSolicitud.Text, "###,###,###,###,##0.00")))) & Chr(vbKeyReturn) & Chr(vbKeyReturn) & _
                        "¿ Seguro de continuar ?"

            If MsgBox(strMensaje, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
               Me.Refresh: Exit Sub
            End If
            
            Me.MousePointer = vbHourglass
           
            Set adoRegistro = New ADODB.Recordset
            
            With adoComm
                                
                .CommandType = adCmdText
                
                '*** Guardar Solicitud ***
                .CommandText = "{ call up_PRManSolicitudParticipe('" & _
                    strCodFondo & "','" & gstrCodAdministradora & "','" & _
                    strNumSolicitud & "','" & gstrCodParticipe & "','" & strCodCliente & "','" & strCodComisionista & "'," & numSecCondicion & ",'" & _
                    Trim(txtNumPapeleta.Text) & "','" & strNumCertificado & "','" & _
                    strCodSucursalSolicitud & "','" & strCodSucursalDestino & "','" & _
                    strCodAgenciaSolicitud & "','" & strCodAgenciaDestino & "','" & _
                    strCodEjecutivo & "','" & strCodEjecutivoDestino & "','" & _
                    strCodTipoOperacion & "','" & strCodClaseOperacion & "','" & _
                    Convertyyyymmdd(CVDate(lblFechaSolicitud.Caption)) & Space(1) & Format(dtpHoraSolicitud.Value, "hh:mm") & "','" & _
                    strFechaProceso & "','" & _
                    strCodMonedaFondo & "',"
                
                If strCodTipoOperacion = Codigo_Operacion_Suscripcion Then
                    .CommandText = .CommandText & dblTasaSuscripcion & ","
                Else
                    .CommandText = .CommandText & dblTasaRescate & ","
                End If
                
                .CommandText = .CommandText & CDbl(txtValorCuota.Value) & "," & dblCantCuotasSolicitud & "," & _
                    dblMontoSolicitud & "," & dblMontoComision & "," & _
                    dblMontoIgv & "," & dblMontoNetoSolicitud & ",'" & _
                    strCodTipoFormaPago & "','" & strCodBanco & "','" & _
                    strCodBancoDestino & "','" & strTipoCuenta & "','" & _
                    strTipoCuentaDestino & "','" & strNumCuenta & "','" & _
                    strNumCuentaDestino & "','" & Trim(txtNumCheque.Text) & "','" & _
                    "','','" & _
                    "X','','" & _
                    Convertyyyymmdd(CVDate(Valor_Fecha)) & "','"
                
                If blnValorConocido Then
                    .CommandText = .CommandText & "X','"
                Else
                    .CommandText = .CommandText & "','"
                End If
                
                .CommandText = .CommandText & "','" & _
                "',0,'" & strIndPagoParcial & "','" & _
                Estado_Solicitud_Ingresada & "','" & _
                gstrLogin & "','" & Convertyyyymmdd(gdatFechaActual) & "','" & _
                gstrLogin & "','" & Convertyyyymmdd(gdatFechaActual) & "','" & _
                strCodTipoOperacion & "','" & Trim(lblDescripParticipe.Caption) & "','" & _
                "','','" & strCodFileActivo & "','" & strCodAnaliticaActivo & "','" & strCodCuentaActivo & "','I') }"
                                
                Set adoRegistro = adoConn.Execute(.CommandText)

                If Not adoRegistro.EOF Then
                    strNumSolicitud = adoRegistro("NumSolicitud")
                    strNumPapeleta = adoRegistro("NumFolio")
                End If
                
                adoRegistro.Close: Set adoRegistro = Nothing


            End With

            Me.MousePointer = vbDefault
            
            strMensaje = Mensaje_Adicion_Exitosa & Chr(vbKeyReturn) & _
                        "Se asignó a esta solicitud el Nro. de Operacion de Solicitud" & Space(2) & ">" & Space(2) & strNumPapeleta
            
            MsgBox strMensaje, vbExclamation
                                                                            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"

            cmdOpcion.Visible = True
            With tabSolicitud
                .TabEnabled(0) = True
                .Tab = 0
            End With
                        
            Call Buscar
        End If
    End If



'    Dim adoRegistro As ADODB.Recordset
'    Dim intRegistro As Integer
'
'    If strEstado = Reg_Consulta Then Exit Sub
'
'    If strEstado = Reg_Adicion Then
'        If TodoOK() Then
'            Dim strTipoIdentidad    As String, strNumIdentidad  As String
'            Dim strClaseCliente     As String, strNumSolicitud  As String
'
'            Me.MousePointer = vbHourglass
'
'            If strCodTipoOperacion = Codigo_Operacion_Suscripcion Then
'                If blnValorConocido Then
'                    strCodClaseOperacion = Codigo_Clase_SuscripcionConocida
'                Else
'                    strCodClaseOperacion = Codigo_Clase_SuscripcionDesconocida
'                End If
'            Else
'                If blnValorConocido Then
'                    strCodClaseOperacion = Codigo_Clase_RescateTotalConocido
'                Else
'                    strCodClaseOperacion = Codigo_Clase_RescateTotalDesconocido
'                End If
'            End If
'
'            With adoComm
'                .CommandType = adCmdStoredProc
'
'                '*** Obtener el número del parámetro **
'                .CommandText = "up_ACObtenerUltNumero"
'                .Parameters.Append .CreateParameter("CodFondo", adChar, adParamInput, 3, strCodFondo)
'                .Parameters.Append .CreateParameter("CodAdministradora", adChar, adParamInput, 3, gstrCodAdministradora)
'                .Parameters.Append .CreateParameter("CodParametro", adChar, adParamInput, 6, Valor_NumSolicitud)
'                .Parameters.Append .CreateParameter("NuevoNumero", adChar, adParamOutput, 10, "")
'                .Execute
'
'                If Not .Parameters("NuevoNumero") Then
'                    strNumSolicitud = .Parameters("NuevoNumero").Value
'                    .Parameters.Delete ("CodFondo")
'                    .Parameters.Delete ("CodAdministradora")
'                    .Parameters.Delete ("CodParametro")
'                    .Parameters.Delete ("NuevoNumero")
'                End If
'
'                .CommandType = adCmdText
'
'                '*** Guardar Solicitud ***
'                .CommandText = "{ call up_PRManSolicitudParticipe('" & _
'                    strCodFondo & "','" & gstrCodAdministradora & "','" & _
'                    gstrCodParticipe & "','" & strNumSolicitud & "','" & _
'                    Trim(txtNumPapeleta.Text) & "','','" & _
'                    strCodSucursalSolicitud & "','" & strCodSucursalDestino & "','" & _
'                    strCodAgenciaSolicitud & "','" & strCodAgenciaDestino & "','" & _
'                    strCodEjecutivo & "','" & strCodEjecutivoDestino & "','" & _
'                    strCodTipoOperacion & "','" & strCodClaseOperacion & "','" & _
'                    Convertyyyymmdd(CVDate(lblFechaSolicitud.Caption)) & Space(1) & Format(dtpHoraSolicitud.Value, "hh:mm") & "','" & _
'                    Convertyyyymmdd(CVDate(lblFechaSolicitud.Caption)) & "','" & _
'                    strCodMonedaFondo & "',"
'                If strCodTipoOperacion = Codigo_Operacion_Suscripcion Then
'                    .CommandText = .CommandText & dblTasaSuscripcion & ","
'                Else
'                    .CommandText = .CommandText & dblTasaRescate & ","
'                End If
'                .CommandText = .CommandText & txtValorCuota.Value & "," & CDec(txtCantCuotasSolicitud.Text) & "," & _
'                    CDec(lblMontoSolicitud.Caption) & "," & CDec(txtMontoComision.Text) & "," & _
'                    CDec(txtMontoIgv.Text) & "," & CDec(txtMontoNetoSolicitud.Text) & ",'" & _
'                    strCodTipoFormaPago & "','" & strCodBanco & "','" & _
'                    strCodBancoDestino & "','','" & _
'                    "','" & strNumCuenta & "','" & _
'                    strNumCuentaDestino & "','" & Trim(txtNumCheque.Text) & "','" & _
'                    "','','" & _
'                    "X','','" & _
'                    Convertyyyymmdd(CVDate(Valor_Fecha)) & "','"
'                If blnValorConocido Then
'                    .CommandText = .CommandText & "X','"
'                Else
'                    .CommandText = .CommandText & "','"
'                End If
'                .CommandText = .CommandText & "','" & _
'                    "',0,'" & strIndPagoParcial & "','" & _
'                    Estado_Solicitud_Ingresada & "','" & _
'                    gstrLogin & "','" & Convertyyyymmdd(gdatFechaActual) & "','" & _
'                    gstrLogin & "','" & Convertyyyymmdd(gdatFechaActual) & "','" & _
'                    Trim(lblDescripParticipe.Caption) & "','" & _
'                    "I') }"
'                adoConn.Execute .CommandText
'
'                '*** Guardar Detalle Solicitud ***
'                If strCodTipoOperacion = Codigo_Operacion_Suscripcion Then
'                    .CommandText = "{ call up_PRManSolicitudParticipeDetalle('" & _
'                        strCodFondo & "','" & gstrCodAdministradora & "','" & _
'                        gstrCodParticipe & "','" & strNumSolicitud & "'," & _
'                        "1,''," & CDec(txtCantCuotasSolicitud.Text) & ",'" & _
'                        gstrCodParticipe & "','" & Trim(lblDescripParticipe.Caption) & "','" & _
'                        "X','I') }"
'                    adoConn.Execute .CommandText
'                Else
'                    Dim strUltNumCertificado    As String
'
'                    If gstrTipoAdministradora = Codigo_Tipo_Fondo_Inversion Then
'                        Set adoRegistro = New ADODB.Recordset
'
''                        With adoComm
'                            intRegistro = 1
'
'                            .CommandText = "SELECT * FROM ParticipeCertificado " & _
'                                "WHERE CodParticipe='" & gstrCodParticipe & "' AND CodFondo='" & strCodFondo & "' AND " & _
'                                "CodAdministradora='" & gstrCodAdministradora & "' AND IndVigente='X' AND IndBloqueo=''"
'                            Set adoRegistro = .Execute
'
'                            Do While Not adoRegistro.EOF
'                                .CommandText = "{ call up_PRManSolicitudParticipeDetalle('" & _
'                                    strCodFondo & "','" & gstrCodAdministradora & "','" & _
'                                    gstrCodParticipe & "','" & strNumSolicitud & "'," & _
'                                    intRegistro & ",'" & Trim(adoRegistro("NumCertificado")) & "'," & _
'                                    CDec(adoRegistro("CantCuotas")) * -1 & ",'" & _
'                                    gstrCodParticipe & "','" & Trim(lblDescripParticipe.Caption) & "','" & _
'                                    "X','I') }"
'                                adoConn.Execute .CommandText
'
'                                intRegistro = intRegistro + 1
'                                adoRegistro.MoveNext
'                            Loop
'                            adoRegistro.Close: Set adoRegistro = Nothing
''                        End With
'                    Else
'                        Dim dblCuotasRescate    As Double
'
'                        Set adoRegistro = New ADODB.Recordset
'                        dblCuotasRescate = CDbl(txtCantCuotasSolicitud.Text)
'
''                        With adoComm
'                            intRegistro = 1
'                            .CommandText = "SELECT * FROM ParticipeCertificado " & _
'                                "WHERE CodParticipe='" & gstrCodParticipe & "' AND CodFondo='" & strCodFondo & "' AND " & _
'                                "CodAdministradora='" & gstrCodAdministradora & "' AND IndVigente='X' AND IndBloqueo='' " & _
'                                "ORDER BY FechaSuscripcion"
'                            Set adoRegistro = .Execute
'
'                            Do While Not adoRegistro.EOF
'                                dblCuotasRescate = dblCuotasRescate - CDbl(adoRegistro("CantCuotas"))
'                                strUltNumCertificado = Trim(adoRegistro("NumCertificado"))
'
'                                If dblCuotasRescate > 0 Then
'                                    .CommandText = "{ call up_PRManSolicitudParticipeDetalle('" & _
'                                        strCodFondo & "','" & gstrCodAdministradora & "','" & _
'                                        gstrCodParticipe & "','" & strNumSolicitud & "'," & _
'                                        intRegistro & ",'" & Trim(adoRegistro("NumCertificado")) & "'," & _
'                                        CDec(adoRegistro("CantCuotas")) * -1 & ",'" & _
'                                        gstrCodParticipe & "','" & Trim(lblDescripParticipe.Caption) & "','" & _
'                                        "X','I') }"
'                                    adoConn.Execute .CommandText
'
'                                    intRegistro = intRegistro + 1
'                                Else
'                                    dblCuotasRescate = dblCuotasRescate + CDbl(adoRegistro("CantCuotas"))
'
'                                    .CommandText = "{ call up_PRManSolicitudParticipeDetalle('" & _
'                                        strCodFondo & "','" & gstrCodAdministradora & "','" & _
'                                        gstrCodParticipe & "','" & strNumSolicitud & "'," & _
'                                        intRegistro & ",'" & Trim(adoRegistro("NumCertificado")) & "'," & _
'                                        dblCuotasRescate * -1 & ",'" & _
'                                        gstrCodParticipe & "','" & Trim(lblDescripParticipe.Caption) & "','" & _
'                                        "X','I') }"
'                                    adoConn.Execute .CommandText
'
'                                    dblCuotasRescate = Abs(dblCuotasRescate - CDbl(adoRegistro("CantCuotas")))
'
'                                    intRegistro = intRegistro + 1
'
'                                    .CommandText = "{ call up_PRManSolicitudParticipeDetalle('" & _
'                                        strCodFondo & "','" & gstrCodAdministradora & "','" & _
'                                        gstrCodParticipe & "','" & strNumSolicitud & "'," & _
'                                        intRegistro & ",''," & _
'                                        dblCuotasRescate & ",'" & _
'                                        gstrCodParticipe & "','" & Trim(lblDescripParticipe.Caption) & "','" & _
'                                        "X','I') }"
'                                    adoConn.Execute .CommandText
'
'                                    Exit Do
'                                End If
'
'                                adoRegistro.MoveNext
'                            Loop
'                            adoRegistro.Close: Set adoRegistro = Nothing
''                        End With
'                    End If
'
'                    '*** Actualizar el último número del certificado rescatado ***
'                    .CommandText = "UPDATE ParticipeSolicitud SET NumCertificado='" & strUltNumCertificado & "' " & _
'                        "WHERE NumSolicitud='" & strNumSolicitud & "' AND CodParticipe='" & gstrCodParticipe & "' AND " & _
'                        "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
'                    adoConn.Execute .CommandText
'                End If
'
'                '*** Pago Parcial de Cuotas de Suscripción ***
'                If strIndPagoParcial = Valor_Indicador Then
'                    Dim datFechaPagoParcial As Date
'                    Dim intContador         As Integer
'                    Dim dblCuotasPagadas    As Double, dblCuotasPagadasReal    As Double
'
'                    Set adoRegistro = New ADODB.Recordset
'
'                    .CommandText = "SELECT NumSecuencial,FechaDesde,FechaHasta,PorcenPago,MontoPago FROM FondoPagoSuscripcion " & _
'                        "WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' " & _
'                        "ORDER BY NumSecuencial"
'                    Set adoRegistro = .Execute
'
'                    Do While Not adoRegistro.EOF
'                        datFechaPagoParcial = CVDate(adoRegistro("FechaHasta"))
'
'                        If CDbl(adoRegistro("PorcenPago")) > 0 Then
'                            dblCuotasPagadas = Round(CDbl(txtCantCuotasSolicitud.Text) * CDbl(adoRegistro("PorcenPago")) * 0.01, 5)
'                            dblCuotasPagadasReal = Round(CDbl(txtCantCuotasSolicitud.Text) * CDbl(adoRegistro("PorcenPago")) * 0.01, 5)
'
'                            .CommandText = "{ call up_PRManParticipePagoSuscripcionTmp('" & _
'                                strCodFondo & "','" & gstrCodAdministradora & "','" & gstrCodParticipe & "','" & _
'                                strNumSolicitud & "'," & CInt(adoRegistro("NumSecuencial")) & ",'" & _
'                                Convertyyyymmdd(datFechaPagoParcial) & "'," & _
'                                CDec(CCur(txtMontoNetoSolicitud.Text) * adoRegistro("PorcenPago") * 0.01) & ",'" & _
'                                Convertyyyymmdd(CVDate(Valor_Fecha)) & "'," & dblCuotasPagadas & "," & _
'                                dblCuotasPagadasReal & ",'I') }"
'                        Else
'                            dblCuotasPagadas = Round((CDbl(adoRegistro("MontoPago")) * CDbl(txtMontoNetoSolicitud.Text)) / CDbl(txtCantCuotasSolicitud.Text), 5)
'                            dblCuotasPagadasReal = Round((CDbl(adoRegistro("MontoPago")) * CDbl(txtMontoNetoSolicitud.Text)) / CDbl(txtCantCuotasSolicitud.Text), 5)
'
'                            .CommandText = "{ call up_PRManParticipePagoSuscripcionTmp('" & _
'                                strCodFondo & "','" & gstrCodAdministradora & "','" & gstrCodParticipe & "','" & _
'                                strNumSolicitud & "'," & CInt(adoRegistro("NumSecuencial")) & ",'" & _
'                                Convertyyyymmdd(datFechaPagoParcial) & "'," & _
'                                CDec(adoRegistro("MontoPago")) & ",'" & _
'                                Convertyyyymmdd(CVDate(Valor_Fecha)) & "'," & dblCuotasPagadas & "," & _
'                                dblCuotasPagadasReal & ",'I') }"
'                        End If
'                        adoConn.Execute .CommandText
'
'                        adoRegistro.MoveNext
'                    Loop
'                    adoRegistro.Close: Set adoRegistro = Nothing
'                End If
'
'                '*** Actualizar Secuenciales ***
'                .CommandText = "{ call up_ACActUltNumero('" & strCodFondo & "','" & gstrCodAdministradora & "','" & _
'                    Valor_NumSolicitud & "','" & strNumSolicitud & "') }"
'                adoConn.Execute .CommandText
'
'            End With
'
'            Me.MousePointer = vbDefault
'
'            MsgBox Mensaje_Adicion_Exitosa, vbExclamation
'
'            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
'
'            cmdOpcion.Visible = True
'            With tabSolicitud
'                .TabEnabled(0) = True
'                .Tab = 0
'            End With
'
'            Call Buscar
'        End If
'    End If
    
End Sub

Private Function TodoOK() As Boolean

    TodoOK = False
    
    If cboTipoOperacion.ListIndex < 0 Then
        MsgBox "Seleccione el Tipo de Operación.", vbCritical
        cboTipoOperacion.SetFocus
        Exit Function
    End If
    
    If cboFondo.ListIndex = 0 Then
        MsgBox "Seleccione el Fondo.", vbCritical
        cboFondo.SetFocus
        Exit Function
    End If
        
    If Trim(txtNumPapeleta.Text) = Valor_Caracter Then
        MsgBox "El Campo Número de Papeleta no es Válido!.", vbCritical
        txtNumPapeleta.SetFocus
        Exit Function
    End If
    
    If Trim(txtNumDocumento.Text) = Valor_Caracter Then
        MsgBox "Debe seleccionar el Partícipe.", vbCritical
        cmdBusqueda.SetFocus
        Exit Function
    End If
    
    If Trim(lblDescripParticipe.Caption) = Valor_Caracter Then
        MsgBox "El Campo Descripción no es Válido!, presione ENTER en el campo Número de Documento.", vbCritical
        txtNumDocumento.SetFocus
        Exit Function
    End If
    
    If strTipoAporte = Tipo_Aporte_NoDinerario Then
        If Not cboActivoAporte.Visible Then
            MsgBox "No existen activos no dinerarios para registrar", vbCritical
            Exit Function
        End If
        If cboActivoAporte.ListIndex <= 0 Then
            MsgBox "No ha seleccionado el activo no dinerario a registrar", vbCritical
            Exit Function
        End If
    End If
    
    If CCur(lblMontoSolicitud.Caption) = 0 And blnValorConocido Then
        MsgBox "El Valor de la Operación no ha sido calculado, presione ENTER en el campo Monto.", vbCritical
        Exit Function
    End If
    
    If txtMontoSolicitud.Value = 0 And Not blnValorConocido And strCodTipoOperacion = Codigo_Operacion_Suscripcion Then
        MsgBox "Debe Ingresar el Valor de la Operación.", vbCritical
        txtMontoSolicitud.SetFocus
        Exit Function
    End If
    
    If cboTipoFormaPago.ListIndex = 0 Then
        MsgBox "Seleccione el Tipo de Forma de Pago.", vbCritical
        cboTipoFormaPago.SetFocus
        Exit Function
    End If
        
    If gstrTipoAdministradora = Codigo_Tipo_Fondo_Inversion Then
        If strCodTipoOperacion = Codigo_Operacion_Suscripcion Then
            Dim adoRegistro         As ADODB.Recordset
            Dim curMontoControl     As Currency
            
            Set adoRegistro = New ADODB.Recordset
            
            With adoComm
                .CommandText = "SELECT SUM(MontoNetoSolicitud) MontoNetoSolicitud FROM ParticipeSolicitud " & _
                    "WHERE EstadoSolicitud <> '" & Estado_Solicitud_Anulada & "' AND " & _
                    "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
                Set adoRegistro = .Execute
                
                If Not adoRegistro.EOF Then
                    If Not IsNull(adoRegistro("MontoNetoSolicitud")) Then
                        curMontoControl = CCur(adoRegistro("MontoNetoSolicitud"))
                    Else
                        curMontoControl = 0
                    End If
                End If
                adoRegistro.Close: Set adoRegistro = Nothing
            End With
            
            curMontoControl = curMontoControl + CCur(txtMontoNetoSolicitud.Text)
            
            If curMontoControl > curMontoEmitido Then
                MsgBox "Las Suscripciones Totales superan el Monto Emitido en " & Format(curMontoControl - curMontoEmitido, "###,###,###,###,##0.00") & " , por favor verifique !", vbCritical
                txtMontoNetoSolicitud.SetFocus
                Exit Function
            End If
        End If
    End If
    
    If strCodTipoOperacion = Codigo_Operacion_Rescate Then
        
        If CDbl(lblCuotasDisponibles.Caption) <= 0 Then
            MsgBox "No tiene cuotas disponibles para Rescatar!", vbCritical
            txtCantCuotasSolicitud.SetFocus
            Exit Function
        End If
        
        If CDbl(lblCuotasDisponibles.Caption) - CDbl(txtCantCuotasSolicitud.Text) < 0 Then
            MsgBox "Las cuotas a Rescatar superan el disponible en cuotas por " & Format(-CDbl(lblCuotasDisponibles.Caption) + CDbl(txtCantCuotasSolicitud.Text), "###,###,###,###,##0.00") & " , por favor verifique !", vbCritical
            txtCantCuotasSolicitud.SetFocus
            Exit Function
        End If
        
'        If CDbl(lblCuotasDisponibles.Caption) - CDbl(txtCantCuotasSolicitud.Text) > 0 Then
'            MsgBox "No se permiten Rescates Parciales de cuotas Desconocidos!", vbCritical
'            txtCantCuotasSolicitud.SetFocus
'            Exit Function
'        End If
    
        If cboCertificado.ListIndex < 0 Then
            MsgBox "Debe seleccionar un certificado a Rescatar!", vbCritical
            cboCertificado.SetFocus
            Exit Function
        End If
    
    End If
    
    '20141203_JCC Inicio
    If strCodTipoFormaPago = Codigo_FormaPago_Transferencia_Mismo_Banco And strCodBanco <> strCodBancoDestino Then
        MsgBox "Para Transferencia a Mismo Banco requiere que la cuenta origen y la cuenta destino pertenezcan a la misma entidad bancaria.", vbCritical
        cboCuenta.SetFocus
        Exit Function
    End If

    If strCodTipoFormaPago = Codigo_FormaPago_Transferencia_Otro_Banco And strCodBanco = strCodBancoDestino Then
        MsgBox "Ha seleccionado forma de pago Transferencia a Otro Banco y la cuenta origen y la cuenta destino pertenecen a la misma entidad bancaria.", vbCritical
        cboCuenta.SetFocus
        Exit Function
    End If
    
    If strCodTipoFormaPago = Codigo_FormaPago_Transferencia_Exterior Then
        Set adoRegistro = New ADODB.Recordset
        
        With adoComm
            .CommandText = "SELECT CodNacionalidad FROM InstitucionPersona WHERE TipoPersona = '" & Codigo_Tipo_Persona_Emisor & "' AND " & _
                            "CodPersona = '" & strCodBanco & "'"
            Set adoRegistro = .Execute
            
            If Trim(adoRegistro("CodNacionalidad")) = Valor_Caracter Or Trim(adoRegistro("CodNacionalidad")) = "001" Then
                MsgBox "Ha seleccionado forma de pago Transferencia del Exterior y la cuenta origen no es de una entidad con nacionalidad Extranjera.", vbCritical
                adoRegistro.Close: Set adoRegistro = Nothing
                cboCuenta.SetFocus
                Exit Function
            End If
        End With
        
        
    End If
    '20141203_JCC Fin
    
    '*** Si todo pasó OK ***
    TodoOK = True

End Function
Private Sub Habilita()

    fraDatosParticipe.Enabled = True
    fraDatosSolicitud.Enabled = True
    fraTipoAporte.Enabled = True
    
End Sub

Public Sub Imprimir()

End Sub

Public Sub Modificar()

'    cmdOpcion.Visible = False
    With tabSolicitud
'        .TabEnabled(0) = False
'        .Tab = 1
        .Tab = 0
    End With
    
End Sub

Private Sub ObtenerCuentasParticipe()
        
    strSql = "SELECT (CB.Banco + CB.TipoCtaCte + CB.NroCtaCte) CODIGO, " & _
                "(RTRIM(AP.DescripParametro) + SPACE(1) + RTRIM(MO.Signo) + SPACE(1) + RTRIM(IP.DescripPersona) + SPACE(1) + RTRIM(CB.NroCtaCte)) DESCRIP " & _
                "FROM ClienteBancarios CB " & _
                "JOIN InstitucionPersona IP ON (IP.TipoPersona = '" & Codigo_Tipo_Persona_Emisor & "' AND IP.IndBanco = '" & Valor_Indicador & "' and IP.CodPersona = CB.Banco) " & _
                "JOIN AuxiliarParametro AP ON (AP.CodTipoParametro = 'CTAFON' AND AP.CodParametro = CB.TipoCtaCte) " & _
                "JOIN Moneda MO ON (MO.CodMoneda = CB.TipoMoneda) WHERE CB.CodCliente = '" & strCodCliente & "'"

    CargarControlLista strSql, cboCuenta, arrCuenta(), ""
    
    If cboCuenta.ListCount > 0 Then
        cboCuenta.ListIndex = 0
        indExisteCuentaParticipe = Valor_Indicador
    Else
        indExisteCuentaParticipe = Valor_Caracter
    End If
        
End Sub

Private Sub ObtenerCuentasFondo()
                
    '20141201_JJCC strSql = "SELECT (FC.CodBanco + FC.NumCuentaBanco) CODIGO,(DescripCuenta + space(1) + NumCuentaBanco) DESCRIP " &
    strSql = "SELECT (FC.CodBanco + FC.TipoCuenta + FC.NumCuentaBanco) CODIGO,(DescripCuenta + space(1) + NumCuentaBanco) DESCRIP " & _
        "FROM FondoCuenta FC JOIN BancoCuenta BC " & _
        "ON(BC.CodFondo = FC.CodFondo AND BC.CodAdministradora = FC.CodAdministradora AND " & _
        "BC.CodFile = FC.CodFile AND BC.CodAnalitica = FC.CodAnalitica AND BC.TipoCuenta = FC.TipoCuenta) " & _
        "WHERE FC.CodFondo='" & strCodFondo & "' AND FC.CodAdministradora='" & gstrCodAdministradora & "' AND " & _
        "TipoOperacion='" & strCodTipoOperacion & "' AND BC.CodMoneda='" & strCodMonedaFondo & "'"
        
    CargarControlLista strSql, cboCuentaFondo, arrCuentaFondo(), ""
    
    If cboCuentaFondo.ListCount > 0 Then
        cboCuentaFondo.ListIndex = 0
    Else
        MsgBox "El Fondo no tiene cuentas definidas...", vbCritical, gstrNombreEmpresa
    End If
        
End Sub

Private Sub ObtenerParametrosFondo()

    Dim adoRegistro As ADODB.Recordset
    
    Set adoRegistro = New ADODB.Recordset
                
    With adoComm
        '*** Hora de Corte ***
        .CommandText = "{ call up_ACSelDatosParametro(24,'" & strCodFondo & "','" & gstrCodAdministradora & "') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            strHoraCorte = adoRegistro("HoraCorte")
            strCodTipoValuacion = adoRegistro("TipoValuacion")
        End If
        adoRegistro.Close
                        
        If strHoraCorte = "00:00" Then
            MsgBox "No ha definido la hora de corte.", vbCritical, Me.Caption
            cboFondo.ListIndex = 0
            Exit Sub
        End If
        
        If dtpHoraSolicitud.Value > CVDate(strHoraCorte) Then
            blnValorConocido = False
            txtCantCuotasSolicitud.Enabled = False
            Call ColorControlDeshabilitado(txtCantCuotasSolicitud)
            'lblValorCuota.Caption = "0"
            txtValorCuota.Text = "0"
        Else
            blnValorConocido = True
            txtCantCuotasSolicitud.Enabled = True
            Call ColorControlHabilitado(txtCantCuotasSolicitud)
            If strCodTipoValuacion <> Codigo_Asignacion_TMenos1 Then
                txtCantCuotasSolicitud.Enabled = False
                Call ColorControlDeshabilitado(txtCantCuotasSolicitud)
            End If
        End If
        
        If strCodTipoValuacion <> Codigo_Asignacion_TMenos1 Then txtValorCuota.Text = "0" 'lblValorCuota.Caption = "0"
                            
        '*** Obtener Código de Comisión ***
        .CommandText = "SELECT CodComision FROM FondoComision WHERE CodFondo='" & strCodFondo & "' AND " & _
            "CodAdministradora='" & gstrCodAdministradora & "' AND CodOperacion='" & strCodTipoOperacion & "'"
        Set adoRegistro = .Execute
        
        strCodComision = Valor_Caracter
        If Not adoRegistro.EOF Then
            strCodComision = Trim(adoRegistro("CodComision"))
        End If
        adoRegistro.Close
                        
        '*** Valores Minimos y Máximos, Cantidad de Partes Pago Cuotas de Suscripción ***
        .CommandText = "{ call up_ACSelDatosParametro(25,'" & strCodFondo & "','" & gstrCodAdministradora & "') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            dblCantCuotaMinSuscripcionInicial = CDbl(adoRegistro("CantCuotaMinSuscripcionInicial"))
            dblMontoMinSuscripcionInicial = CDbl(adoRegistro("MontoMinSuscripcionInicial"))
            dblCantMinCuotaSuscripcion = CDbl(adoRegistro("CantMinCuotaSuscripcion"))
            dblMontoMinSuscripcion = CDbl(adoRegistro("MontoMinSuscripcion"))
            dblPorcenMaxParticipe = CDbl(adoRegistro("PorcenMaxParticipe"))
            curMontoEmitido = CCur(adoRegistro("MontoEmitido"))
            
            datFechaEtapaOperativa = adoRegistro("FechaInicioEtapaOperativa")
            datFechaEtapaPreOperativa = adoRegistro("FechaInicioEtapaPreOperativa")
            
            chkPagoParcial.Value = vbUnchecked
            If CInt(adoRegistro("CantPartesPagoSuscripcion")) > 1 Then chkPagoParcial.Value = vbChecked
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
                        
    End With
    
End Sub
Private Sub ObtenerParametrosFondoOP()

    Dim adoRegistro As ADODB.Recordset
    Dim dblComision As Double
    
    Set adoRegistro = New ADODB.Recordset
                
    With adoComm
        '*** Hora de Corte ***
        .CommandText = "{ call up_ACSelDatosParametro(24,'" & strCodFondo & "','" & gstrCodAdministradora & "') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            strHoraCorte = adoRegistro("HoraCorte")
            strCodTipoValuacion = adoRegistro("TipoValuacion")
        End If
        adoRegistro.Close
                        
        If strHoraCorte = "00:00" Then
            MsgBox "No ha definido la hora de corte.", vbCritical, Me.Caption
            cboFondo.ListIndex = 0
            Exit Sub
        End If
        
'        SI CRITERIO ES T - 1 ENTONCES
'           SI HORA DE CORTE >= HORA DE OPERACIÓN ENTONCES
'               ACTIVA INGRESO DE CUOTAS (SUSC Y RESC) A VALOR CONOCIDO (CAMPOS VC, Q, COMISIONES, TOTAL)
'               FECHA PROCESO = FECHA OPERACIÓN
'           SINO
'               SI ES SUSCRIPCION ENTONCES
'                   ACTIVA INGRESO DE SUSC A VALOR DESCONOCIDO (CAMPO TOTAL)
'                   DESACTIVA INGRESO DE SUSC A VALOR CONOCIDO (CAMPOS VC, Q, COMISIONES)
'               SI ES RESCATE ENTONCES
'                   ACTIVA INGRESO DE RESC A VALOR DESCONOCIDO (CAMPO Q)
'                   DESACTIVA INGRESO DE RESC A VALOR CONOCIDO (CAMPOS VC, COMISIONES, TOTAL)
'
'            FECHA PROCESO = FECHA OPERACIÓN + 1
'
'       SI CRITERIO ES T + 1 ENTONCES
'               SI ES SUSCRIPCION ENTONCES
'                   ACTIVA INGRESO DE SUSC A VALOR DESCONOCIDO (CAMPO TOTAL)
'                   DESACTIVA INGRESO DE SUSC A VALOR CONOCIDO (CAMPOS VC, Q, COMISIONES)
'               SI ES RESCATE ENTONCES
'                   ACTIVA INGRESO DE RESC A VALOR DESCONOCIDO (CAMPO Q)
'                   DESACTIVA INGRESO DE RESC A VALOR CONOCIDO (CAMPOS VC, COMISIONES, TOTAL)
'
'           SI HORA DE CORTE >= HORA DE OPERACIÓN ENTONCES
'               FECHA PROCESO = FECHA OPERACIÓN + 1
'           SINO
'               FECHA PROCESO = FECHA OPERACIÓN + 2
        If strCodTipoValuacion = Codigo_Asignacion_TMenos1 Then
            If CVDate(strHoraCorte) >= dtpHoraSolicitud.Value Then
                blnValorConocido = True
                fraDatosSolicitud.Visible = True
                fraDatosSolicitudDesconocida.Visible = False
                strFechaProceso = Convertyyyymmdd(lblFechaSolicitud.Caption)
            Else
                blnValorConocido = False
                If strCodTipoOperacion = Codigo_Operacion_Suscripcion Then
                    fraDatosSolicitud.Visible = False
                    fraDatosSolicitudDesconocida.Visible = True
                    txtCantCuotasSolicitudD.Enabled = False
                    txtCantCuotasSolicitudD.BackColor = &H8000000F
                    txtMontoSolicitud.Enabled = True
                    txtMontoSolicitud.BackColor = &HFFFFFF
                    txtCantCuotasSolicitudD.Text = "0.00"
                    txtMontoSolicitud.Text = "0.00"
                ElseIf strCodTipoOperacion = Codigo_Operacion_Rescate Then
                    fraDatosSolicitud.Visible = False
                    fraDatosSolicitudDesconocida.Visible = True
                    txtCantCuotasSolicitudD.Enabled = True
                    txtCantCuotasSolicitudD.BackColor = &HFFFFFF
                    txtMontoSolicitud.Enabled = False
                    txtMontoSolicitud.BackColor = &H8000000F
                    txtCantCuotasSolicitudD.Text = "0.00"
                    txtMontoSolicitud.Text = "0.00"
                End If
                
                If dtpHoraSolicitud.Value < CVDate(strHoraCorte) Then
                    strFechaProceso = Convertyyyymmdd(lblFechaSolicitud.Caption)
                Else
                    strFechaProceso = Convertyyyymmdd(DateAdd("d", 1, lblFechaSolicitud.Caption))
                End If
                
            End If
        End If
            
        If strCodTipoValuacion = Codigo_Asignacion_T Then
            blnValorConocido = False
            If strCodTipoOperacion = Codigo_Operacion_Suscripcion Then
                fraDatosSolicitud.Visible = False
                fraDatosSolicitudDesconocida.Visible = True
                txtCantCuotasSolicitudD.Enabled = False
                txtCantCuotasSolicitudD.BackColor = &H8000000F
                txtMontoSolicitud.Enabled = True
                txtMontoSolicitud.BackColor = &HFFFFFF
                txtCantCuotasSolicitudD.Text = "0.00"
                txtMontoSolicitud.Text = "0.00"
            ElseIf strCodTipoOperacion = Codigo_Operacion_Rescate Then
                fraDatosSolicitud.Visible = False
                fraDatosSolicitudDesconocida.Visible = True
                txtCantCuotasSolicitudD.Enabled = True
                txtCantCuotasSolicitudD.BackColor = &HFFFFFF
                txtMontoSolicitud.Enabled = False
                txtMontoSolicitud.BackColor = &H8000000F
                txtCantCuotasSolicitudD.Text = "0.00"
                txtMontoSolicitud.Text = "0.00"
            End If
            
            If dtpHoraSolicitud.Value < CVDate(strHoraCorte) Then
                strFechaProceso = Convertyyyymmdd(lblFechaSolicitud.Caption)
            Else
                strFechaProceso = Convertyyyymmdd(DateAdd("d", 1, lblFechaSolicitud.Caption))
            End If
            
        End If
            
            
        If strCodTipoValuacion = Codigo_Asignacion_TMas1 Then
            blnValorConocido = False
            If strCodTipoOperacion = Codigo_Operacion_Suscripcion Then
                fraDatosSolicitud.Visible = False
                fraDatosSolicitudDesconocida.Visible = True
                txtCantCuotasSolicitudD.Enabled = False
                txtCantCuotasSolicitudD.BackColor = &H8000000F
                txtMontoSolicitud.Enabled = True
                txtMontoSolicitud.BackColor = &HFFFFFF
                txtCantCuotasSolicitudD.Text = "0.00"
                txtMontoSolicitud.Text = "0.00"
            ElseIf strCodTipoOperacion = Codigo_Operacion_Rescate Then
                fraDatosSolicitud.Visible = False
                fraDatosSolicitudDesconocida.Visible = True
                txtCantCuotasSolicitudD.Enabled = True
                txtCantCuotasSolicitudD.BackColor = &HFFFFFF
                txtMontoSolicitud.Enabled = False
                txtMontoSolicitud.BackColor = &H8000000F
                txtCantCuotasSolicitudD.Text = "0.00"
                txtMontoSolicitud.Text = "0.00"
            End If
            
            If dtpHoraSolicitud.Value < CVDate(strHoraCorte) Then
                strFechaProceso = Convertyyyymmdd(DateAdd("d", 1, lblFechaSolicitud.Caption))
            Else
                strFechaProceso = Convertyyyymmdd(DateAdd("d", 2, lblFechaSolicitud.Caption))
            End If
            
        End If
                            
        '*** Obtener Código de Comisión ***
        .CommandText = "SELECT CodComision FROM FondoComision WHERE CodFondo='" & strCodFondo & "' AND " & _
            "CodAdministradora='" & gstrCodAdministradora & "' AND CodOperacion='" & strCodTipoOperacion & "'"
        Set adoRegistro = .Execute
        
        strCodComision = Valor_Caracter
        If Not adoRegistro.EOF Then
            strCodComision = Trim(adoRegistro("CodComision"))
        End If
        adoRegistro.Close
        
        'Calcula la comision de suscripcion o rescate
        dblComision = ObtenerComisionParticipacion(strCodComision, strCodFondo, gstrCodAdministradora)

        If strCodTipoOperacion = Codigo_Operacion_Suscripcion Then
            dblTasaSuscripcion = dblComision
        Else
            dblTasaRescate = dblComision
        End If
                        
        '*** Valores Minimos y Máximos, Cantidad de Partes Pago Cuotas de Suscripción ***
        .CommandText = "{ call up_ACSelDatosParametro(25,'" & strCodFondo & "','" & gstrCodAdministradora & "') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            dblCantCuotaMinSuscripcionInicial = CDbl(adoRegistro("CantCuotaMinSuscripcionInicial"))
            dblMontoMinSuscripcionInicial = CDbl(adoRegistro("MontoMinSuscripcionInicial"))
            dblCantMinCuotaSuscripcion = CDbl(adoRegistro("CantMinCuotaSuscripcion"))
            dblMontoMinSuscripcion = CDbl(adoRegistro("MontoMinSuscripcion"))
            dblPorcenMaxParticipe = CDbl(adoRegistro("PorcenMaxParticipe"))
            curMontoEmitido = CCur(adoRegistro("MontoEmitido"))
            
            datFechaEtapaOperativa = adoRegistro("FechaInicioEtapaOperativa")
            datFechaEtapaPreOperativa = adoRegistro("FechaInicioEtapaPreOperativa")
            
            chkPagoParcial.Value = vbUnchecked
            If CInt(adoRegistro("CantPartesPagoSuscripcion")) > 1 Then chkPagoParcial.Value = vbChecked
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
                        
    End With
    
End Sub

Public Sub Salir()

    Unload Me
    
End Sub

Private Sub Cmb_Cod_fond_Click()
    
'    Dim adoresultAux As New Recordset, adoresult As New Recordset
'    Dim v_Ite As Variant
'    Dim res As Integer
'    Dim adoRecord As New Recordset, adoHoraServ As New Recordset
'
'    If Cmb_Cod_fond.ListIndex >= 0 Then
'        s_CodFon = Mid(aMapCod_Fond(Cmb_Cod_fond.ListIndex), 1, 2)
'
'        adoComm.CommandText = "Sp_INF_SelectSolici '13', ''"
'        Set adoHoraServ = adoComm.Execute
'
'        adoComm.CommandText = "Sp_INF_SelectSolici '14', '" & s_CodFon & "'"
'        Set adoRecord = adoComm.Execute
'
'        If (Format(adoHoraServ!HorServ, "HH:MM") < Format(adoRecord!HOR_INIC, "HH:MM")) Or (Format(adoHoraServ!HorServ, "HH:MM") > Format(adoRecord!HOR_TERM, "HH:MM")) Then
'            MsgBox "El Ingreso de las Solicitudes de Suscripción es desde las " & Format(adoRecord!HOR_INIC, "HH:MM") & " hasta las " & Format(adoRecord!HOR_TERM, "HH:MM"), vbInformation, gstrNombreEmpresa
'            adoRecord.Close: Set adoRecord = Nothing
'            adoHoraServ.Close: Set adoHoraServ = Nothing
'            frm_DatPar.Enabled = False
'            frm_ValOri.Enabled = False
'            fraFormaPago.Enabled = False
'            fraCtaFondo.Enabled = False
'            Exit Sub
'        Else
'            adoRecord.Close: Set adoRecord = Nothing
'            adoHoraServ.Close: Set adoHoraServ = Nothing
'            frm_DatPar.Enabled = True
'            frm_ValOri.Enabled = True
'            fraFormaPago.Enabled = True
'            fraCtaFondo.Enabled = True
'        End If
'
'        adoComm.CommandText = "SELECT COUNT(*) CntReg FROM FMCTASYR WHERE COD_FOND='" & s_CodFon & "'"
'        Set adoresultAux = adoComm.Execute
'
'        If Not adoresultAux.EOF Then
'            If adoresultAux("CntReg") = 0 Then
'                MsgBox "No existe cuenta corriente para el fondo", vbCritical
'                adoresultAux.Close: Set adoresultAux = Nothing
'                frm_DatPar.Enabled = False: frm_ValOri.Enabled = False
'                fraFormaPago.Enabled = False
'                fraCtaFondo.Enabled = False
'                Exit Sub
'            Else
'                frm_DatPar.Enabled = True
'                frm_ValOri.Enabled = True
'                fraFormaPago.Enabled = True
'                fraCtaFondo.Enabled = True
'            End If
'        Else
'            MsgBox "No existe cuenta corriente para el fondo", vbCritical
'            adoresultAux.Close: Set adoresultAux = Nothing
'            frm_DatPar.Enabled = False: frm_ValOri.Enabled = False
'            fraFormaPago.Enabled = False
'            fraCtaFondo.Enabled = False
'            Exit Sub
'        End If
'        adoresultAux.Close
'
'
'        adoComm.CommandText = "Sp_INF_SelectSolici '01', '" & s_CodFon & "'"
'        Set adoresultAux = adoComm.Execute
'        s_CodMon = adoresultAux!COD_MONE
'        adoresultAux.Close: Set adoresultAux = Nothing
'
'        frm_DatPar.Caption = "Datos del Partícipe (Fondo " & IIf(s_CodMon = "S", "Soles)", "Dólares)")
'        frm_ValOri.Caption = "Valorización de Cuotas (" & IIf(s_CodMon = "S", "Soles :", "Dólares :") & ")"
'        pnl_descam(22).Caption = Trim$(Cmb_Cod_fond.Text)
'
'        adoComm.CommandText = "Sp_INF_SelectSolici '02', '" & s_CodFon & "'"
'        LCmbLoad adoComm.CommandText, cmbCtaFondo, aMapCodCtaFondo(), ""
'
'        If cmbCtaFondo.ListCount > 0 Then cmbCtaFondo.ListIndex = 0
'
'        adoComm.CommandText = "Sp_INF_SelectSolici '03', '" & s_CodFon & "'"
'        Set adoresult = adoComm.Execute
'        If Not adoresult.EOF Then
'            gstrpHorSusc = Format(adoresult!HOR_VCON, "HH:MM")
'            gtip_valu = adoresult!TIP_VALU
'        Else
'            gstrpHorSusc = "00:00"
'        End If
'
'        gstrflgvcon = ""
'        gstrtipval = ""
'
'        If gtip_valu = "1" Then '*** T-1 ***
'            n_CntCuo.Enabled = True: mhrMontSus.Enabled = True
'            pnl_descam(10).Visible = True: Lbl_Valcuo.Visible = True
'            If Format(tim_horsoli.Text, "HH:MM") <= Format(gstrpHorSusc, "HH:MM") Then
'                gstrflgvcon = "X"
'            Else
'                gstrflgvcon = ""
'            End If
'            strVC = "SC"
'        End If
'
'        If gtip_valu = "2" Then '*** T ***
'            If Format(tim_horsoli.Text, "HH:MM") > Format(gstrpHorSusc, "HH:MM") Then
'                n_CntCuo.Enabled = False: mhrMontSus.Enabled = True
'                lblDescrip(12).Visible = False: Lbl_Valcuo.Visible = False
'                If Format(tim_horsoli.Text, "HH:MM") <= Format(gstrpHorSusc, "HH:MM") Then
'                    gstrflgvcon = "X"
'                Else
'                    gstrflgvcon = ""
'                End If
'                strVC = "SD"
'            Else
'                n_CntCuo.Enabled = True: mhrMontSus.Enabled = True
'                lblDescrip(12).Visible = True: Lbl_Valcuo.Visible = True
'                If Format(tim_horsoli.Text, "HH:MM") <= Format(gstrpHorSusc, "HH:MM") Then
'                    gstrflgvcon = "X"
'                Else
'                    gstrflgvcon = ""
'                End If
'                strVC = "SC"
'            End If
'        End If
'
'        If gtip_valu = "3" Then '*** T + 1 ***
'            n_CntCuo.Enabled = False: mhrMontSus.Enabled = True
'            lblDescrip(12).Visible = True: Lbl_Valcuo.Visible = True
'            If Format(tim_horsoli.Text, "HH:MM") <= Format(gstrpHorSusc, "HH:MM") Then
'                gstrflgvcon = "X"
'            Else
'                gstrflgvcon = ""
'            End If
'            strVC = "SD"
'        End If
'
'        xFlgActCuota = "N"
'        xFlgActMonto = "N"
'
'        '*** Fecha disponible para el fondo ***
'        With adoComm
'            .CommandText = "Sp_INF_SelectSolici '04', '" & s_CodFon & "'"
'            Set adoresultAux = .Execute
'            If Not adoresultAux.EOF Then
'                'Dat_FchSoli.Text = FmtFec(adoresultAux!fch_cuot, "yyyymmdd", "win", res)
'                Dat_FchSoli.Value = Convertddmmyyyy(adoresultAux!fch_cuot)
'                '*** T-1 ***
'                If adoresult!TIP_VALU = "1" And gstrflgvcon = "X" Then
'                    Lbl_Valcuo.Caption = Format(adoresultAux!Val_cuo2, "#0.00000000")
'                    dblValcuot = Format(adoresultAux!Val_cuo2, "#0.00000000")
'                Else
'                    Lbl_Valcuo.Caption = Format(adoresultAux!Val_cuot, "#0.00000000")
'                    dblValcuot = Format(adoresultAux!Val_cuot, "#0.00000000")
'                End If
'
'                If Not IsNull(adoresultAux!CNT_INIC) Then
'                    n_CntIni = adoresultAux!CNT_INIC
'                Else
'                    n_CntIni = 0
'                End If
'                calcular
'            End If
'            adoresultAux.Close: Set adoresultAux = Nothing
'
'            .CommandText = "Sp_INF_SelectSolici '05', '" & s_CodFon & "'"
'            Set adoresultAux = .Execute
'
'            '** Ojo no utilizar la función IIf(,,) si alguno de los argumentos puede ser nulo
'            If Not adoresultAux.EOF Then
'                If Not IsNull(adoresultAux!TAS_SUSC) Then
'                    n_TasSus = adoresultAux!TAS_SUSC
'                Else
'                    n_TasSus = 0
'                End If
'
'                If Not IsNull(adoresultAux!VAL_MINI) Then
'                    n_CntMin = adoresultAux!VAL_MINI
'                Else
'                    n_CntMin = 0
'                End If
'
'                If Not IsNull(adoresultAux!MTO_PSUS) Then
'                    dblMtoPsus = adoresultAux!MTO_PSUS
'                Else
'                    dblMtoPsus = 0
'                End If
'
'                If Not IsNull(adoresultAux!CUO_SMIN) Then
'                    dblCuoSMin = adoresultAux!CUO_SMIN
'                Else
'                    dblCuoSMin = 0
'                End If
'
'                If Not IsNull(adoresultAux!MTO_SMIN) Then
'                    dblMtoSMin = adoresultAux!MTO_SMIN
'                Else
'                    dblMtoSMin = 0
'                End If
'
'                If Not IsNull(adoresultAux!TAS_FOND) Then
'                    dblTasFond = adoresultAux!TAS_FOND
'                Else
'                    dblTasFond = 0
'                End If
'
'                If n_CntIni > 0 Then
'                    If Not IsNull(adoresultAux!POR_MAXI) Then
'                        n_MaxFon = n_CntIni * (adoresultAux!POR_MAXI) / 100
'                    Else
'                        n_MaxFon = 0
'                    End If
'                    MsgBox "La cantidad máxima de cuotas a suscribir para el fondo " & Trim$(Cmb_Cod_fond) & " es " & Format(n_MaxFon, "###,##0.00000") & ".", vbExclamation
'                Else
'                    If Not IsNull(adoresultAux!POR_MAXI) Then
'                        MsgBox "El porcentaje máximo de participación para el fondo " & Trim$(Cmb_Cod_fond) & " es de " & adoresultAux!POR_MAXI & " %.", vbExclamation
'                    Else
'                        MsgBox "El fondo " & Trim(Cmb_Cod_fond) & " no tiene restricciones respecto al porcentage máximo de participación.", vbExclamation
'                    End If
'                End If
'            End If
'            adoresultAux.Close: Set adoresultAux = Nothing
'        End With
'        adoresult.Close: Set adoresult = Nothing
'        Call LimpiaScr
'
'        If (Cmb_Cod_prom.Text <> "") And (frm_DatPar.Enabled = True) Then
'            adoComm.CommandText = "Sp_INF_SelectSolici '06', '" & s_CodPro & "'"
'            Set adoresultAux = adoComm.Execute
'            If Not adoresultAux.EOF Then
'                v_Ite = Trim(adoresultAux!SUB_RIES)
'                cmbCtaFondo.ListIndex = LBsqIteArr(aMapCodCtaFondo(), v_Ite)
'            End If
'            adoresultAux.Close: Set adoresultAux = Nothing
'            Txt_CodUnico.SetFocus
'        Else
'            Cmb_Cod_prom.SetFocus
'        End If
'    End If
    
End Sub





Private Sub Cmd_Bsq_Click()

'    Dim adoCntCodUnico As New Recordset
'    Dim blnFlgCntCodUnico As Boolean
'
'    Dim s_TmpVar As String, strTipIden As String * 2, strNroIden As String * 15
'    Dim adoresultAux As New Recordset, adoresultAux1 As New Recordset
'
'    adoComm.CommandTimeout = 300
'
'    blnFlgCntCodUnico = False
'    If Trim(Txt_CodUnico.Text) <> "" Then
'        adoConn.CursorLocation = adUseClient
'        adoComm.CommandText = "Sp_INF_SelectSolici '07', '" & Trim(Txt_CodUnico.Text) & "'"
'        adoCntCodUnico.Open adoComm.CommandText, adoConn, adOpenForwardOnly, adLockBatchOptimistic
'
'        If adoCntCodUnico.RecordCount > 1 Then
'            blnFlgCntCodUnico = True
'            MsgBox "El código unico ingresado actualmente tiene más de un contrato. Por favor Seleccione al Partícipe en la pantalla de búsqueda", vbInformation, gstrNombreEmpresa
'        End If
'        adoCntCodUnico.Close: Set adoCntCodUnico = Nothing
'    End If
'
'
'    If (Trim(Txt_CodUnico) = "") Or (blnFlgCntCodUnico = True) Then
'        If Trim$(Txt_CodUnico) <> "" Then
'            frmINFBusqPart.chk_DocIde = True
'            frmINFBusqPart.txt_NroDoc = Mid(Trim(Txt_CodUnico.Text), 3, 10)
'        End If
'
'        frmINFBusqPart.Show 1
'        frmPROOpeSusc.Refresh
'        If Trim(gstrLlamaSoli) = "" Then
'            s_CodPar = ""
'            Call LimpiaScr
'            s_TmpVar = ""
'        Else
'            s_TmpVar = gstrLlamaSoli
'        End If
'        s_CodPar = Mid(s_TmpVar, 1, 15)
'        Me.MousePointer = vbHourglass
'        If Trim(s_TmpVar) <> "" Then
'            adoComm.CommandText = "Sp_INF_SelectSolici '08', '" & s_CodPar & "'"
'            Set adoresultAux = adoComm.Execute
'            Txt_CodUnico.Text = adoresultAux!COD_UNICO
'        Else
'            Txt_CodUnico.SetFocus: Me.MousePointer = vbDefault
'            Exit Sub
'        End If
'    Else
'        Me.MousePointer = vbHourglass
'        '** Datos del Partícipe
'        adoComm.CommandText = "Sp_INF_SelectSolici '09', '" & Trim(Txt_CodUnico.Text) & "'"
'        Set adoresultAux = adoComm.Execute
'        If adoresultAux.EOF Then
'            If MsgBox("No se ha encontrado al partícipe identificado con código único " & Txt_CodUnico.Text & Chr$(13) & Chr$(10) & "Desea realizar una búsqueda en el Maestro de Partícipes?", vbYesNo + vbQuestion + vbDefaultButton1) = vbYes Then
'                adoresultAux.Close: Set adoresultAux = Nothing
'                frmINFBusqPart.Show 1
'                frmPROOpeSusc.Refresh
'                If Trim$(gstrLlamaSoli) = "" Then
'                    s_CodPar = ""
'                    Call LimpiaScr
'                    s_TmpVar = ""
'                Else
'                    s_TmpVar = gstrLlamaSoli
'                End If
'                s_CodPar = Mid(s_TmpVar, 1, 15)
'                If Trim(s_TmpVar) <> "" Then
'                    adoComm.CommandText = "Sp_INF_SelectSolici '08', '" & s_CodPar & "'"
'                    Set adoresultAux = adoComm.Execute
'                    Txt_CodUnico.Text = adoresultAux!COD_UNICO
'                Else
'                    Txt_CodUnico.SetFocus: Me.MousePointer = vbDefault
'                    Exit Sub
'                End If
'            Else
'                Call LimpiaScr
'                Txt_CodUnico.SetFocus
'                Exit Sub: Me.MousePointer = vbDefault
'            End If
'        End If
'    End If
'    frm_DatGen.Caption = "Datos Generales " & IIf(adoresultAux!FLG_DIROK = "X", "(Dirección del Partícipe Errada)", "") & ":"
'    s_CodPar = Trim(adoresultAux!Cod_part): strTipIden = Trim(adoresultAux!TIP_IDEN)
'    s_CodProTit = Trim(adoresultAux!COD_PROM): strNroIden = Trim(adoresultAux!NRO_IDEN)
'    pnl_DscPar.Caption = Trim(adoresultAux!DSC_PART)
'    s_NroCust = IIf(IsNull(adoresultAux!NRO_CUST), "", adoresultAux!NRO_CUST)
'    chk_Flg_Cust.Enabled = True
'
'    If IsNull(adoresultAux!FLG_CUST) Or Trim(adoresultAux!FLG_CUST) = "" Then
'        chk_Flg_Cust.Value = False
'        chk_Flg_Cust.Enabled = False
'    Else
'        chk_Flg_Cust.Value = True
'        '*** Se desactiva DNA
'        chk_Flg_Cust.Enabled = False
'        '*** Fin
'    End If
'
'    If Not IsNull(adoresultAux!RUC_PART) Then
'        lbl_NroRuc = adoresultAux!RUC_PART
'        s_FlgFac = IIf(Len(Trim(adoresultAux!RUC_PART)) > 0, "F", "B")
'    Else
'        s_FlgFac = "B"
'        lbl_NroRuc = ""
'    End If
'
'    If adoresultAux!COD_PAIS = "001" Then
'        s_FlgExt = ""
'    Else
'        s_FlgExt = "X"
'    End If
'
'    s_ClsPer = adoresultAux!CLS_PART
'    adoresultAux.Close: Set adoresultAux = Nothing
'
'    adoComm.CommandText = "Sp_INF_SelectSolici '21', '" & s_CodPar & "', '" & s_CodFon & "'"
'    Set adoresultAux = adoComm.Execute
'    If Not adoresultAux.EOF Then
'        If adoresultAux!TAS_OPER >= 0 Then
'            n_TasSus.Text = adoresultAux!TAS_OPER
'        End If
'    End If
'    adoresultAux.Close: Set adoresultAux = Nothing
'
'    adoComm.CommandText = "Sp_INF_SelectSolici '40', '" & s_CodPar & "', '" & s_CodFon & "', '" & Format(Dat_FchSoli.Value, "yyyymmdd") & "'"
'    Set adoresultAux = adoComm.Execute
'    If Not adoresultAux.EOF Then
'        If adoresultAux!TAS_OPER >= 0 Then
'            n_TasSus.Text = adoresultAux!TAS_OPER
'        End If
'    End If
'    adoresultAux.Close: Set adoresultAux = Nothing
'
'    n_SldCuo = 0
'    n_SldMon = 0
'    With adoComm
'        '.CommandText = "    Sp_INF_SelectSolici '22', '" & Trim(strTipIden) & "', '" & Trim(strNroIden) & "'"
'        .CommandText = "    Sp_INF_SelectSolici '22', '" & s_CodPar & "'"
'        Set adoresultAux = .Execute
'        Do While Not adoresultAux.EOF
'            .CommandText = "Sp_INF_SelectSolici '23', '" & s_CodFon & "', '" & adoresultAux!Cod_part & "'"
'            Set adoresultAux1 = .Execute
'            If Not adoresultAux1.EOF Then
'                n_SldMon = n_SldMon + IIf(IsNull(adoresultAux1!TOTCUO), 0, adoresultAux1!TOTCUO)
'            End If
'            adoresultAux1.Close: Set adoresultAux1 = Nothing
'
'            adoresultAux.MoveNext
'        Loop
'        adoresultAux.Close: Set adoresultAux = Nothing
'    End With
'
'    lbl_sldmon = Format(n_SldMon, "###,##0.00000")
'    If n_SldMon > n_MaxFon Then
'        MsgBox "El Porcentaje máximo de suscripción excede en " & Format(n_SldMon - n_MaxFon, "0.00000") & " cuotas.", vbCritical
'        'Exit Sub
'    Else
'        n_MaxPar = n_MaxFon - n_SldMon
'    End If
'    FormPago(0).Value = True
'    If mhrMontSus.Enabled = True Then
'        mhrMontSus.SetFocus
'    End If
'    If n_CntCuo.Enabled = True Then
'        n_CntCuo.SetFocus
'    End If
'
'    Me.MousePointer = vbDefault

End Sub

Private Sub cboActivoAporte_Click()
    
    Dim adoRegistro As ADODB.Recordset
    
    Set adoRegistro = New ADODB.Recordset
    
    strCodActivoAporte = Valor_Caracter
    If cboActivoAporte.ListCount <= 0 Then Exit Sub
    strCodActivoAporte = Trim(arrActivoAporte(cboActivoAporte.ListIndex))
    
    strSql = "SELECT ValorNominal FROM ParticipeActivoAporte WHERE CodFile + '|' + CodAnalitica + '|' + CodCuentaActivo='" & strCodActivoAporte & "'"
    
    With adoComm
        
        .CommandText = strSql
        
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            While Not adoRegistro.EOF
                txtMontoNetoSolicitud.Text = adoRegistro.Fields("ValorNominal")
                Call txtMontoNetoSolicitud_KeyPress(13)
                adoRegistro.MoveNext
            Wend
        End If
        
    End With
    
End Sub

Private Sub cboAgencia_Click()

    Dim strSql As String, intRegistro   As Integer
    
    strCodAgencia = ""
    If cboAgencia.ListIndex < 0 Then Exit Sub
    
    strCodAgencia = Trim(arrAgencia(cboAgencia.ListIndex))
    
    'strSQL = "{ call up_ACSelDatosParametro(11,'" & strCodAgencia & "') }"
    'CargarControlLista strSQL, cboPromotor, arrPromotor(), Sel_Todos
    
    'If cboPromotor.ListCount > -1 Then cboPromotor.ListIndex = 0
    'intRegistro = ObtenerItemLista(arrPromotor(), gstrCodPromotor)
    'If intRegistro >= 0 Then cboPromotor.ListIndex = intRegistro
    
End Sub

Private Sub cboBanco_Click()

    strCodBanco = ""
    If cboBanco.ListIndex < 0 Then Exit Sub
    
    strCodBanco = Trim(arrBanco(cboBanco.ListIndex))
    
End Sub


Private Sub cboCertificado_Click()

    Dim strFechaSuscripcionCertificado As String
    Dim adoRegistro As New ADODB.Recordset
    
    strNumCertificado = Valor_Caracter
    If cboCertificado.ListIndex < 0 Then Exit Sub

    strNumCertificado = Mid(arrCertificado(cboCertificado.ListIndex), 1, 10)
    strFechaSuscripcionCertificado = Mid(arrCertificado(cboCertificado.ListIndex), 11)
    
    If gstrCodParticipe <> Valor_Caracter Then
        'lblCuotasDisponibles.Caption = CStr(ObtenerCuotasParticipe(gstrCodParticipe, strNumCertificado, strCodFondo, gstrCodAdministradora, Valor_Caracter, Valor_Caracter))
        
        'Obtiene el tipo de comision de rango, si existe
        With adoComm
            
            .CommandText = "SELECT CodComision FROM FondoComision WHERE CodFondo = '" & strCodFondo & "' AND " & _
                "CodAdministradora = '" & gstrCodAdministradora & "' AND CodOperacion = '" & strCodTipoOperacion & "' AND " & _
                "IndRango = 'X' AND IndVigente = 'X'"
            Set adoRegistro = .Execute
            
            strCodComision = Valor_Caracter
            If Not adoRegistro.EOF Then
                strCodComision = Trim(adoRegistro("CodComision"))
            End If
            adoRegistro.Close
            
            'Calcula la comision rescate por rango
            dblTasaRescate = ObtenerComisionParticipacionRango(strCodComision, strCodFondo, gstrCodAdministradora, strFechaSuscripcionCertificado, Convertyyyymmdd(lblFechaSolicitud.Caption))
                
            'Obtener comisionista
        
            
        End With
        
        Call ObtenerCuotasTotalesParticipe(strNumCertificado)
        
        If strCodTipoOperacion = Codigo_Operacion_Rescate And gstrTipoAdministradora = Codigo_Tipo_Fondo_Inversion Then
            txtCantCuotasSolicitud.Text = CStr(lblCuotasDisponibles.Caption)
        End If
        
        'lblCuotasBloqueadas.Caption = CStr(ObtenerCuotasParticipe(gstrCodParticipe, strNumCertificado, strCodFondo, gstrCodAdministradora, Valor_Caracter, Valor_Indicador))
    End If


End Sub

Private Sub cboComisionista_Click()

    strCodComisionista = Valor_Caracter
    numSecCondicion = 0
    
    If cboComisionista.ListIndex < 0 Then Exit Sub
    
    strCodComisionista = Mid$(arrComisionista(cboComisionista.ListIndex), 1, 8)
    numSecCondicion = Mid$(arrComisionista(cboComisionista.ListIndex), 9)
End Sub

Private Sub cboCuenta_Click()

    If cboCuenta.ListIndex < 0 Then Exit Sub
    
    strCodBanco = Left(arrCuenta(cboCuenta.ListIndex), 8)
    strTipoCuenta = Trim(Mid(arrCuenta(cboCuenta.ListIndex), 9, 2))
    strNumCuenta = Trim(Mid(arrCuenta(cboCuenta.ListIndex), 11, 30))
    
End Sub

Private Sub cboCuentaFondo_Click()

    strCodBancoDestino = Valor_Caracter
    strTipoCuentaDestino = Valor_Caracter
    strNumCuentaDestino = Valor_Caracter
    
    If cboCuentaFondo.ListIndex < 0 Then Exit Sub
    
    strCodBancoDestino = Left(arrCuentaFondo(cboCuentaFondo.ListIndex), 8)
    strTipoCuentaDestino = Trim(Mid(arrCuentaFondo(cboCuentaFondo.ListIndex), 9, 2)) '20141201_JJCC
    strNumCuentaDestino = Trim(Mid(arrCuentaFondo(cboCuentaFondo.ListIndex), 11, 30)) '20141201_JJCC
    
End Sub

Private Sub cboEjecutivo_Click()

    Dim adoRegistro As ADODB.Recordset
    
    strCodEjecutivo = Valor_Caracter
    If cboEjecutivo.ListIndex < 0 Then Exit Sub
    
    strCodEjecutivo = Trim(arrEjecutivo(cboEjecutivo.ListIndex))
    
    'If cboFondo.ListIndex > 0 And cboTipoOperacion.ListIndex > 0 And Trim(txtNumPapeleta.Text) <> Valor_Caracter Then
    If cboFondo.ListIndex > 0 And cboTipoOperacion.ListIndex > 0 Then
        Call Habilita
        If strHoraCorte = "00:00" Then Call Deshabilita
        If chkPagoParcial.Value Then
            If dblPorcenPago = 0 And curMontoPago = 0 Then Call Deshabilita
        End If
    Else
        Call Deshabilita
    End If
    
    Set adoRegistro = New ADODB.Recordset
    
    adoComm.CommandText = "SELECT CodSucursal,CodAgencia FROM InstitucionPersona " & _
    "WHERE TipoPersona='01' AND CodPersona='" & strCodEjecutivo & "'"
    Set adoRegistro = adoComm.Execute
    
    If Not adoRegistro.EOF Then
        strCodSucursalSolicitud = Trim(adoRegistro("CodSucursal"))
        strCodAgenciaSolicitud = Trim(adoRegistro("CodAgencia"))
    End If
    adoRegistro.Close: Set adoRegistro = Nothing
    
End Sub

Private Sub cboFondo_Click()

    Dim adoRegistro As ADODB.Recordset, adoTemporal As ADODB.Recordset
    
    strCodFondo = Valor_Caracter
    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim(arrFondo(cboFondo.ListIndex))
   
    cboComisionista.Clear
    
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Fecha Vigente, Valor Cuota, Valor Cuota T-1, Moneda y Cantidad Inicial de Cuotas del Fondo ***
        .CommandText = "{ call up_ACSelDatosParametro(23,'" & strCodFondo & "','" & gstrCodAdministradora & "') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            lblFechaSolicitud.Caption = CStr(adoRegistro("FechaCuota"))
            dblValorCuota = Round(CDbl(adoRegistro("ValorCuotaInicial")), Decimales_ValorCuota)
            'lblValorCuota.Caption = CStr(dblValorCuota)
            txtValorCuota.Text = CStr(dblValorCuota)
            
            strCodMonedaFondo = Trim(adoRegistro("CodMoneda"))
            dblCantCuotaInicio = CDbl(adoRegistro("CantCuotaInicio"))
                        
            gdatFechaActual = adoRegistro("FechaCuota")
            gstrFechaActual = Convertyyyymmdd(gdatFechaActual)
            frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
            
            Call ObtenerParametrosFondoOP
            
            '*** Se inició la Etapa PreOperativa ? ***
            If Trim(lblFechaSolicitud.Caption) <> "dd/mm/yyyy" And datFechaEtapaPreOperativa <> CDate(Valor_Fecha) Then
                If CDate(lblFechaSolicitud.Caption) < datFechaEtapaPreOperativa Then
                    MsgBox "La Fecha de Inicio de la Etapa PreOperativa es" & Space(1) & CStr(datFechaEtapaPreOperativa), vbCritical
                    Exit Sub
                End If
            End If
            
            If chkPagoParcial.Value Then
                Set adoTemporal = New ADODB.Recordset
                
                .CommandText = "SELECT PorcenPago,MontoPago FROM FondoPagoSuscripcion " & _
                    "WHERE FechaDesde>='" & Convertyyyymmdd(adoRegistro("FechaCuota")) & "' AND FechaHasta>'" & Convertyyyymmdd(adoRegistro("FechaCuota")) & "' AND " & _
                    "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' ORDER BY NumSecuencial"
                Set adoTemporal = .Execute
                
                If Not adoTemporal.EOF Then
                    dblPorcenPago = CDbl(adoTemporal("PorcenPago"))
                    curMontoPago = CCur(adoTemporal("MontoPago"))
                End If
                adoTemporal.Close: Set adoTemporal = Nothing
                
                If dblPorcenPago = 0 And curMontoPago = 0 Then
                    MsgBox "No ha definido el cronograma de pago de cuotas parciales", vbCritical, Me.Caption
                    Exit Sub
                End If
            End If
            
            If cboTipoOperacion.ListIndex > -1 And cboFondo.ListIndex > 0 Then Call ObtenerCuentasFondo
            
            If strHoraCorte = "00:00" Then Exit Sub
        
            If strCodTipoOperacion = Codigo_Operacion_Suscripcion Then
                If datFechaEtapaOperativa <> CDate(Valor_Fecha) Then
                    If CDate(lblFechaSolicitud.Caption) >= datFechaEtapaOperativa Then
                        If dblCantCuotaInicio > 0 Then
                            If dblPorcenMaxParticipe > 0 Then
                                dblCantMaxCuotaFondo = dblCantCuotaInicio * dblPorcenMaxParticipe * 0.01
                            Else
                                dblCantMaxCuotaFondo = 0
                            End If
                            MsgBox "La cantidad máxima de cuotas a suscribir para el fondo " & Trim(cboFondo.Text) & " es " & Format(dblCantMaxCuotaFondo, "###,##0.00") & " ...", vbInformation
                        Else
                            MsgBox "El fondo " & Trim(cboFondo.Text) & " no tiene restricciones respecto al porcentaje máximo de participación" & _
                                vbNewLine & vbNewLine & "Aún no existen cuotas suscritas.", vbInformation
                        End If
                    End If
                End If
                
                If gstrTipoAdministradora = Codigo_Tipo_Fondo_Inversion Then
                    Call ObtenerFechaValorCuota
                End If
            End If
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    'If cboTipoOperacion.ListIndex > -1 And cboEjecutivo.ListIndex > -1 And Trim(txtNumPapeleta.Text) <> Valor_Caracter Then
    If cboTipoOperacion.ListIndex > -1 And cboEjecutivo.ListIndex > -1 Then
        Call Habilita
    Else
        Call Deshabilita
    End If
    
End Sub

Private Sub cboFondoSolicitud_Click()

    Dim adoRegistro     As ADODB.Recordset
    
    strCodFondoSolicitud = Valor_Caracter
    If cboFondoSolicitud.ListIndex < 0 Then Exit Sub
    
    strCodFondoSolicitud = Trim(arrFondoSolicitud(cboFondoSolicitud.ListIndex))
    
    Set adoRegistro = New ADODB.Recordset
    '*** Fecha Vigente, Valor Cuota, Valor Cuota T-1, Moneda y Cantidad Inicial de Cuotas del Fondo ***
    adoComm.CommandText = "{ call up_ACSelDatosParametro(23,'" & strCodFondoSolicitud & "','" & gstrCodAdministradora & "') }"
    Set adoRegistro = adoComm.Execute
        
    If Not adoRegistro.EOF Then
        dtpFechaDesde.Value = adoRegistro("FechaCuota")
        dtpFechaHasta.Value = dtpFechaDesde.Value
        gdatFechaActual = adoRegistro("FechaCuota")
        frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
    End If
    adoRegistro.Close: Set adoRegistro = Nothing
    
End Sub

'Private Sub cboNumCuenta_Click()
'
'    strNumCuenta = Valor_Caracter
'    If cboNumCuenta.ListIndex < 0 Then Exit Sub
'
'    strNumCuenta = Trim(arrNumCuenta(cboNumCuenta.ListIndex))
'
'End Sub

Private Sub cboPromotor_Click()

    strCodPromotor = Valor_Caracter
    If cboPromotor.ListIndex < 0 Then Exit Sub
    
    strCodPromotor = Trim(arrPromotor(cboPromotor.ListIndex))
    
End Sub

Private Sub cboSucursal_Click()

    Dim strSql As String, intRegistro   As Integer
    
    strCodSucursal = Valor_Caracter
    If cboSucursal.ListIndex < 0 Then Exit Sub
    
    strCodSucursal = Trim(arrSucursal(cboSucursal.ListIndex))
    
    strSql = "{ call up_ACSelDatosParametro(10,'" & strCodSucursal & "') }"
    CargarControlLista strSql, cboAgencia, arrAgencia(), Sel_Todos
    
    If cboAgencia.ListCount > -1 Then cboAgencia.ListIndex = 0
    intRegistro = ObtenerItemLista(arrAgencia(), gstrCodAgencia)
    If intRegistro >= 0 Then cboAgencia.ListIndex = intRegistro
    
End Sub

Private Sub cboTipoAporte_Click()
    strTipoAporte = Valor_Caracter
    If cboTipoAporte.ListCount <= 0 Then Exit Sub
    strTipoAporte = Trim(arrTipoAporte(cboTipoAporte.ListIndex))
    
    If strTipoAporte = Tipo_Aporte_Dinerario Then
        lblDescrip(40).Visible = False
        cboActivoAporte.Visible = False
        fraFormaPago.Enabled = True
    ElseIf strTipoAporte = Tipo_Aporte_NoDinerario Then
        If gstrCodParticipe <> Valor_Caracter Then
            If cboActivoAporte.ListCount > 0 Then
                lblDescrip(40).Visible = True
                cboActivoAporte.Visible = True
            End If
        Else
            MsgBox "No ha seleccionado el participe", vbCritical, Me.Caption
        End If
        fraFormaPago.Enabled = False
    End If
    
End Sub

Private Sub cboTipoFormaPago_Click()

    Dim intRegistro As Integer
    
    strCodTipoFormaPago = Valor_Caracter
    strCodBanco = Valor_Caracter
    strTipoCuenta = Valor_Caracter
    strNumCuenta = Valor_Caracter
    
    If cboTipoFormaPago.ListIndex < 0 Then Exit Sub
    
    strCodTipoFormaPago = Trim(arrTipoFormaPago(cboTipoFormaPago.ListIndex))
    
    If strCodCliente <> Valor_Caracter Then
        Call ObtenerCuentasParticipe
        
        Select Case strCodTipoFormaPago
            Case Codigo_FormaPago_Transferencia_Mismo_Banco, Codigo_FormaPago_Transferencia_Otro_Banco, Codigo_FormaPago_Transferencia_Exterior
                If indExisteCuentaParticipe = Valor_Indicador Then
                    cboCuenta.Enabled = True
                    cboCuenta.Visible = True
                    cboBanco.Enabled = False
                    cboBanco.Visible = False
                    txtNumCheque.Enabled = False
                    txtNumCheque.Visible = False
                    txtNumCheque.Text = Valor_Caracter
                    lblDescrip(18).Caption = "Cuenta"
                    lblDescrip(18).Visible = True
                    lblDescrip(19).Caption = Valor_Caracter
                    lblDescrip(19).Visible = False
                Else
                    MsgBox "El Partícipe no tiene cuentas definidas...", vbCritical, gstrNombreEmpresa
                    
                    intRegistro = ObtenerItemLista(arrTipoFormaPago(), Codigo_FormaPago_Efectivo)
                    If intRegistro >= 0 Then cboTipoFormaPago.ListIndex = intRegistro
                End If
            Case Codigo_FormaPago_Cuenta
                cboCuenta.Enabled = False
                cboCuenta.Visible = False
                cboBanco.Enabled = False
                cboBanco.Visible = False
                txtNumCheque.Visible = False
                txtNumCheque.Enabled = False
                txtNumCheque.Text = Valor_Caracter
                lblDescrip(18).Caption = Valor_Caracter
                lblDescrip(18).Visible = False
                lblDescrip(19).Caption = Valor_Caracter
                lblDescrip(19).Visible = False
            Case Codigo_FormaPago_Cheque
                cboCuenta.Enabled = False
                cboCuenta.Visible = False
                cboBanco.Enabled = True
                cboBanco.Visible = True
                txtNumCheque.Visible = True
                txtNumCheque.Enabled = True
                lblDescrip(18).Caption = "Banco"
                lblDescrip(18).Visible = True
                lblDescrip(19).Caption = "Núm.Cheque"
                lblDescrip(19).Visible = True
            Case Codigo_FormaPago_Efectivo
                cboCuenta.Enabled = False
                cboCuenta.Visible = False
                cboBanco.Enabled = False
                cboBanco.Visible = False
                txtNumCheque.Enabled = False
                txtNumCheque.Visible = False
                txtNumCheque.Text = Valor_Caracter
                lblDescrip(18).Caption = Valor_Caracter
                lblDescrip(18).Visible = False
                lblDescrip(19).Caption = Valor_Caracter
                lblDescrip(19).Visible = False
        End Select
    End If
End Sub

Private Sub cboTipoOperacion_Click()

    strCodTipoOperacion = Valor_Caracter
    If cboTipoOperacion.ListIndex < 0 Then Exit Sub
    
    strCodTipoOperacion = Trim(arrTipoOperacion(cboTipoOperacion.ListIndex))
    
    If strCodTipoOperacion = Codigo_Operacion_Rescate Then
        If gstrTipoAdministradora = Codigo_Tipo_Fondo_Inversion Then
            fraFormaPago.Caption = "Forma Pago Separación"
        Else
            fraFormaPago.Caption = "Forma Pago Rescate"
        End If
        lblDescrip(31).Caption = "Valor de Redención"
        dtpFechaValorCuota.Enabled = True
        cboCertificado.Visible = True
        'lblDescrip(8).Visible = True
    
        Call ObtenerCertificadosParticipe
    
    Else
        dtpFechaValorCuota.Enabled = False
        fraFormaPago.Caption = "Forma Pago Suscripción"
        lblDescrip(31).Caption = "Valor de Suscripción"
        cboCertificado.Visible = False
    End If
    
    If gstrTipoAdministradora = Codigo_Tipo_Fondo_Mutuo Then
        lblDescrip(31).Visible = False: dtpFechaValorCuota.Visible = False
        chkPagoParcial.Visible = False
    End If
    
    Call ObtenerParametrosFondoOP
    
    If cboTipoOperacion.ListIndex >= 0 And cboFondo.ListIndex > 0 Then Call ObtenerCuentasFondo
    
    'If cboFondo.ListIndex > 0 And cboEjecutivo.ListIndex > -1 And Trim(txtNumPapeleta.Text) <> Valor_Caracter Then
    If cboFondo.ListIndex > 0 And cboEjecutivo.ListIndex > -1 Then
        Call Habilita
    Else
        Call Deshabilita
    End If
    
    
End Sub

Private Sub cboTipoSolicitud_Click()

    strCodTipoSolicitud = Valor_Caracter
    If cboTipoSolicitud.ListIndex < 0 Then Exit Sub
    
    strCodTipoSolicitud = Trim(arrTipoSolicitud(cboTipoSolicitud.ListIndex))
    
End Sub

Private Sub chkPagoParcial_Click()

    If chkPagoParcial.Value Then
        strIndPagoParcial = Valor_Indicador
    Else
        strIndPagoParcial = Valor_Caracter
    End If
    
End Sub


Private Sub cmdBusqueda_Click()

'    gstrFormulario = "frmSolicitudParticipe"
'    frmBusquedaParticipe.Show vbModal
    
    Dim sSql As String
    Dim intRegistro As Integer
   
    Screen.MousePointer = vbHourglass
   
    Dim frmBus As frmBuscar
    
    Set frmBus = New frmBuscar
    
    With frmBus.TBuscarRegistro1
           
        .ADOConexion = adoConn
        .ADOConexion.CommandTimeout = 0

        .iTipoGrilla = 2
        
        frmBus.Caption = " Relación de Participes"
        .sSql = "{ call up_ACSelDatos(30) }"
        
        .OutputColumns = "1,2,3,4,5,6,7,8,9,10,11"
        .HiddenColumns = "1,2,5,6,7,10,11"
        
        .BuscarTabla
        
        Screen.MousePointer = vbNormal
        frmBus.Show 1
       
        If .iParams.Count = 0 Then Exit Sub
        
        If .iParams(1).Valor <> "" Then
            gstrCodParticipe = Trim(.iParams(1).Valor)
            txtTipoDocumento.Text = Trim(.iParams(2).Valor)
            txtNumDocumento.Text = Trim(.iParams(3).Valor)
            lblDescripParticipe.Caption = Trim(.iParams(4).Valor)
            lblDescripTipoParticipe.Caption = Trim(.iParams(7).Valor)
            strCodTipoDocumento = Trim(.iParams(6).Valor)
            txtTitularSolicitante.Text = Trim(.iParams(8).Valor)
            strCodCliente = Trim(.iParams(9).Valor)
            'strCodComisionista = Trim(.iParams(10).Valor)
            
'            strSql = "SELECT CodFile + '|' + CodAnalitica CODIGO, DescripActivo DESCRIP, CodCuentaActivo CodCuentaActivo " & _
'                    "FROM ParticipeActivoAporte WHERE CodParticipe='" & gstrCodParticipe & "' AND IndVigente='X' AND " & _
'                    "CodFile + '|' + CodAnalitica NOT IN (SELECT CodFileActivo + '|' + CodAnaliticaActivo FROM ParticipeSolicitud WHERE EstadoSolicitud<>'" & Estado_Solicitud_Anulada & "')"
'            CargarControlLista strSql, cboActivoAporte, arrActivoAporte(), Sel_Defecto
            
            strSql = "SELECT CodFile + '|' + CodAnalitica + '|' + CodCuentaActivo CODIGO, DescripActivo DESCRIP " & _
                    "FROM ParticipeActivoAporte WHERE CodParticipe='" & gstrCodParticipe & "' AND IndVigente='X' AND " & _
                    "CodFile + '|' + CodAnalitica NOT IN (SELECT CodFileActivo + '|' + CodAnaliticaActivo FROM ParticipeSolicitud WHERE EstadoSolicitud<>'" & Estado_Solicitud_Anulada & "')"
            CargarControlLista strSql, cboActivoAporte, arrActivoAporte(), Sel_Defecto
            
            If cboActivoAporte.ListCount > 0 Then cboActivoAporte.ListIndex = 0
            
            If strTipoAporte = Tipo_Aporte_NoDinerario Then
                If cboActivoAporte.ListCount > 1 Then
                    lblDescrip(40).Visible = True
                    cboActivoAporte.Visible = True
                End If
            End If
        Else
            gstrCodParticipe = Valor_Caracter
            txtTipoDocumento.Text = Valor_Caracter
            txtNumDocumento.Text = Valor_Caracter
            lblDescripParticipe.Caption = Valor_Caracter
            lblDescripTipoParticipe.Caption = Valor_Caracter
            txtTitularSolicitante.Text = Valor_Caracter
        End If
            
       
    End With
    
    Set frmBus = Nothing
     
    Call ObtenerCuotasTotalesParticipe
    
    Call ObtenerCertificadosParticipe
     
    'Obtener lista de comisionistas
    strSql = "{ call up_ACLstFondoComisionistaContraparte('" & gstrCodFondoContable & "','" & gstrCodAdministradora & "','" & _
                Codigo_Tipo_Comisionista_Participe & "','" & Codigo_Tipo_Persona_Participe & "','" & "" & "','" & _
                strCodMonedaFondo & "','" & gstrFechaActual & "') }"
    CargarControlLista strSql, cboComisionista, arrComisionista(), Valor_Caracter
    If cboComisionista.ListCount = 1 Then
        cboComisionista.ListIndex = 0
    End If
    
    fraFormaPago.Enabled = True
    intRegistro = ObtenerItemLista(arrTipoFormaPago(), Codigo_FormaPago_Efectivo)
    If intRegistro >= 0 Then cboTipoFormaPago.ListIndex = intRegistro
    
End Sub
Private Sub ObtenerCertificadosParticipe()
    
    If strCodTipoOperacion = Codigo_Operacion_Rescate And gstrCodParticipe <> "" Then
        strSql = "{ call up_ACSelDatosParametro(51,'" & gstrCodParticipe & "','" & strCodCliente & "','" & strCodFondo & "','" & gstrCodAdministradora & "') }"
        CargarControlLista strSql, cboCertificado, arrCertificado(), Valor_Caracter
    End If

End Sub
Private Sub ObtenerCuotasTotalesParticipe(Optional pstrNumCertificado As String = "")
    
    Dim strCodUnico As String 'AGREGE ESTO PARA QUE NO SE CAIGA EL SISTEMA '¿"AGREGE"? ¿No será: "AGREGUÉ"?
    If gstrCodParticipe <> "" Then
        lblCuotasDisponibles.Caption = CStr(ObtenerCuotasParticipe(gstrCodParticipe, pstrNumCertificado, strCodFondo, gstrCodAdministradora, Valor_Caracter, Valor_Caracter))
        lblCuotasBloqueadas.Caption = CStr(ObtenerCuotasParticipe(gstrCodParticipe, pstrNumCertificado, strCodFondo, gstrCodAdministradora, Valor_Caracter, Valor_Indicador))
    End If

End Sub

Private Sub ObtenerDatosParticipe()

    Dim adoRegistro As ADODB.Recordset

    Set adoRegistro = New ADODB.Recordset
    adoRegistro.CursorLocation = adUseClient
    adoRegistro.CursorType = adOpenStatic

    adoComm.CommandText = "SELECT PC.CodParticipe,AP1.DescripParametro TipoIdentidad,PCD.NumIdentidad,DescripParticipe,FechaIngreso,PCD.TipoIdentidad CodIdentidad,PC.TipoMancomuno, AP2.DescripParametro DescripMancomuno " & _
        "FROM ParticipeContratoDetalle PCD JOIN ParticipeContrato PC " & _
        "ON(PCD.CodParticipe=PC.CodParticipe AND PCD.TipoIdentidad='" & strCodTipoDocumento & "' AND PCD.NumIdentidad='" & Trim(txtNumDocumento.Text) & "') " & _
        "JOIN AuxiliarParametro AP1 ON(AP1.CodParametro=PCD.TipoIdentidad AND CodTipoParametro='TIPIDE') " & _
        "JOIN AuxiliarParametro AP2 ON(AP2.CodParametro=PC.TipoMancomuno AND AP2.CodTipoParametro='TIPMAN')"
    adoRegistro.Open adoComm.CommandText, adoConn

    If Not adoRegistro.EOF Then
        If adoRegistro.RecordCount > 1 Then
            gstrFormulario = "frmSolicitudParticipe"
            'frmBusquedaParticipe.optCriterio(1).Value = vbChecked
            'frmBusquedaParticipe.txtNumDocumento = Trim(txtNumDocumento.Text)
            'Call frmBusquedaParticipe.Buscar
            'frmBusquedaParticipe.Show vbModal
        Else
            gstrCodParticipe = Trim(adoRegistro("CodParticipe"))
            lblDescripTipoParticipe.Caption = Trim(adoRegistro("DescripMancomuno"))
            lblDescripParticipe.Caption = Trim(adoRegistro("DescripParticipe"))
        End If
    End If
    adoRegistro.Close
            
End Sub

Private Sub dtpFechaDesde_Click()

    If IsNull(dtpFechaDesde.Value) Then
        dtpFechaHasta.Value = Null
   ' Else
   '     dtpFechaDesde.Value = gdatFechaActual
   '     dtpFechaHasta.Value = dtpFechaDesde.Value
    End If
    
End Sub


Private Sub dtpFechaHasta_Click()

    If IsNull(dtpFechaHasta.Value) Then
        dtpFechaDesde.Value = Null
    'Else
        'dtpFechaDesde.Value = gdatFechaActual
        'dtpFechaHasta.Value = dtpFechaDesde.Value
    End If
    
End Sub


Private Sub dtpFechaValorCuota_Change()
        
    If gstrTipoAdministradora = Codigo_Tipo_Fondo_Inversion Then
        Dim adoRegistro     As ADODB.Recordset
        Dim strFechaInicio  As String, strFechaFin  As String
        
        strFechaInicio = Convertyyyymmdd(dtpFechaValorCuota.Value)
        strFechaFin = Convertyyyymmdd(DateAdd("d", 1, dtpFechaValorCuota.Value))
        
        Set adoRegistro = New ADODB.Recordset
        With adoComm
            .CommandText = "SELECT ValorCuotaInicial,ValorCuotaFinal FROM FondoValorCuota " & _
                "WHERE (FechaCuota>='" & strFechaInicio & "' AND FechaCuota<'" & strFechaFin & "') AND " & _
                "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
            Set adoRegistro = .Execute
            
            If Not adoRegistro.EOF Then
                If blnValorConocido Then
                    'lblValorCuota.Caption = CStr(adoRegistro("ValorCuotaInicial"))
                    txtValorCuota.Text = CStr(adoRegistro("ValorCuotaInicial"))
                    dblValorCuota = Round(CDbl(adoRegistro("ValorCuotaInicial")), Decimales_ValorCuota)
                Else
                    'lblValorCuota.Caption = CStr(adoRegistro("ValorCuotaFinal"))
                    txtValorCuota.Text = CStr(adoRegistro("ValorCuotaFinal"))
                    dblValorCuota = Round(CDbl(adoRegistro("ValorCuotaFinal")), Decimales_ValorCuota)
                End If
            End If
            adoRegistro.Close: Set adoRegistro = Nothing
        End With
    End If
    
End Sub

Private Sub Form_Activate()

    Call CargarReportes
    
End Sub

Private Sub Form_Deactivate()

    ReDim garrTipoDocumento(0)
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
        
    Call ValidaExisteTipoCambio(Codigo_TipoCambio_SBS, gstrFechaActual)
    
End Sub

Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Lista de Solicitudes"
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Text = "Imprimir Solicitud de Suscripción"
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo3").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo3").Text = "Imprimir Solicitud de Rescate"
    
End Sub


Private Sub DarFormato()

    Dim intCont As Integer
    Dim elemento As Object
    
    For intCont = 0 To (lblDescrip.Count - 1)
       ' Call FormatoEtiqueta(lblDescrip(intCont), vbLeftJustify)
    Next
    
    For Each elemento In Me.Controls
    
        If TypeOf elemento Is TDBGrid Then
            Call FormatoGrilla(elemento)
        End If
    
    Next
            
End Sub

Private Sub CargarListas()

    Dim intRegistro As Integer
    
    '*** Fondos ***
    strSql = "{ call up_ACSelDatosParametro(74,'" & gstrCodAdministradora & "','" & gstrCodFondoContable & "') }"
    CargarControlLista strSql, cboFondoSolicitud, arrFondoSolicitud(), Valor_Caracter
    CargarControlLista strSql, cboFondo, arrFondo(), Sel_Defecto
    
    If cboFondoSolicitud.ListCount > 0 Then cboFondoSolicitud.ListIndex = 0
    
    '*** Tipo Solicitud/Operación ***
    strSql = "SELECT CodTipoSolicitud CODIGO,DescripTipoSolicitud DESCRIP FROM TipoSolicitud WHERE CodCorto<>'T' and CodCorto<>'M' and CodCorto<>'O' ORDER BY DescripTipoSolicitud"
    CargarControlLista strSql, cboTipoSolicitud, arrTipoSolicitud(), Sel_Todos
    CargarControlLista strSql, cboTipoOperacion, arrTipoOperacion(), Valor_Caracter
    
    If cboTipoSolicitud.ListCount > 0 Then cboTipoSolicitud.ListIndex = 0
                        
    '*** Sucursal ***
    strSql = "{ call up_ACSelDatos(15) }"
    CargarControlLista strSql, cboSucursal, arrSucursal(), Sel_Todos
    
    If cboSucursal.ListCount > 0 Then cboSucursal.ListIndex = 0
    intRegistro = ObtenerItemLista(arrSucursal(), gstrCodSucursal)
    If intRegistro >= 0 Then cboSucursal.ListIndex = 0
    'If intRegistro >= 0 Then cboSucursal.ListIndex = intRegistro
    
    '*** Operador ***
    'strSQL = "SELECT CodPersona CODIGO,DescripPersona DESCRIP FROM InstitucionPersona WHERE TipoPersona='01' ORDER BY DescripPersona"
    'CargarControlLista strSQL, cboEjecutivo, arrEjecutivo(), Valor_Caracter
    
    strSql = "{ call up_ACSelDatos(41) }"
    CargarControlLista strSql, cboPromotor, arrPromotor(), Sel_Defecto
    
    If cboPromotor.ListCount > 0 Then cboPromotor.ListIndex = 0
    intRegistro = ObtenerItemLista(arrPromotor(), gstrCodPromotor)
    If intRegistro >= 0 Then cboPromotor.ListIndex = 0
    'If intRegistro >= 0 Then cboPromotor.ListIndex = intRegistro
    
    strSql = "{ call up_ACSelDatos(41) }"
    CargarControlLista strSql, cboEjecutivo, arrEjecutivo(), Sel_Defecto
    
    If cboEjecutivo.ListCount > 0 Then cboEjecutivo.ListIndex = 0
    intRegistro = ObtenerItemLista(arrEjecutivo(), gstrCodPromotor)
    If intRegistro >= 0 Then cboEjecutivo.ListIndex = intRegistro
    
    '*** Tipo Documento Identidad ***
'    strSQL = "{ call up_ACSelDatos(11) }"
'    CargarControlLista strSQL, cboTipoDocumento, garrTipoDocumento(), Sel_Defecto
    
    '*** Forma de Pago ***
    strSql = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='MEDPAG' ORDER BY DescripParametro"
    CargarControlLista strSql, cboTipoFormaPago, arrTipoFormaPago(), Sel_Defecto
    
    '*** Banco ***
    strSql = "{ call up_ACSelDatos(22) }"
    CargarControlLista strSql, cboBanco, arrBanco(), Sel_Defecto
    
    '*** Tipo Aporte ***
    strSql = "SELECT CodParametro CODIGO, DescripParametro DESCRIP From AuxiliarParametro WHERE CodTipoParametro='TACPIN'"
    CargarControlLista strSql, cboTipoAporte, arrTipoAporte(), Valor_Caracter
    
    If cboTipoAporte.ListCount > 0 Then cboTipoAporte.ListIndex = 0
    
    'Obtener lista de comisionistas
    strSql = "{ call up_ACLstFondoComisionistaContraparte('" & gstrCodFondoContable & "','" & gstrCodAdministradora & "','" & _
                Codigo_Tipo_Comisionista_Inversion & "','" & Codigo_Tipo_Persona_Participe & "','" & "" & "','" & _
                gstrCodMoneda & "','" & gstrFechaActual & "') }"
    CargarControlLista strSql, cboComisionista, arrComisionista(), Valor_Caracter
    If cboComisionista.ListCount = 1 Then
        cboComisionista.ListIndex = 0
    End If
    
End Sub

Private Sub InicializarValores()

    Dim adoRegistro As ADODB.Recordset
    Dim intCont     As Integer
    
    '*** Valores Iniciales ***
    strEstado = Reg_Defecto
    blnValorConocido = False
    tabSolicitud.Tab = 0

    dtpFechaDesde.Value = gdatFechaActual
    dtpFechaHasta.Value = gdatFechaActual
    datFechaEtapaOperativa = CDate(Valor_Fecha)
    datFechaEtapaPreOperativa = CDate(Valor_Fecha)
    
    Set adoRegistro = New ADODB.Recordset
    
    intCont = 0
    ReDim arrLeyendaSuscripcion(intCont)
    ReDim arrLeyendaRescate(intCont)
    
    With adoComm
        '*** Suscripción ***
        .CommandText = "SELECT RTRIM(TSD.CodCorto) + space(1) + '=' + space(1) + RTRIM(TS.DescripTipoSolicitud) + space(1) + RTRIM(TSD.DescripDetalleTipoSolicitud)  ValorLeyenda " & _
            "FROM TipoSolicitud TS JOIN TipoSolicitudDetalle TSD " & _
            "ON(TSD.CodTipoSolicitud=TS.CodTipoSolicitud AND " & _
            "TS.CodTipoSolicitud='" & Codigo_Operacion_Suscripcion & "') " & _
            "WHERE TS.CodCorto<>'T'"
        Set adoRegistro = .Execute
        
        Do Until adoRegistro.EOF
            ReDim Preserve arrLeyendaSuscripcion(intCont)
            arrLeyendaSuscripcion(intCont) = adoRegistro("ValorLeyenda")
            
            adoRegistro.MoveNext
            intCont = intCont + 1
        Loop
        adoRegistro.Close
        
        intCont = 0
        '*** Rescate ***
        .CommandText = "SELECT RTRIM(TSD.CodCorto) + space(1) + '=' + space(1) + RTRIM(TS.DescripTipoSolicitud) + space(1) + RTRIM(TSD.DescripDetalleTipoSolicitud)  ValorLeyenda " & _
            "FROM TipoSolicitud TS JOIN TipoSolicitudDetalle TSD " & _
            "ON(TSD.CodTipoSolicitud=TS.CodTipoSolicitud AND " & _
            "TS.CodTipoSolicitud='" & Codigo_Operacion_Rescate & "') " & _
            "WHERE TS.CodCorto<>'T'"
        Set adoRegistro = .Execute
        
        Do Until adoRegistro.EOF
            ReDim Preserve arrLeyendaRescate(intCont)
            arrLeyendaRescate(intCont) = adoRegistro("ValorLeyenda")
            
            adoRegistro.MoveNext
            intCont = intCont + 1
        Loop
        adoRegistro.Close: Set adoRegistro = Nothing
            
    End With
    
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgConsulta.Columns(1).Width = tdgConsulta.Width * 0.01 * 14
    tdgConsulta.Columns(2).Width = tdgConsulta.Width * 0.01 * 5
    tdgConsulta.Columns(3).Width = tdgConsulta.Width * 0.01 * 9
    tdgConsulta.Columns(4).Width = tdgConsulta.Width * 0.01 * 6
    tdgConsulta.Columns(5).Width = tdgConsulta.Width * 0.01 * 30
    tdgConsulta.Columns(6).Width = tdgConsulta.Width * 0.01 * 13
    
    '*** Verificando Nivel de Acceso de Usuario ***
'    strNivAcceso = AccesoForm(gstrNomOpc, gstrNumInd)

    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmSolicitudParticipe = Nothing
    gstrCodParticipe = Valor_Caracter
    tmrHora.Enabled = False
    Call OcultarReportes
    frmMainMdi.stbMdi.Panels(3).Text = "Acción..."
    
End Sub



Private Sub lblCuotasBloqueadas_Change()

    Call FormatoMillarEtiqueta(lblCuotasBloqueadas, Decimales_CantCuota)
    
End Sub

Private Sub lblCuotasDisponibles_Change()

    Call FormatoMillarEtiqueta(lblCuotasDisponibles, Decimales_CantCuota)
    
End Sub

Private Sub lblDescrip_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim intContador As Integer
    
    If Index = 21 Then
        intContador = UBound(arrLeyendaSuscripcion)
        
        lstLeyenda.AddItem "Leyenda :"
        lstLeyenda.AddItem ""
        
        For intContador = 0 To (UBound(arrLeyendaSuscripcion))
            lstLeyenda.AddItem arrLeyendaSuscripcion(intContador)
        Next
        
        lstLeyenda.Height = lblDescrip(Index).Height * (intContador + 2)
        lstLeyenda.Left = lblDescrip(Index).Left
        lstLeyenda.Top = lblDescrip(Index).Top + lblDescrip(Index).Height
        lstLeyenda.Width = 3300
        lstLeyenda.Visible = True
    End If
    
    If Index = 30 Then
        intContador = UBound(arrLeyendaRescate)
        
        lstLeyenda.AddItem "Leyenda :"
        lstLeyenda.AddItem ""
        
        For intContador = 0 To (UBound(arrLeyendaRescate))
            lstLeyenda.AddItem arrLeyendaRescate(intContador)
        Next
        
        lstLeyenda.Height = lblDescrip(Index).Height * (intContador + 2)
        lstLeyenda.Left = lblDescrip(Index).Left
        lstLeyenda.Top = lblDescrip(Index).Top + lblDescrip(Index).Height
        lstLeyenda.Width = 3800
        lstLeyenda.Visible = True
    End If
    
End Sub

Private Sub lblDescrip_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Index = 21 Or Index = 30 Then
        lstLeyenda.Clear
        lstLeyenda.Visible = False
    End If
    
End Sub

Private Sub lblDescripParticipe_Change()

'    If gstrCodParticipe <> Valor_Caracter Then
'        lblCuotasDisponibles.Caption = CStr(ObtenerCuotasParticipe(gstrCodParticipe, strCodFondo, gstrCodAdministradora, Valor_Caracter, Valor_Caracter))
'
'        If strCodTipoOperacion = Codigo_Operacion_Rescate And gstrTipoAdministradora = Codigo_Tipo_Fondo_Inversion Then
'            txtCantCuotasSolicitud.Text = CStr(lblCuotasDisponibles.Caption)
'        End If
'
'        lblCuotasBloqueadas.Caption = CStr(ObtenerCuotasParticipe(gstrCodParticipe, strCodFondo, gstrCodAdministradora, Valor_Caracter, Valor_Indicador))
'    End If
    
End Sub

Private Sub lblMontoSolicitud_Change()

    Call FormatoMillarEtiqueta(lblMontoSolicitud, Decimales_Monto)
    
End Sub

Private Sub lblValorCuota_Change()

    Call FormatoMillarEtiqueta(lblValorCuota, Decimales_ValorCuota)
    
End Sub

Private Sub tabSolicitud_Click(PreviousTab As Integer)

    Select Case tabSolicitud.Tab
        Case 1
            If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vQuery)
            If strEstado = Reg_Defecto Then tabSolicitud.Tab = 0
        
    End Select
    
End Sub



Private Sub tdgConsulta_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)

    If ColIndex = 6 Then
        Call DarFormatoValor(Value, Decimales_CantCuota)
    End If
            
    If ColIndex = 7 Then
        Call DarFormatoValor(Value, Decimales_Monto)
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

Private Sub tmrHora_Timer()

    dtpHoraSolicitud.Value = ObtenerHoraServidor
    
End Sub

Private Sub txtCantCuotasSolicitud_Change()
    
    'Call FormatoCajaTexto(txtCantCuotasSolicitud, Decimales_CantCuota)
        
End Sub

Private Sub txtCantCuotasSolicitud_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        blnMonto = False: blnCuota = True
        Call Calcular
    End If
    
    'Call ValidaCajaTexto(KeyAscii, "M", txtCantCuotasSolicitud, Decimales_CantCuota)
            
End Sub

Private Sub txtCantCuotasSolicitud_LostFocus()

    blnMonto = False: blnCuota = True
    Call Calcular

End Sub

Private Sub txtMontoComision_Change()

    'Call FormatoCajaTexto(txtMontoComision, Decimales_Monto)
    
End Sub

Private Sub txtMontoComision_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        blnMonto = False: blnCuota = True
        Call Calcular
    End If

End Sub

Private Sub txtMontoComision_LostFocus()

    blnMonto = False: blnCuota = True
    Call Calcular

End Sub

Private Sub txtMontoIgv_Change()

    Call FormatoCajaTexto(txtMontoIgv, Decimales_Monto)
    
End Sub

Private Sub txtMontoNetoSolicitud_Change()

    Call FormatoCajaTexto(txtMontoNetoSolicitud, Decimales_Monto)
    
End Sub

Private Sub txtMontoNetoSolicitud_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        blnMonto = True: blnCuota = False
        Call Calcular
    End If
    
    Call ValidaCajaTexto(KeyAscii, "M", txtMontoNetoSolicitud, Decimales_Monto)
    
End Sub

Private Sub txtMontoSolicitud_Change()

    lblMontoNetoSolicitud.Caption = CStr(txtMontoSolicitud.Value)

End Sub

Private Sub txtNumDocumento_Change()

    Dim adoRegistro     As ADODB.Recordset
        
    If chkPagoParcial.Value Then
        Set adoRegistro = New ADODB.Recordset
        
        adoComm.CommandText = "SELECT FechaPago FROM ParticipePagoSuscripcion " & _
            "WHERE CodParticipe='" & gstrCodParticipe & "' AND CodFondo='" & strCodFondo & "' AND " & _
            "CodAdministradora='" & gstrCodAdministradora & "'"
        Set adoRegistro = adoComm.Execute
        
'        If adoRegistro.EOF Then
'            If curMontoPago > 0 Or dblPorcenPago > 0 Then
'                If curMontoPago > 0 Then txtMontoNetoSolicitud.Text = curMontoPago
'                If dblPorcenPago > 0 Then txtMontoNetoSolicitud.Text = curMontoEmitido * dblPorcenPago * 0.01
'            End If
'        End If
        
        Do While Not adoRegistro.EOF
            If IsNull(adoRegistro("FechaPago")) Or Convertyyyymmdd(adoRegistro("FechaPago")) = Convertyyyymmdd(CVDate(Valor_Fecha)) Then
                MsgBox "El Partícipe tiene cuotas pendientes de pago", vbCritical, Me.Caption
                txtMontoNetoSolicitud.Text = "0"
                Call Deshabilita
                Exit Do
            End If
        
            adoRegistro.MoveNext
        Loop
        adoRegistro.Close
    End If
    Set adoRegistro = Nothing
    
End Sub

Private Sub txtNumDocumento_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Call ObtenerDatosParticipe
    End If
    
End Sub

Private Sub txtNumPapeleta_LostFocus()

    txtNumPapeleta.Text = Format(txtNumPapeleta.Text, "000000000000000")
    
    If cboFondo.ListIndex > 0 And cboEjecutivo.ListIndex > -1 And cboTipoOperacion.ListIndex > -1 Then
        Call Habilita
        If strHoraCorte = "00:00" Then Call Deshabilita
        If chkPagoParcial.Value Then
            If dblPorcenPago = 0 And curMontoPago = 0 Then Call Deshabilita
        End If
    Else
        Call Deshabilita
    End If
    
    If strEstado = Reg_Adicion Then
        If Not ValidarNumFolio(Trim(txtNumPapeleta.Text), strCodTipoOperacion, strCodFondo, Me) Then Exit Sub
    End If
    
End Sub

Private Sub txtValorCuota_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        blnMonto = False: blnCuota = True
        Call Calcular
    End If

End Sub

Private Sub txtValorCuota_LostFocus()

    blnMonto = False: blnCuota = True
    Call Calcular

End Sub
