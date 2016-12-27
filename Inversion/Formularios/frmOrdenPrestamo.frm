VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{830D5A06-9B70-4F8C-98B6-7A19C4E7760B}#1.0#0"; "TAMControls.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmOrdenPrestamo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Préstamos - Operaciones de Financiamiento"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13395
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   13395
   Begin TAMControls2.ucBotonEdicion2 cmdOpcion 
      Height          =   735
      Left            =   810
      TabIndex        =   95
      Top             =   7920
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
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
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
      Left            =   9870
      Picture         =   "frmOrdenPrestamo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   91
      Top             =   6930
      Width           =   1200
   End
   Begin TabDlg.SSTab tabRFCortoPlazo 
      Height          =   7785
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   13732
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabsPerRow      =   5
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
      TabPicture(0)   =   "frmOrdenPrestamo.frx":0562
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "tdgConsulta"
      Tab(0).Control(1)=   "fraCriterio"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Condiciones del Préstamo"
      TabPicture(1)   =   "frmOrdenPrestamo.frx":057E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraDatosTasa"
      Tab(1).Control(1)=   "fraDatosBasicos"
      Tab(1).Control(2)=   "fraDatosTitulo"
      Tab(1).Control(3)=   "gbParametrosCuponera"
      Tab(1).Control(4)=   "fraTramos"
      Tab(1).Control(5)=   "cmdSiguiente"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Detalle del Calendario"
      TabPicture(2)   =   "frmOrdenPrestamo.frx":059A
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "cmdGuardar"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cmdAnterior"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "fraSimulacion"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "fraComisionMontoFL1"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      Begin VB.Frame fraComisionMontoFL1 
         Caption         =   "Comisiones y Montos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5865
         Left            =   150
         TabIndex        =   14
         Top             =   450
         Width           =   2715
         Begin VB.TextBox txtObservacion 
            BackColor       =   &H00FFFFFF&
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
            Height          =   645
            Left            =   210
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   102
            Top             =   4890
            Width           =   2280
         End
         Begin TAMControls.TAMTextBox txtInteresTotal 
            Height          =   315
            Left            =   1410
            TabIndex        =   104
            Top             =   1080
            Width           =   1150
            _ExtentX        =   2037
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
            Container       =   "frmOrdenPrestamo.frx":05B6
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
         Begin TAMControls.TAMTextBox txtIGVInteresTotal 
            Height          =   315
            Left            =   1410
            TabIndex        =   105
            Top             =   1440
            Width           =   1150
            _ExtentX        =   2037
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
            Container       =   "frmOrdenPrestamo.frx":05D2
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
         Begin TAMControls.TAMTextBox txtCantCuotas 
            Height          =   315
            Left            =   1410
            TabIndex        =   106
            Top             =   1800
            Width           =   1150
            _ExtentX        =   2037
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
            Container       =   "frmOrdenPrestamo.frx":05EE
            Estilo          =   3
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            EnterTab        =   -1  'True
            Apariencia      =   1
            Borde           =   1
         End
         Begin TAMControls.TAMTextBox txtTotalPago 
            Height          =   315
            Left            =   1410
            TabIndex        =   107
            Top             =   2160
            Width           =   1150
            _ExtentX        =   2037
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
            Container       =   "frmOrdenPrestamo.frx":060A
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
         Begin TAMControls.TAMTextBox txtPorcenComision 
            Height          =   315
            Left            =   1410
            TabIndex        =   110
            Top             =   2520
            Width           =   1155
            _ExtentX        =   2037
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
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            Container       =   "frmOrdenPrestamo.frx":0626
            Text            =   "0.000000"
            Decimales       =   6
            Estilo          =   4
            ColorEnfoque    =   8454143
            EnterTab        =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   6
         End
         Begin TAMControls.TAMTextBox txtMontoComision 
            Height          =   315
            Left            =   1410
            TabIndex        =   111
            Top             =   2880
            Width           =   1155
            _ExtentX        =   2037
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
            Container       =   "frmOrdenPrestamo.frx":0642
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
         Begin TAMControls.TAMTextBox txtTotalRecibido 
            Height          =   315
            Left            =   1410
            TabIndex        =   113
            Top             =   3690
            Width           =   1155
            _ExtentX        =   2037
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
            Container       =   "frmOrdenPrestamo.frx":065E
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
         Begin TAMControls.TAMTextBox txtIGVComision 
            Height          =   315
            Left            =   1410
            TabIndex        =   115
            Top             =   3240
            Width           =   1155
            _ExtentX        =   2037
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
            Container       =   "frmOrdenPrestamo.frx":067A
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
         Begin TAMControls.TAMTextBox txtValorNominalDscto 
            Height          =   315
            Left            =   1410
            TabIndex        =   116
            Top             =   720
            Width           =   1155
            _ExtentX        =   2037
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
            Container       =   "frmOrdenPrestamo.frx":0696
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
         Begin VB.Label lblValorNominalDscto 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor Nom."
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
            Left            =   210
            TabIndex        =   117
            Top             =   795
            Width           =   945
         End
         Begin VB.Label lblIGVComision 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IGV Comisión"
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
            Left            =   210
            TabIndex        =   114
            Top             =   3320
            Width           =   1170
         End
         Begin VB.Line lnDivisoria1 
            X1              =   2590
            X2              =   120
            Y1              =   3600
            Y2              =   3600
         End
         Begin VB.Label lblTotalRecibido 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tot. Recibido"
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
            Left            =   210
            TabIndex        =   112
            Top             =   3770
            Width           =   1170
         End
         Begin VB.Label lblMontoComision 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Comisión"
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
            Left            =   210
            TabIndex        =   109
            Top             =   2960
            Width           =   750
         End
         Begin VB.Label lblPorcenComision 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "% Comisión"
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
            Left            =   210
            TabIndex        =   108
            Top             =   2600
            Width           =   945
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Instrucciones"
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
            Left            =   210
            TabIndex        =   103
            Top             =   4620
            Width           =   1155
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº de Cuotas"
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
            Index           =   34
            Left            =   210
            TabIndex        =   99
            Top             =   1880
            Width           =   1125
         End
         Begin VB.Label lblDescripMoneda 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Height          =   195
            Index           =   0
            Left            =   1830
            TabIndex        =   18
            Top             =   480
            Width           =   405
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
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
            Index           =   36
            Left            =   210
            TabIndex        =   17
            Top             =   2240
            Width           =   450
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Interés Total"
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
            Index           =   92
            Left            =   210
            TabIndex        =   16
            Top             =   1160
            Width           =   1080
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IGV Total"
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
            Index           =   111
            Left            =   210
            TabIndex        =   15
            Top             =   1520
            Width           =   825
         End
      End
      Begin VB.Frame fraSimulacion 
         Caption         =   "Simulación de Cuotas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5865
         Left            =   2940
         TabIndex        =   100
         Top             =   450
         Width           =   10215
         Begin TrueOleDBGrid60.TDBGrid tdgCalendario 
            Bindings        =   "frmOrdenPrestamo.frx":06B2
            Height          =   5355
            Left            =   150
            OleObjectBlob   =   "frmOrdenPrestamo.frx":06CE
            TabIndex        =   101
            Top             =   300
            Width           =   9885
         End
      End
      Begin VB.CommandButton cmdAnterior 
         Caption         =   "Anterior"
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
         Left            =   555
         Picture         =   "frmOrdenPrestamo.frx":6B00
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   6870
         Width           =   1200
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "Guardar"
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
         Left            =   11490
         Picture         =   "frmOrdenPrestamo.frx":6F85
         Style           =   1  'Graphical
         TabIndex        =   97
         Top             =   6870
         Width           =   1200
      End
      Begin VB.CommandButton cmdSiguiente 
         Caption         =   "Siguiente "
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
         Left            =   -63510
         Picture         =   "frmOrdenPrestamo.frx":7579
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   6870
         Width           =   1200
      End
      Begin VB.Frame fraTramos 
         Caption         =   "Especificación de tramos"
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
         Height          =   2835
         Left            =   -65700
         TabIndex        =   77
         Top             =   3930
         Width           =   3855
         Begin TrueOleDBGrid60.TDBGrid tdgTramos 
            Height          =   1485
            Left            =   210
            OleObjectBlob   =   "frmOrdenPrestamo.frx":79FF
            TabIndex        =   83
            Top             =   1155
            Width           =   3435
         End
         Begin VB.OptionButton optAmortizacion 
            Caption         =   "Por amortización"
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
            Left            =   210
            TabIndex        =   79
            Top             =   800
            Width           =   1785
         End
         Begin VB.OptionButton optCuota 
            Caption         =   "Por cuota"
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
            Left            =   2190
            TabIndex        =   78
            Top             =   800
            Value           =   -1  'True
            Width           =   1215
         End
         Begin TAMControls.TAMTextBox txtCantidadTramos 
            Height          =   315
            Left            =   2070
            TabIndex        =   80
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
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
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            Container       =   "frmOrdenPrestamo.frx":AE35
            Estilo          =   3
            Apariencia      =   1
            Borde           =   1
         End
         Begin VB.Label lblCantidadTramos 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad de tramos"
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
            Left            =   210
            TabIndex        =   81
            Top             =   440
            Width           =   1845
         End
      End
      Begin VB.Frame gbParametrosCuponera 
         Caption         =   "Parámetros de calendario de cuotas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3675
         Left            =   -70620
         TabIndex        =   65
         Top             =   3930
         Width           =   4845
         Begin VB.ComboBox cboAmortizacion 
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
            Left            =   2250
            Style           =   2  'Dropdown List
            TabIndex        =   89
            Top             =   360
            Width           =   2430
         End
         Begin TAMControls.TAMTextBox txtPeriodoGracia 
            Height          =   315
            Left            =   2250
            TabIndex        =   86
            Top             =   1800
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   556
            BackColor       =   16777215
            ForeColor       =   0
            Enabled         =   0   'False
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
            Container       =   "frmOrdenPrestamo.frx":AE51
            Estilo          =   3
            Apariencia      =   1
            Borde           =   1
         End
         Begin VB.ComboBox cboUnidadesPeriodo 
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
            Left            =   2880
            Style           =   2  'Dropdown List
            TabIndex        =   72
            Top             =   720
            Width           =   1800
         End
         Begin VB.CheckBox chkAPartir 
            Caption         =   "Cálculo a partir del"
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
            Left            =   210
            TabIndex        =   71
            Top             =   1125
            Width           =   1935
         End
         Begin VB.CheckBox chkCorteAFinPeriodo 
            Caption         =   "Corte a fin de periodo de cuota"
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
            Left            =   210
            TabIndex        =   69
            Top             =   1485
            Width           =   3165
         End
         Begin VB.ComboBox cboDesplazamientoCorte 
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
            Left            =   1600
            Style           =   2  'Dropdown List
            TabIndex        =   68
            Top             =   2880
            Width           =   3100
         End
         Begin VB.ComboBox cboDesplazamientoPago 
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
            Left            =   1600
            Style           =   2  'Dropdown List
            TabIndex        =   67
            Top             =   3240
            Width           =   3100
         End
         Begin VB.ComboBox cbTipoDia 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1410
            Style           =   2  'Dropdown List
            TabIndex        =   66
            Top             =   4590
            Width           =   2505
         End
         Begin MSComCtl2.DTPicker dtpAPartir 
            Height          =   315
            Left            =   2250
            TabIndex        =   70
            Top             =   1080
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   175505409
            CurrentDate     =   40413
         End
         Begin TAMControls.TAMTextBox txtUnidadesPeriodo 
            Height          =   315
            Left            =   2250
            TabIndex        =   73
            Top             =   720
            Width           =   600
            _ExtentX        =   1058
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
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            Container       =   "frmOrdenPrestamo.frx":AE6D
            Estilo          =   3
            Apariencia      =   1
            Borde           =   1
         End
         Begin TAMControls.TAMTextBox txtDiasMinimosPagoInteres 
            Height          =   315
            Left            =   3210
            TabIndex        =   74
            Top             =   2160
            Width           =   600
            _ExtentX        =   1058
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
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            Container       =   "frmOrdenPrestamo.frx":AE89
            Estilo          =   3
            Apariencia      =   1
            Borde           =   1
         End
         Begin VB.Label lblAmortizacion 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Modo de Amortización:"
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
            Left            =   210
            TabIndex        =   90
            Top             =   435
            Width           =   1935
         End
         Begin VB.Label lblCuotasPeriodoGracia 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cuotas"
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
            Left            =   2910
            TabIndex        =   87
            Top             =   1875
            Width           =   600
         End
         Begin VB.Label lblDesplazamientoCort 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de corte:"
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
            Left            =   210
            TabIndex        =   88
            Top             =   2955
            Width           =   1365
         End
         Begin VB.Label lblPeriodoGracia 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Periodo de gracia:"
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
            Left            =   210
            TabIndex        =   85
            Top             =   1875
            Width           =   1905
         End
         Begin VB.Label lblPeriodoCupon 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Periodo de cuota cada:"
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
            Left            =   210
            TabIndex        =   52
            Top             =   780
            Width           =   2010
         End
         Begin VB.Label lblDesplazamientoCorte 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Desplazamiento:"
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
            Left            =   210
            TabIndex        =   53
            Top             =   2595
            Width           =   1410
         End
         Begin VB.Label lblDesplazamientoPago 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de pago:"
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
            Left            =   210
            TabIndex        =   54
            Top             =   3315
            Width           =   1350
         End
         Begin VB.Label lblTipoDia 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de día:"
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
            Left            =   90
            TabIndex        =   76
            Top             =   4650
            Width           =   1425
         End
         Begin VB.Label lblDiasMinimosPagoInteres 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Días mínimos de Pago de Interés:"
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
            Left            =   210
            TabIndex        =   75
            Top             =   2235
            Width           =   2895
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
         Height          =   1545
         Left            =   -74850
         TabIndex        =   39
         Top             =   450
         Width           =   13035
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
            Left            =   11010
            Picture         =   "frmOrdenPrestamo.frx":AEA5
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "Enviar a BackOffice"
            Top             =   360
            Width           =   1200
         End
         Begin VB.ComboBox cboEstado 
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
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   1080
            Width           =   4605
         End
         Begin VB.ComboBox cboTipoInstrumento 
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
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   720
            Width           =   4605
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
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   360
            Width           =   4605
         End
         Begin MSComCtl2.DTPicker dtpFechaOrdenDesde 
            Height          =   285
            Left            =   9240
            TabIndex        =   44
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
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
            Height          =   285
            Left            =   9240
            TabIndex        =   45
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
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
            BackStyle       =   0  'Transparent
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
            Left            =   360
            TabIndex        =   51
            Top             =   1155
            Width           =   600
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   360
            TabIndex        =   50
            Top             =   795
            Width           =   1005
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   8520
            TabIndex        =   49
            Top             =   735
            Width           =   510
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   8520
            TabIndex        =   48
            Top             =   435
            Width           =   555
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   360
            TabIndex        =   47
            Top             =   435
            Width           =   540
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   7140
            TabIndex        =   46
            Top             =   435
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
         Height          =   1935
         Left            =   -74850
         TabIndex        =   20
         Top             =   1980
         Width           =   13005
         Begin VB.TextBox txtDescripOrden 
            BackColor       =   &H00FFFFFF&
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
            Left            =   6600
            MaxLength       =   45
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   720
            Width           =   5250
         End
         Begin VB.ComboBox cboMoneda 
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
            Left            =   9570
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   1440
            Width           =   2280
         End
         Begin VB.TextBox txtNemonico 
            BackColor       =   &H00FFFFFF&
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
            Left            =   2430
            MaxLength       =   15
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   720
            Width           =   2000
         End
         Begin VB.TextBox txtNumDocDscto 
            BackColor       =   &H00FFFFFF&
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
            Left            =   2430
            MaxLength       =   10
            TabIndex        =   21
            Top             =   360
            Width           =   2000
         End
         Begin MSComCtl2.DTPicker dtpFechaOrden 
            Height          =   315
            Left            =   2430
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   1080
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   175505409
            CurrentDate     =   38776
         End
         Begin MSComCtl2.DTPicker dtpFechaEmisionDocumento 
            Height          =   315
            Left            =   6600
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   360
            Visible         =   0   'False
            Width           =   1500
            _ExtentX        =   2646
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
            Format          =   175505409
            CurrentDate     =   38776
         End
         Begin MSComCtl2.DTPicker dtpFechaVencimiento 
            Height          =   315
            Left            =   2430
            TabIndex        =   27
            Top             =   1440
            Width           =   1500
            _ExtentX        =   2646
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
            Format          =   175505409
            CurrentDate     =   38776
         End
         Begin TAMControls.TAMTextBox txtDiasPlazo 
            Height          =   315
            Left            =   6600
            TabIndex        =   28
            Top             =   1080
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
            Container       =   "frmOrdenPrestamo.frx":B400
            Estilo          =   3
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            EnterTab        =   -1  'True
            Apariencia      =   1
            Borde           =   1
         End
         Begin TAMControls.TAMTextBox txtValorNominalDocumento 
            Height          =   315
            Left            =   6600
            TabIndex        =   29
            Top             =   1440
            Width           =   1455
            _ExtentX        =   2566
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
            Container       =   "frmOrdenPrestamo.frx":B41C
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
            BackStyle       =   0  'Transparent
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
            Left            =   210
            TabIndex        =   38
            Top             =   1160
            Width           =   525
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   4750
            TabIndex        =   37
            Top             =   800
            Width           =   1005
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   4750
            TabIndex        =   36
            Top             =   1160
            Width           =   1095
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   8670
            TabIndex        =   35
            Top             =   1515
            Width           =   690
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Emisión Documento"
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
            Left            =   4750
            TabIndex        =   34
            Top             =   440
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nemónico"
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
            Index           =   85
            Left            =   210
            TabIndex        =   33
            Top             =   800
            Width           =   840
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nro. de Doc. Referencia"
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
            TabIndex        =   32
            Top             =   440
            Width           =   2100
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vencimiento Operación"
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
            Left            =   210
            TabIndex        =   31
            Top             =   1520
            Width           =   1965
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Height          =   195
            Index           =   183
            Left            =   4750
            TabIndex        =   30
            Top             =   1520
            Width           =   1185
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
         Height          =   1515
         Left            =   -74850
         TabIndex        =   1
         Top             =   450
         Width           =   13005
         Begin VB.ComboBox cboClaseInstrumento 
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
            Left            =   2060
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1080
            Width           =   4185
         End
         Begin VB.ComboBox cboTipoInstrumentoOrden 
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
            Left            =   2060
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   720
            Width           =   4185
         End
         Begin VB.ComboBox cboFondoOrden 
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
            Left            =   2060
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   360
            Width           =   4185
         End
         Begin VB.ComboBox cboLineaCliente 
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
            Left            =   2060
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1920
            Visible         =   0   'False
            Width           =   4185
         End
         Begin VB.ComboBox cboAcreedor 
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
            Left            =   7700
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   360
            Width           =   4185
         End
         Begin VB.ComboBox cboGestor 
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
            Left            =   7700
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   720
            Width           =   4185
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Acreedor:"
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
            Left            =   6600
            TabIndex        =   13
            Top             =   440
            Width           =   840
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            TabIndex        =   12
            Top             =   800
            Width           =   1005
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            TabIndex        =   11
            Top             =   440
            Width           =   540
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            TabIndex        =   10
            Top             =   1160
            Width           =   480
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   330
            TabIndex        =   9
            Top             =   1980
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   6600
            TabIndex        =   8
            Top             =   800
            Width           =   570
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmOrdenPrestamo.frx":B438
         Height          =   5355
         Left            =   -74850
         OleObjectBlob   =   "frmOrdenPrestamo.frx":B452
         TabIndex        =   19
         Top             =   2100
         Width           =   13035
      End
      Begin VB.Frame fraDatosTasa 
         Caption         =   "Datos de la Tasa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3675
         Left            =   -74850
         TabIndex        =   55
         Top             =   3930
         Width           =   4155
         Begin VB.CheckBox chkigv 
            Alignment       =   1  'Right Justify
            Caption         =   "Cálculo con IGV"
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
            Left            =   180
            TabIndex        =   84
            Top             =   2550
            Value           =   1  'Checked
            Width           =   1965
         End
         Begin VB.ComboBox cboBaseAnual 
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
            Left            =   2000
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   2160
            Width           =   2000
         End
         Begin VB.ComboBox cboTipoTasa 
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
            Left            =   2000
            Style           =   2  'Dropdown List
            TabIndex        =   58
            Top             =   1080
            Width           =   2000
         End
         Begin VB.ComboBox cboPeriodoTasa 
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
            ItemData        =   "frmOrdenPrestamo.frx":13CF2
            Left            =   2000
            List            =   "frmOrdenPrestamo.frx":13CF4
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   1440
            Width           =   2000
         End
         Begin VB.ComboBox cboPeriodoCapitalizacion 
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
            ItemData        =   "frmOrdenPrestamo.frx":13CF6
            Left            =   2000
            List            =   "frmOrdenPrestamo.frx":13CF8
            Style           =   2  'Dropdown List
            TabIndex        =   56
            Top             =   1800
            Width           =   2000
         End
         Begin TAMControls.TAMTextBox txtTasa 
            Height          =   315
            Left            =   1995
            TabIndex        =   93
            Top             =   360
            Width           =   1995
            _ExtentX        =   3519
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
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            Container       =   "frmOrdenPrestamo.frx":13CFA
            Text            =   "0.000000"
            Decimales       =   6
            Estilo          =   4
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   6
         End
         Begin TAMControls.TAMTextBox txtTasaMora 
            Height          =   315
            Left            =   2000
            TabIndex        =   94
            Top             =   720
            Width           =   1995
            _ExtentX        =   3519
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
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            Container       =   "frmOrdenPrestamo.frx":13D16
            Text            =   "0.000000"
            Decimales       =   6
            Estilo          =   4
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   6
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tasa Moratoria"
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
            Left            =   210
            TabIndex        =   82
            Top             =   800
            Width           =   1305
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   210
            TabIndex        =   64
            Top             =   440
            Width           =   1005
         End
         Begin VB.Label lblPeriodoTasa 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Periodo:"
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
            Left            =   210
            TabIndex        =   63
            Top             =   1520
            Width           =   885
         End
         Begin VB.Label lblTipoTasa 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de tasa:"
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
            Left            =   210
            TabIndex        =   62
            Top             =   1160
            Width           =   1335
         End
         Begin VB.Label lblBaseCalculo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Base de cálculo:"
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
            Left            =   210
            TabIndex        =   61
            Top             =   2220
            Width           =   1485
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Capitalización:"
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
            Left            =   210
            TabIndex        =   60
            Top             =   1880
            Width           =   1305
         End
      End
   End
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   11430
      TabIndex        =   96
      Top             =   7920
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
      Caption0        =   "&Salir"
      Tag0            =   "9"
      ToolTipText0    =   "Salir"
      UserControlWidth=   1200
   End
   Begin MSAdodcLib.Adodc adoConsulta 
      Height          =   330
      Left            =   6630
      Top             =   8160
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
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "frmOrdenPrestamo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrFondo()                  As String
Dim arrFondoOrden()             As String
Dim arrTipoInstrumento()        As String
Dim arrTipoInstrumentoOrden()   As String
Dim arrEstado()                 As String
Dim arrAcreedor()               As String
Dim arrMoneda()                 As String
Dim arrGestor()                 As String
Dim arrBaseAnual()              As String
Dim arrTipoTasa()               As String
Dim arrClaseInstrumento()       As String
Dim arrSubClaseInstrumento()    As String
Dim arrPeriodoTasa()            As String
Dim arrPeriodoCuota()           As String
Dim arrPeriodoCapitalizacion()  As String
Dim arrBaseCalculo()            As String
Dim arrLineaClienteLista()      As String
Dim arrDesplazamientoPago()     As String
Dim arrDesplazamientoCorte()    As String
Dim arrTipoAmortizacion()       As String

Dim strTipoTasa                 As String

Dim strCodFondo                 As String
Dim strCodFondoOrden            As String
Dim strCodTipoInstrumento       As String
Dim strCodTipoInstrumentoOrden  As String
Dim strCodEstado                As String
Dim strCodTipoOrden             As String
Dim strCodOperacion             As String
Dim strCodAcreedor              As String
Dim strCodMoneda                As String
Dim strPeriodoTasa              As String
Dim strPeriodoCapitalizacion    As String
Dim strCodGestor                As String
Dim strCodBaseAnual             As String
Dim strCodTipoTasa              As String
Dim strCodClaseInstrumento      As String
Dim strEstado                   As String
Dim strSQL                      As String
Dim strCodFile                  As String
Dim strIndIGV                   As String
Dim strPeriodoCuota             As String
Dim strTipoAmortizacion         As String
Dim strIndFechaAPartir          As String
Dim strIndFinPeriodo            As String
Dim strCodDesplazamientoCorte   As String
Dim strCodDesplazamientoPago    As String
Dim datFechaInicioCalendario    As Date
Dim strNemonico                 As String
Dim strNumOrden                 As String
Dim strTipoOrden                As String
Dim intCantTramos               As Integer
Dim strTipoTramo                As String
Dim dblPorcenImptoInteres       As Double

Dim strEstadoOrden              As String
Dim dblTipoCambio               As Double
Dim dblTasaInteres              As Double
Dim dblTasaInteresMoratoria     As Double
Dim intCantCuotas               As Integer

Dim blnCargadoDesdeCartera      As Boolean
Dim blnCargarCabeceraAnexo      As Boolean
Dim blnCancelaPrepago           As Boolean

Dim blnFormReady                As Boolean
Dim blnPreInfoReady             As Boolean
Dim blnLockPorcenComision       As Boolean
Dim blnLock                     As Boolean

Dim intDiasInteresMinimo        As Integer
Dim strTramoXML                 As String

Dim adoAuxiliar                 As ADODB.Recordset
Dim adoTramos                   As ADODB.Recordset
Dim adoCalendario               As ADODB.Recordset

Private Sub InicializarValores()
    txtNumDocDscto.Text = ""
    txtNemonico.Text = ""
    
    If cboAcreedor.ListCount > 0 Then
        cboAcreedor.ListIndex = 0
    End If
    If cboGestor.ListCount > 0 Then
        cboGestor.ListIndex = 0
    End If
    
    dtpFechaOrden.Value = Convertddmmyyyy(gstrFechaActual)
    dtpFechaVencimiento.Value = Convertddmmyyyy(gstrFechaActual)
    dtpFechaEmisionDocumento.Value = Convertddmmyyyy(gstrFechaActual)
    dtpAPartir.Value = Convertddmmyyyy(gstrFechaActual)
    txtDescripOrden.Text = ""
    txtDiasPlazo.Text = 0
    txtValorNominalDocumento.Text = 0
    txtTasa.Text = 0
    txtTasaMora.Text = 0
    txtCantidadTramos.Text = 1
    txtUnidadesPeriodo.Text = 1
    txtPeriodoGracia.Text = 0
    txtDiasMinimosPagoInteres.Text = 0
    
    txtMontoComision.Text = 0
    txtIGVComision.Text = 0
    txtPorcenComision.Text = 0
    
    chkAPartir.Value = False
    chkCorteAFinPeriodo.Value = False
    chkigv.Value = vbChecked
    strIndIGV = Valor_Indicador
    datFechaInicioCalendario = dtpFechaOrden.Value
    
    dtpFechaOrdenDesde.Value = gdatFechaActual
    dtpFechaOrdenHasta.Value = dtpFechaOrdenDesde.Value
    
    LimpiarRecordSetTramos
       
End Sub

Private Sub LimpiarRecordSetTramos()
    Set adoTramos = New ADODB.Recordset
    
End Sub

Private Sub cboAcreedor_Click()
    strCodAcreedor = Trim$(arrAcreedor(cboAcreedor.ListIndex))
    adoComm.CommandText = "SELECT DescripNemonico FROM InstitucionPersona WHERE CodPersona = '" & strCodAcreedor & "' AND TipoPersona = '" & Codigo_Tipo_Persona_Emisor & "'"
    Set adoAuxiliar = adoComm.Execute
    
    If Not adoAuxiliar.EOF Then
        strNemonico = adoAuxiliar("DescripNemonico").Value
    Else
        strNemonico = ""
    End If
    Call GenerarNemonico
End Sub

Private Sub GenerarNemonico()
    txtNemonico.Text = strNemonico & "-" & txtNumDocDscto.Text
    txtDescripOrden.Text = "Préstamo - " & txtNemonico.Text
End Sub

Private Sub cboAmortizacion_Click()
    strTipoAmortizacion = Trim$(arrTipoAmortizacion(cboAmortizacion.ListIndex))
    
    'Si es cuotas por tramos se activa el frame correspondiente
    fraTramos.Enabled = strTipoAmortizacion = "04"
    If (strTipoTramo <> "01") And (strTipoTramo <> "02") Then
        If optAmortizacion.Value Then
            strTipoTramo = "01"
        Else
            strTipoTramo = "02"
        End If
    End If
    
    Call CalcularCantidadCuotas
    
    Call GenerarTramos
End Sub

Private Sub cboBaseAnual_Click()
    strCodBaseAnual = Trim$(arrBaseAnual(cboBaseAnual.ListIndex))
End Sub

Private Sub cboClaseInstrumento_Click()
    strCodClaseInstrumento = Trim$(arrClaseInstrumento(cboClaseInstrumento.ListIndex))
End Sub

Private Sub cboDesplazamientoCorte_Click()
    strCodDesplazamientoCorte = Trim$(arrDesplazamientoCorte(cboDesplazamientoCorte.ListIndex))
End Sub

Private Sub cboDesplazamientoPago_Click()
    strCodDesplazamientoPago = Trim$(arrDesplazamientoPago(cboDesplazamientoPago.ListIndex))
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
    strCodFondoOrden = Trim$(arrFondoOrden(cboFondoOrden.ListIndex))
End Sub

Private Sub cboGestor_Click()
    strCodGestor = Trim$(arrGestor(cboGestor.ListIndex))
End Sub

Private Sub cboMoneda_Click()
    strCodMoneda = Trim$(arrMoneda(cboMoneda.ListIndex))
End Sub

Private Sub cboPeriodoCapitalizacion_Click()
    strPeriodoCapitalizacion = Trim$(arrPeriodoCapitalizacion(cboPeriodoCapitalizacion.ListIndex))
End Sub

Private Sub cboPeriodoTasa_Click()
    strPeriodoTasa = Trim$(arrPeriodoTasa(cboPeriodoTasa.ListIndex))
End Sub

Private Sub cboTipoInstrumento_Change()
    strCodTipoInstrumento = Trim$(arrTipoInstrumento(cboTipoInstrumento.ListIndex))
End Sub

Private Sub cboTipoInstrumentoOrden_Click()
    strCodTipoInstrumentoOrden = Valor_Caracter

    If cboTipoInstrumentoOrden.ListIndex < 0 Then Exit Sub
    
    strCodTipoInstrumentoOrden = Trim$(arrTipoInstrumentoOrden(cboTipoInstrumentoOrden.ListIndex))

    'Asignar nemónico
    Call GenerarNemonico
    
    strSQL = "SELECT IFTON.CodTipoOperacion CODIGO,DescripParametro DESCRIP " & "FROM InversionFileTipoOperacionNegociacion IFTON JOIN AuxiliarParametro AUX ON(AUX.CodParametro=IFTON.CodTipoOperacion AND AUX.CodTipoParametro = 'OPECAJ') " & "WHERE IFTON.CodFile='" & strCodTipoInstrumentoOrden & "' ORDER BY DescripParametro"

    strCodFile = strCodTipoInstrumentoOrden

    '*** Clase de Instrumento ***
    strSQL = "SELECT CodDetalleFile CODIGO,DescripDetalleFile DESCRIP FROM InversionDetalleFile WHERE CodFile='" & strCodTipoInstrumentoOrden & "' AND IndVigente='X' ORDER BY DescripDetalleFile"
    CargarControlLista strSQL, cboClaseInstrumento, arrClaseInstrumento(), Sel_Defecto
    
    If cboClaseInstrumento.ListCount > 0 Then
        cboClaseInstrumento.ListIndex = 0
        cboClaseInstrumento.Enabled = True
    End If
    
End Sub

Private Sub cboTipoTasa_Click()
    strTipoTasa = Trim$(arrTipoTasa(cboTipoTasa.ListIndex))
    
    If strTipoTasa = "02" Then 'Tasa Nominal
        cboPeriodoCapitalizacion.Enabled = True
    Else
        If cboPeriodoCapitalizacion.ListCount > 0 Then cboPeriodoCapitalizacion.ListIndex = 0
        cboPeriodoCapitalizacion.Enabled = False
    End If
    
End Sub

Private Sub cboUnidadesPeriodo_Click()
    strPeriodoCuota = Trim$(arrPeriodoCuota(cboUnidadesPeriodo.ListIndex))
    
    Call CalcularCantidadCuotas
    If strTipoAmortizacion = "04" Then
        Call GenerarTramos
    End If
End Sub

Private Sub chkApartir_Click()
    If chkAPartir.Value = vbChecked Then
        strIndFechaAPartir = Valor_Indicador
        datFechaInicioCalendario = dtpAPartir.Value
        txtDiasPlazo.Text = DateDiff("d", datFechaInicioCalendario, dtpFechaVencimiento.Value)
    Else
        strIndFechaAPartir = Valor_Caracter
        datFechaInicioCalendario = dtpFechaOrden.Value
        txtDiasPlazo.Text = DateDiff("d", datFechaInicioCalendario, dtpFechaVencimiento.Value)
    End If

    dtpAPartir.Enabled = chkAPartir.Value = vbChecked
End Sub

Private Sub chkCorteAFinPeriodo_Click()
    If chkCorteAFinPeriodo.Value = vbChecked Then
        strIndFinPeriodo = Valor_Indicador
    Else
        strIndFinPeriodo = Valor_Caracter
    End If
    
    Call CalcularCantidadCuotas
    If strTipoAmortizacion = "04" Then
        Call GenerarTramos
    End If
End Sub

Private Sub chkigv_Click()
    If chkigv.Value = vbChecked Then
        strIndIGV = Valor_Indicador
    Else
        strIndIGV = Valor_Caracter
    End If
End Sub

Private Sub cmdAnterior_Click()

    With tabRFCortoPlazo
        .TabEnabled(0) = False
        .TabEnabled(1) = True
        .TabEnabled(2) = False
        .Tab = 1
        cmdCancelar.Visible = True
    End With
    
End Sub

Public Sub Adicionar()
    Dim adoAuxiliar As ADODB.Recordset
    
    Dim intCantidadOperaciones     As Integer
    Dim intNumeroDocumentosEnAnexo As Integer
    Dim adoRegistro                As ADODB.Recordset
    
    If Not EsDiaUtil(gdatFechaActual) Then
        MsgBox "No se puede negociar en un día no útil!", vbCritical, Me.Caption
        Exit Sub
    End If
    
    If cboTipoInstrumento.ListCount > 0 Then

        strEstado = Reg_Adicion
        LlenarFormulario strEstado
        
        cmdOpcion.Visible = False
        cmdGuardar.Visible = True
        cmdCancelar.Visible = True
        
        InicializarValores
       
        With tabRFCortoPlazo
            .TabEnabled(0) = False
            .TabEnabled(1) = True
            .TabEnabled(2) = False
            .Tab = 1
        End With
    Else
        MsgBox "Acceso a Negociación Denegada", vbCritical, Me.Caption
    End If
    
End Sub
Private Sub LlenarFormulario(strModo As String)

    Dim intRegistro As Integer
  
    Select Case strModo

        Case Reg_Adicion
            intRegistro = 1
    End Select
  
End Sub

Public Sub Buscar()

    Dim strFechaOrdenDesde       As String, strFechaOrdenHasta        As String
    Dim strFechaLiquidacionDesde As String, strFechaLiquidacionHasta  As String
    Dim datFechaSiguiente        As Date
    Dim adoAux              As ADODB.Recordset
    
    Me.MousePointer = vbHourglass
    
    '*** Fecha Vigente, Moneda ***
    adoComm.CommandText = "{ call up_ACSelDatosParametro(23,'" & strCodFondo & "','" & gstrCodAdministradora & "','000') }"
    Set adoAux = adoComm.Execute
    
    If Not adoAux.EOF Then
        gdatFechaActual = CVDate(adoAux("FechaCuota"))
        strCodMoneda = Trim$(adoAux("CodMoneda"))
        frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
    End If

    adoAux.Close: Set adoAux = Nothing
    
    If Not IsNull(dtpFechaOrdenDesde.Value) And Not IsNull(dtpFechaOrdenHasta.Value) Then
        strFechaOrdenDesde = Convertyyyymmdd(dtpFechaOrdenDesde.Value)
        datFechaSiguiente = DateAdd("d", 1, dtpFechaOrdenHasta.Value)
        strFechaOrdenHasta = Convertyyyymmdd(datFechaSiguiente)
    End If
    
    strSQL = "SELECT IOR.NumOrden,FechaOrden,Nemotecnico,NumDocumento,EstadoOrden,IOR.CodFile,CodAnalitica,TipoOrden,IOR.CodMoneda," & _
       "(RTRIM(DescripParametro) + SPACE(1) + DescripOrden) DescripOrden,CantOrden,ValorNominal,PorcenDsctoValorNominal, " & _
       "ValorNominalDscto,MontoInteres, MontoImptoInteres," & _
       "MontoVencimiento, CodSigno DescripMoneda, IOR.CodDetalleFile,  IOR.CodFondo," & _
       "IP1.DescripPersona DesAcreedor, IP2.DescripPersona DesGestor, IOR.CodAcreedor, IOR.CodGestor  " & _
       "FROM FinanciamientoOrden IOR JOIN AuxiliarParametro AUX ON(AUX.CodParametro=IOR.TipoOrden AND AUX.CodTipoParametro = 'OPECAJ') " & _
       "JOIN Moneda MON ON(MON.CodMoneda=IOR.CodMoneda) " & _
       "LEFT JOIN InstitucionPersona IP1 ON (IP1.CodPersona = IOR.CodAcreedor AND IP1.TipoPersona = '" & Codigo_Tipo_Persona_Emisor & "') " & _
       "LEFT JOIN InstitucionPersona IP2 ON (IP2.CodPersona = IOR.CodGestor AND IP2.TipoPersona = '" & Codigo_Tipo_Persona_Emisor & "') " & _
       "WHERE IOR.TipoOrden = '" & Codigo_Orden_Compra & "' AND  IOR.CodAdministradora='" & gstrCodAdministradora & "' AND IOR.CodFondo='" & gstrCodFondoContable & "' "
        
    If strCodTipoInstrumento <> Valor_Caracter Then
        strSQL = strSQL & "AND IOR.CodFile='" & strCodTipoInstrumento & "' "
    Else
        strSQL = strSQL & "AND IOR.CodFile IN ('" & CodFile_Financiamiento_Prestamos & "')"
    End If

    If Not IsNull(dtpFechaOrdenDesde.Value) And Not IsNull(dtpFechaOrdenHasta.Value) Then
        strSQL = strSQL & "AND (FechaOrden >='" & strFechaOrdenDesde & "' AND FechaOrden <'" & strFechaOrdenHasta & "') "
    End If
        
    If strCodEstado <> Valor_Caracter Then
        strSQL = strSQL & "AND EstadoOrden='" & strCodEstado & "' "
    End If
        
    strSQL = strSQL & "ORDER BY IOR.NumOrden"
    
    With adoConsulta
        .ConnectionString = gstrConnectConsulta
        .RecordSource = strSQL
        .Refresh
    End With

    tdgConsulta.Refresh

    If adoConsulta.Recordset.RecordCount > 0 Then strEstado = Reg_Consulta

    Me.MousePointer = vbDefault
    
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

Public Sub Grabar()

    Call Accion(vSave)

End Sub

Private Sub cmdCancelar_Click()
    Call Cancelar
End Sub
Private Sub cboEstado_Click()

    strCodEstado = Valor_Caracter

    If cboEstado.ListIndex < 0 Then Exit Sub
    
    strCodEstado = Trim$(arrEstado(cboEstado.ListIndex))
    
    Call Buscar
End Sub
Private Sub cmdEnviar_Click()

    Dim strFechaDesde As String
    Dim intRegistro   As Integer, intContador         As Integer
    
    If adoConsulta.Recordset.RecordCount = 0 Then Exit Sub
    
    strFechaDesde = Convertyyyymmdd(dtpFechaOrdenDesde.Value)
    
    intContador = tdgConsulta.SelBookmarks.Count - 1
    
    If intContador < 0 Then
        MsgBox "No se ha seleccionado ningún registro", vbCritical, Me.Caption
        Exit Sub
    End If
        
    For intRegistro = 0 To intContador
        tdgConsulta.Row = tdgConsulta.SelBookmarks(intRegistro) - 1
               
        If strCodEstado = Estado_Orden_Ingresada Then
            adoComm.CommandText = "UPDATE FinanciamientoOrden SET EstadoOrden='" & Estado_Orden_Enviada & "'," & "UsuarioEdicion='" & gstrLogin & "',FechaEdicion='" & gstrFechaActual & Space$(1) & Format$(Time, "hh:mm") & "' " & "WHERE NumOrden='" & Trim$(tdgConsulta.Columns(0)) & "' AND CodFondo='" & gstrCodFondoContable & "' AND " & "CodAdministradora='" & gstrCodAdministradora & "' AND EstadoOrden='" & Estado_Orden_Ingresada & "'"
        ElseIf strCodEstado = Estado_Orden_Enviada Then
            adoComm.CommandText = "UPDATE FinanciamientoOrden SET EstadoOrden='" & Estado_Orden_Ingresada & "'," & "UsuarioEdicion='" & gstrLogin & "',FechaEdicion='" & gstrFechaActual & Space$(1) & Format$(Time, "hh:mm") & "' " & "WHERE NumOrden='" & Trim$(tdgConsulta.Columns(0)) & "' AND CodFondo='" & gstrCodFondoContable & "' AND " & "CodAdministradora='" & gstrCodAdministradora & "' AND EstadoOrden='" & Estado_Orden_Enviada & "'"
        End If

        adoConn.Execute adoComm.CommandText
    Next
    
    If strCodEstado = Estado_Orden_Ingresada Then
        MsgBox Mensaje_Envio_Exitoso, vbExclamation, gstrNombreEmpresa
    ElseIf strCodEstado = Estado_Orden_Enviada Then
        MsgBox Mensaje_Desenvio_Exitoso, vbExclamation, gstrNombreEmpresa
    ElseIf strCodEstado = Estado_Orden_Procesada Or strCodEstado = "" Then
        MsgBox "Las órdenes seleccionadas ya han sido confirmadas.", vbExclamation, gstrNombreEmpresa
    ElseIf strCodEstado = Estado_Orden_Anulada Then
        MsgBox "No puede enviarse a backoffice una orden anulada.", vbExclamation, gstrNombreEmpresa
    End If

    Call Buscar
End Sub

Private Sub cmdGuardar_Click()
    Call GrabarNew
End Sub

Private Sub Cancelar()
    cmdOpcion.Visible = True
    cmdCancelar.Visible = False

    With tabRFCortoPlazo
        .TabEnabled(0) = True
        .TabEnabled(1) = False
        .TabEnabled(2) = False
        .Tab = 0
    End With

    Call Buscar
End Sub

Public Sub Eliminar()

    Dim strNumOrden  As String
    Dim intRegistro  As Integer

    For intRegistro = 0 To tdgConsulta.SelBookmarks.Count - 1
        adoConsulta.Recordset.MoveFirst
        adoConsulta.Recordset.Move CLng(tdgConsulta.SelBookmarks(intRegistro) - 1), 0
        tdgConsulta.Refresh
        
        strNumOrden = Trim$(adoConsulta.Recordset("NumOrden"))
        strCodEstado = Trim$(adoConsulta.Recordset("EstadoOrden"))
        
        If strEstado = Reg_Consulta Or strEstado = Reg_Edicion Then
            Dim strMensaje As String
            
            'verificar si la orden no está ya anulada
            
            If strCodEstado <> Estado_Orden_Anulada And strCodEstado <> Estado_Orden_Procesada Then
            
                strMensaje = "Se procederá a eliminar la ORDEN " & strNumOrden & " por la " & tdgConsulta.Columns(3) & vbNewLine & vbNewLine & vbNewLine & "¿ Seguro de continuar ?"
                
                If MsgBox(strMensaje, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
            
                    '*** Anular Orden ***
                    adoComm.CommandText = "UPDATE FinanciamientoOrden SET EstadoOrden='" & Estado_Orden_Anulada & "'," & "UsuarioEdicion='" & gstrLogin & "',FechaEdicion='" & gstrFechaActual & Space$(1) & Format$(Time, "hh:mm") & "' " & "WHERE CodFondo='" & gstrCodFondoContable & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND NumOrden='" & strNumOrden & "'"
                        
                    adoConn.Execute adoComm.CommandText
                    
                    MsgBox Mensaje_Eliminacion_Exitosa, vbExclamation, Me.Caption
                    
                    tabRFCortoPlazo.TabEnabled(0) = True
                    tabRFCortoPlazo.Tab = 0
                    Call Buscar
                    
                    Exit Sub
                End If
                
            Else
    
                If strCodEstado = Estado_Orden_Anulada Then
                    MsgBox "La orden " & strNumOrden & " ya ha sido anulada.", vbExclamation, "Anular Orden"
                Else
                    MsgBox "La orden " & strNumOrden & " ya ha sido procesada." & vbNewLine & "No se puede anular.", vbCritical, "Anular Orden"
                End If
            End If
        
        End If
    Next
    
End Sub

Private Sub GrabarNew()

    Dim strMsgError     As String
    If (strEstado = Reg_Adicion) And (TodoOK()) Then
        On Error GoTo CtrlError
        strEstadoOrden = Estado_Orden_Ingresada
        
        With adoComm
            .CommandText = "{ call up_FIAdicFinanciamientoOrden('" & gstrCodFondoContable & "','" & gstrCodAdministradora & "','" & strNumOrden & "','" & _
                            Convertyyyymmdd(dtpFechaOrden.Value) & "','" & Estado_Orden_Ingresada & "','" & Codigo_Orden_Compra & "','" & gstrPeriodoActual & "','" & _
                            gstrMesActual & "','','" & CodFile_Financiamiento_Prestamos & "','','" & strCodClaseInstrumento & "','" & txtDescripOrden.Text & _
                            "','" & txtNemonico.Text & "','" & strCodAcreedor & "','" & strCodGestor & "','" & Convertyyyymmdd(dtpFechaEmisionDocumento) & _
                            "','19000101','19000101','" & Convertyyyymmdd(dtpFechaVencimiento.Value) & "'," & txtDiasPlazo.Text & ",'" & Valor_Indicador & _
                            "','','','','" & txtNumDocDscto.Text & "','" & strCodMoneda & "','" & strCodMoneda & "'," & txtValorNominalDocumento.Value & ",1,1," & _
                            txtValorNominalDocumento.Value & ",100," & txtValorNominalDocumento.Value & "," & txtTasa.Value & "," & txtTasaMora.Text & _
                            ",'" & strTipoTasa & "','" & strPeriodoTasa & "','" & strPeriodoCapitalizacion & "','" & strCodBaseAnual & "','" & _
                            strTipoAmortizacion & "','" & strPeriodoCuota & "'," & txtUnidadesPeriodo.Value & ",'" & strIndFinPeriodo & "','" & _
                            strIndIGV & "','19000101','" & strIndFechaAPartir & "','" & Convertyyyymmdd(dtpAPartir.Value) & "'," & _
                            txtDiasMinimosPagoInteres.Value & ",'" & strCodDesplazamientoCorte & "','" & strCodDesplazamientoPago & "'," & _
                            txtPeriodoGracia.Value & "," & intCantTramos & ",'" & strTipoTramo & "'," & txtInteresTotal.Value & "," & _
                            dblPorcenImptoInteres & "," & txtIGVInteresTotal.Value & ",0,0," & txtPorcenComision.Value & "," & txtMontoComision.Value & _
                            "," & gdblTasaIgv * 100 & "," & txtIGVComision.Value & "," & txtTotalPago.Value & "," & txtTotalRecibido.Value & _
                            ",0,0,0,0,0,0,'" & txtObservacion.Text & "','" & gstrLogin & "','" & gstrFechaActual
                            
            If strTipoAmortizacion = "04" Then
          
                .CommandText = .CommandText & "','" & strTramoXML
            End If
            
            .CommandText = .CommandText & "')}"
            
            adoConn.Execute .CommandText
        
        End With
        
        Me.MousePointer = vbDefault
                
        MsgBox Mensaje_Adicion_Exitosa, vbInformation
          
        With tabRFCortoPlazo
            .TabEnabled(0) = True
            .TabEnabled(1) = False
            .TabEnabled(2) = False
            .Tab = 0
            cmdOpcion.Visible = True
            cmdCancelar.Visible = False
        End With
        
        Call Buscar
    End If
    
    Exit Sub
    
CtrlError:
  
    Me.MousePointer = vbDefault
  
    strMsgError = strMsgError & err.Description
    MsgBox strMsgError, vbCritical, "Error"
        
End Sub

Private Function TodoOK() As Boolean
    TodoOK = True
    If strCodClaseInstrumento = Valor_Caracter Then
        MsgBox "Debe seleccionar una clase de instrumento.", vbCritical, Me.Caption
        TodoOK = False
        Exit Function
    End If
    
    If strCodAcreedor = Valor_Caracter Then
        MsgBox "Debe seleccionar un acreedor.", vbCritical, Me.Caption
        TodoOK = False
        Exit Function
    End If
    If strCodGestor = Valor_Caracter Then
        MsgBox "Debe seleccionar un gestor.", vbCritical, Me.Caption
        TodoOK = False
        Exit Function
    End If
    
    If txtDiasPlazo.Value < 1 Then
        MsgBox "El plazo en días no puede ser 0.", vbCritical, Me.Caption
        TodoOK = False
        Exit Function
    End If
    
    If txtValorNominalDocumento.Value < 1 Then
        MsgBox "El valor nominal no puede ser 0.", vbCritical, Me.Caption
        TodoOK = False
        Exit Function
    End If
    
    If txtTasa.Value <= 0 Then
        MsgBox "La tasa de interés no puede ser 0.", vbCritical, Me.Caption
        TodoOK = False
        Exit Function
    End If
    
    If strCodMoneda = Valor_Caracter Then
        MsgBox "Debe seleccionar la moneda del préstamo.", vbCritical, Me.Caption
        TodoOK = False
        Exit Function
    End If
    
    If txtNumDocDscto.Text = Valor_Caracter Then
        MsgBox "Debe indicar el Nº de documento de referencia.", vbCritical, Me.Caption
        TodoOK = False
        Exit Function
    End If
    
    
    If txtTasaMora.Value <= 0 Then
        MsgBox "Se recomienda configurar una tasa moratoria para los intereses adicionales.", vbInformation, Me.Caption
        Exit Function
    End If
    
End Function


Private Sub cmdSiguiente_Click()
    Dim objTramoXML     As DOMDocument60
    Dim strMsgError     As String

    'Comprobar si todo OK
    If Not TodoOK() Then Exit Sub
    
    If chkAPartir.Value = vbChecked Then
        datFechaInicioCalendario = dtpAPartir.Value
    Else
        datFechaInicioCalendario = dtpFechaOrden.Value
    End If
    
    With adoComm
        'calcular numero de cuotas
        .CommandText = "select dbo.uf_IVCalculoCantidadCuotas('" & Convertyyyymmdd(datFechaInicioCalendario) & "','" & _
                        Convertyyyymmdd(dtpFechaVencimiento.Value) & "','" & strCodBaseAnual & "','" & strPeriodoCuota & _
                        "'," & txtUnidadesPeriodo.Value & ",'" & strIndFinPeriodo & "','" & strCodDesplazamientoCorte & "') as CantCuotas"
        Set adoAuxiliar = .Execute
        
        If Not adoAuxiliar.EOF Then
            intCantCuotas = CInt(adoAuxiliar("CantCuotas").Value)
        Else
            intCantCuotas = 0
        End If
    End With
        
    txtCantCuotas.Text = intCantCuotas
    
    strSQL = "{ call up_IVGeneraCalendarioCuotas('" & gstrCodFondoContable & "'," & intCantCuotas & ",'" & Convertyyyymmdd(datFechaInicioCalendario) & _
            "','" & Convertyyyymmdd(dtpFechaVencimiento.Value) & "'," & txtValorNominalDocumento.Value & "," & txtTasa.Value & ",'" & _
            strTipoTasa & "','" & strPeriodoTasa & "','" & strPeriodoCapitalizacion & "','" & strCodBaseAnual & "','" & strIndIGV & "','" & _
            strTipoAmortizacion & "','" & strPeriodoCuota & "'," & txtUnidadesPeriodo.Value & ",'" & strIndFinPeriodo & "'," & _
            txtPeriodoGracia.Value & ",'" & strCodDesplazamientoCorte & "','" & strCodDesplazamientoPago
    
    If strTipoAmortizacion = "04" Then
        'xml de tramos
        Call XMLADORecordset(objTramoXML, "TramoCuota", "Tramo", adoTramos, strMsgError)
        strTramoXML = objTramoXML.xml
        strSQL = strSQL & "','" & strTramoXML
    End If
    
    strSQL = strSQL & "')}"
        
    Set adoCalendario = New ADODB.Recordset
    With adoCalendario
        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .Open strSQL
    End With

    tdgCalendario.DataSource = adoCalendario
    tdgCalendario.Refresh
    
    If adoCalendario.EOF Then Exit Sub
    
    Dim SumaInteres As Double
    Dim sumaigv As Double
    Dim SumaCuotas As Double
    
    SumaInteres = 0
    sumaigv = 0
    SumaCuotas = 0
    
    adoCalendario.MoveFirst
    While Not adoCalendario.EOF
        SumaInteres = SumaInteres + adoCalendario("Interes")
        sumaigv = sumaigv + adoCalendario("IGVInteres")
        SumaCuotas = SumaCuotas + adoCalendario("Cuota")
        adoCalendario.MoveNext
    Wend
    
    adoCalendario.MoveFirst
    
    txtValorNominalDscto.Text = txtValorNominalDocumento.Value
    txtInteresTotal.Text = SumaInteres
    txtIGVInteresTotal.Text = sumaigv
    txtTotalPago.Text = SumaCuotas
    
    Call CalcularTotalRecibido

     With tabRFCortoPlazo
        .TabEnabled(0) = False
        .TabEnabled(1) = False
        .TabEnabled(2) = True
        .Tab = 2
        cmdCancelar.Visible = True
    End With
    
End Sub

Private Sub GenerarTramos()
    If Not blnFormReady Then Exit Sub
    
    Dim intNumTramo     As Integer
    Dim intCuotaInicio  As Integer
    Dim intCuotaFin     As Integer
    
    Set adoTramos = New ADODB.Recordset
    
      With adoTramos.Fields
        .Append "TipoTramo", adVarChar, 2
        .Append "NumTramo", adInteger
        .Append "InicioTramo", adInteger
        .Append "FinTramo", adInteger
        .Append "Valor", adDecimal, 19
    
        .Item("Valor").Precision = 19
        .Item("Valor").NumericScale = 2
       
    End With
    
    intNumTramo = 0
    intCantTramos = txtCantidadTramos.Value
    
    adoTramos.Open
    
    While intNumTramo < intCantTramos
        intNumTramo = intNumTramo + 1
        intCuotaInicio = intNumTramo
        
        If intNumTramo = intCantTramos Then
            intCuotaFin = intCantCuotas
        Else
            intCuotaFin = intNumTramo
        End If
        
        adoTramos.AddNew Array("TipoTramo", "NumTramo", "InicioTramo", "FinTramo", "Valor"), Array(strTipoTramo, intNumTramo, intCuotaInicio, intCuotaFin, 0)

    Wend
    
    tdgTramos.DataSource = adoTramos
    
End Sub

Private Sub dtpAPartir_Change()
    If dtpAPartir.Value < dtpFechaOrden.Value Then
        dtpAPartir.Value = dtpFechaOrden.Value
    End If
    If dtpAPartir.Value > dtpFechaVencimiento.Value Then
        dtpAPartir.Value = dtpFechaVencimiento.Value
    End If
    If chkAPartir.Value = vbChecked Then
        datFechaInicioCalendario = dtpAPartir.Value
        txtDiasPlazo.Text = DateDiff("d", datFechaInicioCalendario, dtpFechaVencimiento.Value)
    End If
     
    Call CalcularCantidadCuotas
    If strTipoAmortizacion = "04" Then
        Call GenerarTramos
    End If
    
End Sub

Private Sub dtpFechaVencimiento_Change()
    If dtpFechaVencimiento.Value < dtpFechaOrden.Value Then
        dtpFechaVencimiento.Value = dtpFechaOrden.Value
    End If
    txtDiasPlazo.Text = DateDiff("d", datFechaInicioCalendario, dtpFechaVencimiento.Value)
    
    Call CalcularCantidadCuotas
    If strTipoAmortizacion = "04" Then
        Call GenerarTramos
    End If
    
End Sub

Private Sub Form_Activate()

    frmMainMdi.stbMdi.Panels(3).Text = Me.Caption
    'Call CargarReportes
    
End Sub

Private Sub Form_Deactivate()

    'Call OcultarReportes
    
End Sub

Private Sub Form_Load()
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me

    
    blnCargadoDesdeCartera = False
    blnFormReady = False
    
    Call InicializarValores
    Call CargarListas
    'Call CargarReportes
    
     With tabRFCortoPlazo
        .TabEnabled(0) = True
        .TabEnabled(1) = False
        .TabEnabled(2) = False
        .Tab = 0
        cmdCancelar.Visible = False
        cmdOpcion.Visible = True
    End With
    
    Call Buscar

    Call ValidarPermisoUsoControl(Trim$(gstrLogin), Me, Trim$(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)
    
    Call ValidaExisteTipoCambio(Codigo_TipoCambio_SBS, gstrFechaActual)
    
    CentrarForm Me
    blnFormReady = True
End Sub

Private Sub CargarListas()
    Dim adoRecord   As ADODB.Recordset
    Dim intRegistro As Integer
    
    '*** Fondos ***
    strSQL = "SELECT CodFondo CODIGO,DescripFondo DESCRIP FROM Fondo WHERE CodAdministradora='" & gstrCodAdministradora & "' AND Estado='01' and CodFondo = '" & gstrCodFondoContable & "' ORDER BY DescripFondo"
    CargarControlLista strSQL, cboFondo, arrFondo(), Valor_Caracter
    CargarControlLista strSQL, cboFondoOrden, arrFondoOrden(), Valor_Caracter
        
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
    If cboFondoOrden.ListCount > 0 Then cboFondoOrden.ListIndex = 0
    
    Dim adoAux As ADODB.Recordset
    adoComm.CommandText = "SELECT CodMoneda FROM Fondo WHERE CodAdministradora='" & gstrCodAdministradora & "' AND Estado='01' and CodFondo = '" & gstrCodFondoContable & "' ORDER BY DescripFondo"
    Set adoAux = adoComm.Execute
    If Not adoAux.EOF Then
        gstrCodMoneda = adoAux("CodMoneda")
    End If
        
    '*** Tipo de Instrumento ***
    strSQL = "SELECT CodFile CODIGO,DescripFile DESCRIP " & "FROM InversionFile " & "WHERE IndVigente='X' AND CodFile = '" & CodFile_Financiamiento_Prestamos & "' ORDER BY CODIGO"
    CargarControlLista strSQL, cboTipoInstrumentoOrden, arrTipoInstrumentoOrden(), Valor_Caracter
    If cboTipoInstrumentoOrden.ListCount > 0 Then cboTipoInstrumentoOrden.ListIndex = 0
    
    '*** Tipo de Instrumento ***
    strSQL = "SELECT CodFile CODIGO,DescripFile DESCRIP " & "FROM InversionFile " & "WHERE IndVigente='X' AND CodFile = '" & CodFile_Financiamiento_Prestamos & "' ORDER BY CODIGO"
    CargarControlLista strSQL, cboTipoInstrumento, arrTipoInstrumento(), Valor_Caracter
    If cboTipoInstrumento.ListCount > 0 Then cboTipoInstrumento.ListIndex = 0
    

    '*** Estados de la Orden ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='ESTORD' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboEstado, arrEstado(), Sel_Todos
    If cboEstado.ListCount > 0 Then cboEstado.ListIndex = 0

    intRegistro = ObtenerItemLista(arrEstado(), Estado_Orden_Ingresada)

    If intRegistro >= 0 Then cboEstado.ListIndex = intRegistro
        
    '*** Acreedor ***
    strSQL = "SELECT (CodPersona) CODIGO,DescripPersona DESCRIP FROM InstitucionPersona WHERE TipoPersona='" & Codigo_Tipo_Persona_Emisor & "' AND IndVigente='X' and IndBanco = 'X' ORDER BY DescripPersona"
    CargarControlLista strSQL, cboAcreedor, arrAcreedor(), Sel_Defecto
    If cboAcreedor.ListCount > 0 Then cboAcreedor.ListIndex = 0
    
    '*** Gestor ***
    strSQL = "SELECT (CodPersona) CODIGO,DescripPersona DESCRIP FROM InstitucionPersona WHERE TipoPersona='" & Codigo_Tipo_Persona_Emisor & "' AND IndVigente='X' AND IndBanco = 'X' ORDER BY DescripPersona"
    CargarControlLista strSQL, cboGestor, arrGestor(), Sel_Defecto
    If cboGestor.ListCount > 0 Then cboGestor.ListIndex = 0
                
    '*** Moneda ***
    strSQL = "{ call up_ACSelDatos(2) }"
    CargarControlLista strSQL, cboMoneda, arrMoneda(), Valor_Caracter
    If cboMoneda.ListCount > 0 Then cboMoneda.ListIndex = ObtenerItemLista(arrMoneda(), gstrCodMoneda)

    '*** Base de Calculo
    strSQL = "{ call up_ACSelDatos(45) }"
    CargarControlLista strSQL, cboBaseAnual, arrBaseAnual(), Valor_Caracter
    If cboBaseAnual.ListCount > 0 Then cboBaseAnual.ListIndex = 0
    
    '*** Tipo Tasa ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='NATTAS' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboTipoTasa, arrTipoTasa(), Valor_Caracter
    If cboTipoTasa.ListCount > 0 Then cboTipoTasa.ListIndex = 0
    
    '*** Periodo de Tasa ***
    strSQL = "SELECT CodParametro CODIGO, DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro = 'TIPFRE'"
    CargarControlLista strSQL, cboPeriodoTasa, arrPeriodoTasa(), Valor_Caracter
    cboPeriodoTasa.ListIndex = 0
    
    '*** Periodo de Capitalizacion ***
    strSQL = "SELECT CodParametro CODIGO, DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro = 'TIPFRE'"
    CargarControlLista strSQL, cboPeriodoCapitalizacion, arrPeriodoCapitalizacion(), Sel_NoAplicable
    If cboPeriodoCapitalizacion.ListCount > 0 Then cboPeriodoCapitalizacion.ListIndex = 0

    '-----unidad de periodo
    strSQL = "{ call up_ACSelDatos(46) }"
    CargarControlLista strSQL, cboUnidadesPeriodo, arrPeriodoCuota(), Valor_Caracter
    If cboUnidadesPeriodo.ListCount > 0 Then cboUnidadesPeriodo.ListIndex = 0
    
    '-----desplazamiento
    strSQL = "{ call up_ACSelDatos(47) }"
    CargarControlLista strSQL, cboDesplazamientoCorte, arrDesplazamientoCorte(), Valor_Caracter
    CargarControlLista strSQL, cboDesplazamientoPago, arrDesplazamientoPago(), Valor_Caracter
    If cboDesplazamientoCorte.ListCount > 0 Then cboDesplazamientoCorte.ListIndex = 0
    If cboDesplazamientoPago.ListCount > 0 Then cboDesplazamientoPago.ListIndex = 0
    
    '-----modalidad de amortizacion
    strSQL = "{ call up_ACSelDatos(53) }"
    CargarControlLista strSQL, cboAmortizacion, arrTipoAmortizacion(), Valor_Caracter
    If cboAmortizacion.ListCount > 0 Then cboAmortizacion.ListIndex = 0


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

Public Sub Salir()

    Unload Me
    
End Sub

Private Sub CalcularCantidadCuotas()
    If Not blnFormReady Then Exit Sub
        
    With adoComm
        'calcular numero de cuotas
        .CommandText = "select dbo.uf_IVCalculoCantidadCuotas('" & Convertyyyymmdd(datFechaInicioCalendario) & "','" & _
                        Convertyyyymmdd(dtpFechaVencimiento.Value) & "','" & strCodBaseAnual & "','" & strPeriodoCuota & _
                        "'," & txtUnidadesPeriodo.Value & ",'" & strIndFinPeriodo & "','" & strCodDesplazamientoCorte & "') as CantCuotas"
        Set adoAuxiliar = .Execute
        
        If Not adoAuxiliar.EOF Then
            intCantCuotas = CInt(adoAuxiliar("CantCuotas").Value)
        Else
            intCantCuotas = 0
        End If
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmOrdenPrestamo = Nothing
   ' Call OcultarReportes
    frmMainMdi.stbMdi.Panels(3).Text = "Acción"
    
End Sub

Private Sub optAmortizacion_Click()
    If optAmortizacion.Value Then
        strTipoTramo = "01"
    End If
    adoTramos.MoveFirst
    While Not adoTramos.EOF
        adoTramos("TipoTramo").Value = strTipoTramo
        adoTramos.MoveNext
    Wend
    
End Sub

Private Sub optCuota_Click()
    If optAmortizacion.Value Then
        strTipoTramo = "02"
    End If
    adoTramos.MoveFirst
    While Not adoTramos.EOF
        adoTramos("TipoTramo").Value = strTipoTramo
        adoTramos.MoveNext
    Wend

End Sub

Private Sub tdgTramos_AfterColUpdate(ByVal ColIndex As Integer)
    Dim intPosicion As Long
    Dim intRow As Long
    Dim numCuota As Integer
    
    If ColIndex < 3 Then
         intPosicion = adoTramos.Bookmark
         intRow = intPosicion
         numCuota = adoTramos("FinTramo").Value
              
        
         While intRow < intCantTramos
             adoTramos.MoveNext
             intRow = intRow + 1
             numCuota = numCuota + 1
             adoTramos("InicioTramo").Value = numCuota
             
             If intRow = intCantTramos Then
                 adoTramos("FinTramo").Value = intCantCuotas
             Else
                 adoTramos("FinTramo").Value = numCuota
             End If
         Wend
         adoTramos("FinTramo").Value = intCantCuotas
         adoTramos.MoveFirst
         adoTramos.AbsolutePosition = intPosicion
    End If
End Sub

Private Sub tdgTramos_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    Dim intPosicion As Long
    Dim intRow As Long
    Dim numCuota As Integer
    
    intPosicion = adoTramos.Bookmark
    If adoTramos.Bookmark = intCantTramos Then
        tdgTramos.Splits(0).Columns(ColIndex).Locked = True
    Else
        tdgTramos.Splits(0).Columns(ColIndex).Locked = False
    End If
         
End Sub

Private Sub tdgTramos_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    Dim intPosicion As Long
    intPosicion = adoTramos.Bookmark
    If ColIndex = 2 Then
        If CInt(tdgTramos.Splits(0).Columns(ColIndex).Value) < tdgTramos.Splits(0).Columns(ColIndex - 1).Value Then
             tdgTramos.Splits(0).Columns(ColIndex).Value = OldValue
        End If
        If CInt(tdgTramos.Splits(0).Columns(ColIndex).Value) > (intCantCuotas - (intCantTramos - intPosicion)) Then
            tdgTramos.Splits(0).Columns(ColIndex).Value = OldValue
        End If
    End If
End Sub

Private Sub tdgTramos_Click()
    tdgTramos.Splits(0).Columns(2).Locked = False
    tdgTramos.Splits(0).Columns(3).Locked = False
End Sub

Private Sub tdgTramos_LostFocus()
    tdgTramos.MoveFirst
End Sub

Private Sub txtCantidadTramos_Change()
    If strTipoAmortizacion = "04" Then
       
        Call CalcularCantidadCuotas
        
        If txtCantidadTramos.Value > intCantCuotas Then
            txtCantidadTramos.Text = intCantCuotas
        ElseIf txtCantidadTramos.Value < 1 Then
            txtCantidadTramos.Text = "1"
        End If
        
        Call GenerarTramos
    End If
End Sub

Private Sub txtDiasMinimosPagoInteres_Change()
    If txtDiasMinimosPagoInteres.Value < 0 Then
        txtDiasMinimosPagoInteres.Text = "0"
    End If

End Sub

Private Sub txtDiasPlazo_Change()
    dtpFechaVencimiento.Value = DateAdd("d", txtDiasPlazo.Value, datFechaInicioCalendario)
End Sub

Private Sub txtMontoComision_LostFocus()
    If txtPorcenComision.Value <> 0 And Not blnLockPorcenComision Then
        txtPorcenComision.Text = txtMontoComision.Value / txtValorNominalDocumento.Value * 100
        txtIGVComision.Text = txtMontoComision.Value * gdblTasaIgv
        Call CalcularTotalRecibido
    ElseIf Not blnLockPorcenComision Then
        txtPorcenComision.Text = txtMontoComision.Value / txtValorNominalDocumento.Value * 100
        txtIGVComision.Text = txtMontoComision.Value * gdblTasaIgv
        Call CalcularTotalRecibido
    End If
    
End Sub

Private Sub txtNumDocDscto_Change()
    Call GenerarNemonico
End Sub

Private Sub txtPeriodoGracia_Change()
    Call CalcularCantidadCuotas
    
    If txtPeriodoGracia.Value < 0 Then
        txtPeriodoGracia.Text = "0"
    ElseIf txtPeriodoGracia.Value >= intCantCuotas Then
        txtPeriodoGracia.Text = CStr(intCantCuotas - 1)
    End If

End Sub

Private Sub txtPorcenComision_LostFocus()
    txtMontoComision.Text = txtPorcenComision.Value * txtValorNominalDocumento.Value / 100
    txtIGVComision.Text = txtMontoComision.Value * gdblTasaIgv
    Call CalcularTotalRecibido
End Sub

Private Sub CalcularTotalRecibido()
    txtTotalRecibido.Text = txtValorNominalDscto.Value - txtMontoComision.Value - txtIGVComision.Value
End Sub

'Private Sub txtPorcenComision_GotFocus()
'    blnLockPorcenComision = True
'End Sub
'
'Private Sub txtPorcenComision_LostFocus()
'    blnLockPorcenComision = False
'End Sub

Private Sub txtUnidadesPeriodo_Change()
    If txtUnidadesPeriodo.Value < 1 Then
        txtUnidadesPeriodo.Text = "1"
    End If
    
    Call CalcularCantidadCuotas
    If strTipoAmortizacion = "04" Then
        Call GenerarTramos
    End If
End Sub

Private Sub txtValorNominalDocumento_Click()
    txtValorNominalDscto.Text = txtValorNominalDocumento.Value
End Sub
