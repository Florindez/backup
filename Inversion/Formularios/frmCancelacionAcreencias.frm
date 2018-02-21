VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{830D5A06-9B70-4F8C-98B6-7A19C4E7760B}#1.0#0"; "TAMControls.ocx"
Begin VB.Form frmCancelacionAcreencias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pagos - Operaciones de Acreencias"
   ClientHeight    =   9270
   ClientLeft      =   5760
   ClientTop       =   3615
   ClientWidth     =   15375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9270
   ScaleWidth      =   15375
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
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
      Left            =   3810
      Picture         =   "frmCancelacionAcreencias.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   149
      Top             =   8460
      Width           =   1200
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
      Left            =   12630
      Picture         =   "frmCancelacionAcreencias.frx":00EA
      Style           =   1  'Graphical
      TabIndex        =   147
      Top             =   7380
      Width           =   1200
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
      Left            =   13140
      Picture         =   "frmCancelacionAcreencias.frx":064C
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   8460
      Width           =   1200
   End
   Begin VB.CommandButton cmdAnular 
      Caption         =   "Anular"
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
      Left            =   2460
      Picture         =   "frmCancelacionAcreencias.frx":0BCE
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   8460
      Width           =   1200
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "Nuevo"
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
      Left            =   1110
      Picture         =   "frmCancelacionAcreencias.frx":0CA5
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   8460
      Width           =   1200
   End
   Begin TabDlg.SSTab tabPagosRFCP 
      Height          =   8295
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   15315
      _ExtentX        =   27014
      _ExtentY        =   14631
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
      TabPicture(0)   =   "frmCancelacionAcreencias.frx":0D6F
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraCriterio"
      Tab(0).Control(1)=   "tdgConsulta"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Operación / Anexo"
      TabPicture(1)   =   "frmCancelacionAcreencias.frx":0D8B
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraDatosBasicos"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraDatosAnexo"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "fraDetalle"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdSiguiente"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Detalle de Pago"
      TabPicture(2)   =   "frmCancelacionAcreencias.frx":0DA7
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraDetallePago"
      Tab(2).Control(1)=   "cmdAnterior"
      Tab(2).Control(2)=   "fraDetallePagoParcial"
      Tab(2).Control(3)=   "fraDetalleTotalPago"
      Tab(2).Control(4)=   "cmdGuardar"
      Tab(2).ControlCount=   5
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
         Left            =   -61080
         Picture         =   "frmCancelacionAcreencias.frx":0DC3
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   7320
         Width           =   1200
      End
      Begin VB.Frame fraDetalleTotalPago 
         Caption         =   "Total de Conceptos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4305
         Left            =   -64620
         TabIndex        =   93
         Top             =   2910
         Width           =   4755
         Begin TAMControls.TAMTextBox txtTotalPrincipalAdeudado 
            Height          =   285
            Left            =   2910
            TabIndex        =   130
            Top             =   360
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
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
            Container       =   "frmCancelacionAcreencias.frx":13B7
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   2
            MaximoValor     =   1E+18
         End
         Begin TAMControls.TAMTextBox txtTotalIGVInteresAdic 
            Height          =   285
            Left            =   2910
            TabIndex        =   131
            Top             =   2520
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
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
            Container       =   "frmCancelacionAcreencias.frx":13D3
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   2
            MaximoValor     =   1E+18
         End
         Begin TAMControls.TAMTextBox txtTotalInteresAdic 
            Height          =   285
            Left            =   2910
            TabIndex        =   132
            Top             =   2160
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
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
            Container       =   "frmCancelacionAcreencias.frx":13EF
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   2
            MaximoValor     =   1E+18
         End
         Begin TAMControls.TAMTextBox txtAjusteDev 
            Height          =   285
            Left            =   2910
            TabIndex        =   133
            Top             =   1440
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
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
            Container       =   "frmCancelacionAcreencias.frx":140B
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   2
            MaximoValor     =   1E+18
         End
         Begin TAMControls.TAMTextBox txtIGVAjusteDev 
            Height          =   285
            Left            =   2910
            TabIndex        =   134
            Top             =   1800
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
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
            Container       =   "frmCancelacionAcreencias.frx":1427
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   2
            MaximoValor     =   1E+18
         End
         Begin TAMControls.TAMTextBox txtTotalIGVInteresAFavor 
            Height          =   285
            Left            =   2910
            TabIndex        =   135
            Top             =   1080
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
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
            Container       =   "frmCancelacionAcreencias.frx":1443
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   2
            MaximoValor     =   1E+18
         End
         Begin TAMControls.TAMTextBox txtTotalInteresAFavor 
            Height          =   285
            Left            =   2910
            TabIndex        =   136
            Top             =   720
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
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
            Container       =   "frmCancelacionAcreencias.frx":145F
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   2
            MaximoValor     =   1E+18
         End
         Begin TAMControls.TAMTextBox txtTotalNoConsumido 
            Height          =   285
            Left            =   2520
            TabIndex        =   137
            Top             =   3930
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   503
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
            Container       =   "frmCancelacionAcreencias.frx":147B
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   2
            MaximoValor     =   1E+18
         End
         Begin TAMControls.TAMTextBox txtMontoPagado 
            Height          =   285
            Left            =   2910
            TabIndex        =   138
            Top             =   3210
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
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
            Container       =   "frmCancelacionAcreencias.frx":1497
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   2
            MaximoValor     =   1E+18
         End
         Begin TAMControls.TAMTextBox txtTotalDeuda 
            Height          =   285
            Left            =   2910
            TabIndex        =   139
            Top             =   2880
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
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
            Container       =   "frmCancelacionAcreencias.frx":14B3
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   2
            MaximoValor     =   1E+18
         End
         Begin TAMControls.TAMTextBox txtMargenDevolucion 
            Height          =   285
            Left            =   2520
            TabIndex        =   140
            Top             =   3570
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   503
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
            Container       =   "frmCancelacionAcreencias.frx":14CF
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   2
            MaximoValor     =   1E+18
         End
         Begin TAMControls.TAMTextBox txtPorcenTotalImptoInteresAFavor 
            Height          =   285
            Left            =   1980
            TabIndex        =   141
            Top             =   1080
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
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
            Container       =   "frmCancelacionAcreencias.frx":14EB
            Text            =   "0.0000"
            Decimales       =   4
            Estilo          =   4
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   4
            MaximoValor     =   1E+18
         End
         Begin TAMControls.TAMTextBox txtPorcenTotalImptoInteresAdic 
            Height          =   285
            Left            =   1980
            TabIndex        =   142
            Top             =   2520
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
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
            Container       =   "frmCancelacionAcreencias.frx":1507
            Text            =   "0.0000"
            Decimales       =   4
            Estilo          =   4
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   4
            MaximoValor     =   1E+18
         End
         Begin TAMControls.TAMTextBox txtPorcenImptoAjusteDev 
            Height          =   285
            Left            =   1980
            TabIndex        =   143
            Top             =   1800
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
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
            Container       =   "frmCancelacionAcreencias.frx":1523
            Text            =   "0.0000"
            Decimales       =   4
            Estilo          =   4
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   4
            MaximoValor     =   1E+18
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "%"
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
            Index           =   56
            Left            =   2760
            TabIndex        =   108
            Top             =   1860
            Width           =   150
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "IGV Ajuste de Int.:"
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
            Index           =   55
            Left            =   210
            TabIndex        =   107
            Top             =   1860
            Width           =   1590
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Ajus. Int. por Pago Adelantado:"
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
            Index           =   54
            Left            =   210
            TabIndex        =   106
            Top             =   1500
            Width           =   2685
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Total No Consumido:"
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
            Index           =   53
            Left            =   210
            TabIndex        =   105
            Top             =   3990
            Width           =   1785
         End
         Begin VB.Line Line1 
            DrawMode        =   1  'Blackness
            X1              =   4590
            X2              =   120
            Y1              =   3510
            Y2              =   3510
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "%"
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
            Index           =   52
            Left            =   2760
            TabIndex        =   104
            Top             =   2550
            Width           =   150
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "%"
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
            Index           =   51
            Left            =   2760
            TabIndex        =   103
            Top             =   1140
            Width           =   150
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Margen a Devolver:"
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
            Index           =   50
            Left            =   210
            TabIndex        =   102
            Top             =   3630
            Width           =   1695
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto Pagado:"
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
            Index           =   49
            Left            =   210
            TabIndex        =   101
            Top             =   3240
            Width           =   1305
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Total Deuda:"
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
            Index           =   48
            Left            =   210
            TabIndex        =   100
            Top             =   2910
            Width           =   1125
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Total IGV Int. Adic.:"
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
            Index           =   47
            Left            =   210
            TabIndex        =   99
            Top             =   2550
            Width           =   1725
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Total Int. Adicional:"
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
            Left            =   210
            TabIndex        =   98
            Top             =   2190
            Width           =   1695
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Total IGV Int. a Fav:"
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
            Left            =   210
            TabIndex        =   97
            Top             =   1140
            Width           =   1770
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Total Interés a Favor:"
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
            Index           =   39
            Left            =   210
            TabIndex        =   96
            Top             =   780
            Width           =   1845
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Total Principal Adeudado:"
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
            Index           =   38
            Left            =   210
            TabIndex        =   95
            Top             =   420
            Width           =   2220
         End
      End
      Begin VB.Frame fraDetallePagoParcial 
         Caption         =   "Detalle de Pago Parcial de la Operación "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4305
         Left            =   -74850
         TabIndex        =   68
         Top             =   2910
         Width           =   10155
         Begin TAMControls.TAMTextBox txtPrincipalAdeudado 
            Height          =   285
            Left            =   3270
            TabIndex        =   109
            Top             =   600
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
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
            Container       =   "frmCancelacionAcreencias.frx":153F
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   2
            MaximoValor     =   1E+18
         End
         Begin VB.Frame fraDetalleDatosFuturos 
            Caption         =   "Datos futuros de la Operación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1965
            Left            =   5190
            TabIndex        =   87
            Top             =   2160
            Width           =   4845
            Begin TAMControls.TAMTextBox txtNuevoPrincipal 
               Height          =   285
               Left            =   3450
               TabIndex        =   118
               Top             =   330
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   503
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
               Container       =   "frmCancelacionAcreencias.frx":155B
               Text            =   "0.00"
               Decimales       =   2
               Estilo          =   4
               Apariencia      =   1
               Borde           =   1
               DecimalesValue  =   2
               MaximoValor     =   1E+18
            End
            Begin TAMControls.TAMTextBox txtNuevoInteresAFavor 
               Height          =   285
               Left            =   3450
               TabIndex        =   119
               Top             =   690
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   503
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
               Container       =   "frmCancelacionAcreencias.frx":1577
               Text            =   "0.00"
               Decimales       =   2
               Estilo          =   4
               Apariencia      =   1
               Borde           =   1
               DecimalesValue  =   2
               MaximoValor     =   1E+18
            End
            Begin TAMControls.TAMTextBox txtImptoNuevoInteresAFavor 
               Height          =   285
               Left            =   3450
               TabIndex        =   120
               Top             =   1050
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   503
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
               Container       =   "frmCancelacionAcreencias.frx":1593
               Text            =   "0.00"
               Decimales       =   2
               Estilo          =   4
               Apariencia      =   1
               Borde           =   1
               DecimalesValue  =   2
               MaximoValor     =   1E+18
            End
            Begin TAMControls.TAMTextBox txtNuevaDeudaParcial 
               Height          =   285
               Left            =   3450
               TabIndex        =   121
               Top             =   1410
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   503
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
               Container       =   "frmCancelacionAcreencias.frx":15AF
               Text            =   "0.00"
               Decimales       =   2
               Estilo          =   4
               Apariencia      =   1
               Borde           =   1
               DecimalesValue  =   2
               MaximoValor     =   1E+18
            End
            Begin TAMControls.TAMTextBox txtPorcenImptoNuevoInteresAFavor 
               Height          =   285
               Left            =   2430
               TabIndex        =   122
               Top             =   1050
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   503
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
               Container       =   "frmCancelacionAcreencias.frx":15CB
               Text            =   "0.0000"
               Decimales       =   4
               Estilo          =   4
               Apariencia      =   1
               Borde           =   1
               DecimalesValue  =   4
               MaximoValor     =   1E+18
            End
            Begin VB.Label lblDescrip 
               AutoSize        =   -1  'True
               Caption         =   "%"
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
               Index           =   37
               Left            =   3240
               TabIndex        =   92
               Top             =   1080
               Width           =   150
            End
            Begin VB.Label lblDescrip 
               AutoSize        =   -1  'True
               Caption         =   "Nueva Deuda:"
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
               Index           =   31
               Left            =   210
               TabIndex        =   91
               Top             =   1470
               Width           =   1245
            End
            Begin VB.Label lblDescrip 
               AutoSize        =   -1  'True
               Caption         =   "Nuevo IGV de Int.a Fav.:"
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
               TabIndex        =   90
               Top             =   1110
               Width           =   2160
            End
            Begin VB.Label lblDescrip 
               AutoSize        =   -1  'True
               Caption         =   "Nuevo Interés a Favor:"
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
               Left            =   210
               TabIndex        =   89
               Top             =   750
               Width           =   1965
            End
            Begin VB.Label lblDescrip 
               AutoSize        =   -1  'True
               Caption         =   "Nuevo Principal Adeudado:"
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
               Left            =   210
               TabIndex        =   88
               Top             =   390
               Width           =   2340
            End
         End
         Begin VB.Frame fraDetalleNotaCreditoParcial 
            Caption         =   "Generación de Nota de Crédito"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1995
            Left            =   5190
            TabIndex        =   81
            Top             =   150
            Width           =   4845
            Begin TAMControls.TAMTextBox txtInteresAjuste 
               Height          =   285
               Left            =   3450
               TabIndex        =   114
               Top             =   360
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   503
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
               Container       =   "frmCancelacionAcreencias.frx":15E7
               Text            =   "0.00"
               Decimales       =   2
               Estilo          =   4
               Apariencia      =   1
               Borde           =   1
               DecimalesValue  =   2
               MaximoValor     =   1E+18
            End
            Begin TAMControls.TAMTextBox txtImptoInteresAjuste 
               Height          =   285
               Left            =   3450
               TabIndex        =   115
               Top             =   720
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   503
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
               Container       =   "frmCancelacionAcreencias.frx":1603
               Text            =   "0.00"
               Decimales       =   2
               Estilo          =   4
               Apariencia      =   1
               Borde           =   1
               DecimalesValue  =   2
               MaximoValor     =   1E+18
            End
            Begin TAMControls.TAMTextBox txtAmortizacionPrincipal 
               Height          =   285
               Left            =   3450
               TabIndex        =   116
               Top             =   1080
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   503
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
               Container       =   "frmCancelacionAcreencias.frx":161F
               Text            =   "0.00"
               Decimales       =   2
               Estilo          =   4
               Apariencia      =   1
               Borde           =   1
               DecimalesValue  =   2
               MaximoValor     =   1E+18
            End
            Begin TAMControls.TAMTextBox txtTotalNotaCredito 
               Height          =   285
               Left            =   3450
               TabIndex        =   117
               Top             =   1440
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   503
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
               Container       =   "frmCancelacionAcreencias.frx":163B
               Text            =   "0.00"
               Decimales       =   2
               Estilo          =   4
               Apariencia      =   1
               Borde           =   1
               DecimalesValue  =   2
               MaximoValor     =   1E+18
            End
            Begin TAMControls.TAMTextBox txtPorcenImptoInteresAjuste 
               Height          =   285
               Left            =   2430
               TabIndex        =   125
               Top             =   720
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   503
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
               Container       =   "frmCancelacionAcreencias.frx":1657
               Text            =   "0.0000"
               Decimales       =   4
               Estilo          =   4
               Apariencia      =   1
               Borde           =   1
               DecimalesValue  =   4
               MaximoValor     =   1E+18
            End
            Begin VB.Label lblDescrip 
               AutoSize        =   -1  'True
               Caption         =   "%"
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
               Left            =   3240
               TabIndex        =   86
               Top             =   750
               Width           =   150
            End
            Begin VB.Label lblDescrip 
               AutoSize        =   -1  'True
               Caption         =   "Total de Nota de Crédito"
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
               Left            =   210
               TabIndex        =   85
               Top             =   1530
               Width           =   2100
            End
            Begin VB.Label lblDescrip 
               AutoSize        =   -1  'True
               Caption         =   "IGV Ajuste de Interés:"
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
               TabIndex        =   84
               Top             =   780
               Width           =   1875
            End
            Begin VB.Label lblDescrip 
               AutoSize        =   -1  'True
               Caption         =   "Amortización del Principal:"
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
               Left            =   210
               TabIndex        =   83
               Top             =   1170
               Width           =   2250
            End
            Begin VB.Label lblDescrip 
               AutoSize        =   -1  'True
               Caption         =   "Ajuste de Interés:"
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
               Left            =   210
               TabIndex        =   82
               Top             =   420
               Width           =   1500
            End
         End
         Begin VB.Frame fraDetalleDeudaPagoParcial 
            Caption         =   "Deuda total de la operación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1125
            Left            =   90
            TabIndex        =   78
            Top             =   3000
            Width           =   5025
            Begin TAMControls.TAMTextBox txtDeudaTotal 
               Height          =   285
               Left            =   3180
               TabIndex        =   126
               Top             =   270
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   503
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
               Container       =   "frmCancelacionAcreencias.frx":1673
               Text            =   "0.00"
               Decimales       =   2
               Estilo          =   4
               Apariencia      =   1
               Borde           =   1
               DecimalesValue  =   2
               MaximoValor     =   1E+18
            End
            Begin TAMControls.TAMTextBox txtMontoPagoParcial 
               Height          =   285
               Left            =   3180
               TabIndex        =   127
               Top             =   630
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   503
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
               Container       =   "frmCancelacionAcreencias.frx":168F
               Text            =   "0.00"
               Decimales       =   2
               Estilo          =   4
               Apariencia      =   1
               Borde           =   1
               DecimalesValue  =   2
               MaximoValor     =   1E+18
            End
            Begin VB.Label lblDescrip 
               AutoSize        =   -1  'True
               Caption         =   "Deuda Total:"
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
               TabIndex        =   80
               Top             =   360
               Width           =   1125
            End
            Begin VB.Label lblDescrip 
               AutoSize        =   -1  'True
               Caption         =   "Monto del Pago Parcial:"
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
               Index           =   32
               Left            =   210
               TabIndex        =   79
               Top             =   690
               Width           =   2055
            End
         End
         Begin VB.Frame fraDetallePagoFF 
            Caption         =   "A Cancelar por Pago Fuera de Fecha"
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
            Left            =   90
            TabIndex        =   74
            Top             =   1740
            Width           =   5025
            Begin TAMControls.TAMTextBox txtImptoInteresAdicional 
               Height          =   285
               Left            =   3180
               TabIndex        =   112
               Top             =   690
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   503
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
               Container       =   "frmCancelacionAcreencias.frx":16AB
               Text            =   "0.00"
               Decimales       =   2
               Estilo          =   4
               Apariencia      =   1
               Borde           =   1
               DecimalesValue  =   2
               MaximoValor     =   1E+18
            End
            Begin TAMControls.TAMTextBox txtInteresAdicional 
               Height          =   285
               Left            =   3180
               TabIndex        =   113
               Top             =   330
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   503
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
               Container       =   "frmCancelacionAcreencias.frx":16C7
               Text            =   "0.00"
               Decimales       =   2
               Estilo          =   4
               Apariencia      =   1
               Borde           =   1
               DecimalesValue  =   2
               MaximoValor     =   1E+18
            End
            Begin TAMControls.TAMTextBox txtPorcenImptoInteresAdicional 
               Height          =   285
               Left            =   2190
               TabIndex        =   123
               Top             =   690
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   503
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
               Container       =   "frmCancelacionAcreencias.frx":16E3
               Text            =   "0.0000"
               Decimales       =   4
               Estilo          =   4
               Apariencia      =   1
               Borde           =   1
               DecimalesValue  =   4
               MaximoValor     =   1E+18
            End
            Begin VB.Label lblDescrip 
               AutoSize        =   -1  'True
               Caption         =   "Interés Adicional:"
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
               Left            =   210
               TabIndex        =   77
               Top             =   360
               Width           =   1485
            End
            Begin VB.Label lblDescrip 
               AutoSize        =   -1  'True
               Caption         =   "IGV Interés Adicional:"
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
               Left            =   210
               TabIndex        =   76
               Top             =   720
               Width           =   1860
            End
            Begin VB.Label lblDescrip 
               AutoSize        =   -1  'True
               Caption         =   "%"
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
               Index           =   33
               Left            =   3000
               TabIndex        =   75
               Top             =   720
               Width           =   150
            End
         End
         Begin TAMControls.TAMTextBox txtInteresAFavor 
            Height          =   285
            Left            =   3270
            TabIndex        =   110
            Top             =   960
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
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
            Container       =   "frmCancelacionAcreencias.frx":16FF
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   2
            MaximoValor     =   1E+18
         End
         Begin TAMControls.TAMTextBox txtImptoInteresAFavor 
            Height          =   285
            Left            =   3270
            TabIndex        =   111
            Top             =   1320
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
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
            Container       =   "frmCancelacionAcreencias.frx":171B
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   2
            MaximoValor     =   1E+18
         End
         Begin TAMControls.TAMTextBox txtPorcenImptoInteresAFavor 
            Height          =   285
            Left            =   2280
            TabIndex        =   124
            Top             =   1320
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
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
            Container       =   "frmCancelacionAcreencias.frx":1737
            Text            =   "0.0000"
            Decimales       =   4
            Estilo          =   4
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   4
            MaximoValor     =   1E+18
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "%"
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
            Index           =   134
            Left            =   3090
            TabIndex        =   73
            Top             =   1350
            Width           =   150
         End
         Begin VB.Label lblDescripMonedaPago 
            Caption         =   "Nuevos Soles (S/.)"
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
            Left            =   3300
            TabIndex        =   72
            Top             =   360
            Width           =   1665
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "IGV de Interés a Favor:"
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
            Left            =   210
            TabIndex        =   71
            Top             =   1350
            Width           =   1995
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Interés a Favor:"
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
            Index           =   16
            Left            =   210
            TabIndex        =   70
            Top             =   1020
            Width           =   1350
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Principal Adeudado:"
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
            Left            =   210
            TabIndex        =   69
            Top             =   660
            Width           =   1725
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
         Left            =   -74850
         Picture         =   "frmCancelacionAcreencias.frx":1753
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   7320
         Width           =   1200
      End
      Begin VB.Frame fraDetallePago 
         Caption         =   "Detalle del Pago"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   -74850
         TabIndex        =   65
         Top             =   480
         Width           =   14985
         Begin TrueOleDBGrid60.TDBGrid tdgDetallePago 
            Bindings        =   "frmCancelacionAcreencias.frx":1BD8
            Height          =   1995
            Left            =   120
            OleObjectBlob   =   "frmCancelacionAcreencias.frx":1BF5
            TabIndex        =   67
            Top             =   270
            Width           =   14745
         End
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
         Left            =   13920
         Picture         =   "frmCancelacionAcreencias.frx":8370
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   7320
         Width           =   1200
      End
      Begin VB.Frame fraDetalle 
         Caption         =   "Detalle del Anexo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   150
         TabIndex        =   54
         Top             =   4230
         Width           =   14985
         Begin VB.CommandButton cmdPagoTotal 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   14610
            Picture         =   "frmCancelacionAcreencias.frx":87F6
            Style           =   1  'Graphical
            TabIndex        =   150
            Top             =   2280
            Width           =   270
         End
         Begin TrueOleDBGrid60.TDBGrid tdgAnexo 
            Bindings        =   "frmCancelacionAcreencias.frx":8B99
            Height          =   1995
            Left            =   120
            OleObjectBlob   =   "frmCancelacionAcreencias.frx":8BB7
            TabIndex        =   59
            Top             =   270
            Width           =   14745
         End
         Begin VB.TextBox txtObservacion 
            Height          =   285
            Left            =   1350
            TabIndex        =   146
            Top             =   2430
            Width           =   7125
         End
         Begin TAMControls.TAMTextBox txtMontoRecibido 
            Height          =   285
            Left            =   13290
            TabIndex        =   128
            Top             =   2250
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   503
            BackColor       =   12632319
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
            Container       =   "frmCancelacionAcreencias.frx":F030
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            ColorEnfoque    =   8454143
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   2
            MaximoValor     =   1E+18
         End
         Begin TAMControls.TAMTextBox txtDeudaRestante 
            Height          =   285
            Left            =   13290
            TabIndex        =   129
            Top             =   2550
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   503
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
            Container       =   "frmCancelacionAcreencias.frx":F04C
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   2
            MaximoValor     =   1E+18
         End
         Begin TrueOleDBGrid60.TDBGrid tdgCuotas 
            Bindings        =   "frmCancelacionAcreencias.frx":F068
            Height          =   1995
            Left            =   120
            OleObjectBlob   =   "frmCancelacionAcreencias.frx":F086
            TabIndex        =   144
            Top             =   270
            Width           =   14745
         End
         Begin VB.Label lblObservacion 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Index           =   57
            Left            =   150
            TabIndex        =   145
            Top             =   2460
            Width           =   1140
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Deuda Restante"
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
            Left            =   11550
            TabIndex        =   60
            Top             =   2610
            Width           =   1395
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto Recibido"
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
            Left            =   11550
            TabIndex        =   55
            Top             =   2310
            Width           =   1350
         End
      End
      Begin VB.Frame fraDatosAnexo 
         Caption         =   "Datos del Anexo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   150
         TabIndex        =   36
         Top             =   2400
         Width           =   14985
         Begin VB.TextBox txtCantidadDocumentosPendientes 
            Height          =   285
            Left            =   5340
            Locked          =   -1  'True
            TabIndex        =   57
            Top             =   690
            Width           =   705
         End
         Begin VB.TextBox txtComisiones 
            Height          =   285
            Left            =   9360
            Locked          =   -1  'True
            TabIndex        =   53
            Top             =   1380
            Width           =   1845
         End
         Begin VB.TextBox txtTotalAnexoDesc 
            Height          =   285
            Left            =   11250
            Locked          =   -1  'True
            TabIndex        =   52
            Top             =   690
            Width           =   1845
         End
         Begin VB.TextBox txtPorcenDesc 
            Height          =   285
            Left            =   9360
            Locked          =   -1  'True
            TabIndex        =   49
            Top             =   1050
            Width           =   1845
         End
         Begin VB.TextBox txtTasaFacial 
            Height          =   285
            Left            =   2490
            Locked          =   -1  'True
            TabIndex        =   48
            Top             =   1410
            Width           =   1515
         End
         Begin VB.TextBox txtTotalAnexo 
            Height          =   285
            Left            =   9360
            Locked          =   -1  'True
            TabIndex        =   47
            Top             =   690
            Width           =   1845
         End
         Begin VB.TextBox txtMoneda 
            Height          =   285
            Left            =   9360
            Locked          =   -1  'True
            TabIndex        =   46
            Top             =   330
            Width           =   3735
         End
         Begin VB.TextBox txtFechaInicio 
            Height          =   285
            Left            =   2490
            Locked          =   -1  'True
            TabIndex        =   45
            Top             =   1050
            Width           =   1515
         End
         Begin VB.TextBox txtCantidadDocumentos 
            Height          =   285
            Left            =   2490
            Locked          =   -1  'True
            TabIndex        =   44
            Top             =   690
            Width           =   705
         End
         Begin VB.TextBox txtNumeroContrato 
            Height          =   285
            Left            =   2490
            Locked          =   -1  'True
            TabIndex        =   43
            Top             =   330
            Width           =   3555
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Pendientes de pago:"
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
            Left            =   3420
            TabIndex        =   56
            Top             =   750
            Width           =   1770
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisiones"
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
            Left            =   7110
            TabIndex        =   51
            Top             =   1440
            Width           =   960
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "% Descuento"
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
            Left            =   7080
            TabIndex        =   50
            Top             =   1080
            Width           =   1125
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   " Tasa Facial"
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
            Left            =   150
            TabIndex        =   42
            Top             =   1440
            Width           =   1065
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Inicio"
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
            Left            =   210
            TabIndex        =   41
            Top             =   1110
            Width           =   1335
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad de Documentos"
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
            TabIndex        =   40
            Top             =   750
            Width           =   2145
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
            Index           =   4
            Left            =   7110
            TabIndex        =   39
            Top             =   360
            Width           =   690
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Total (Sin Dsc | Con Dsc)"
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
            Left            =   7080
            TabIndex        =   38
            Top             =   750
            Width           =   2175
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Número de Contrato"
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
            TabIndex        =   37
            Top             =   390
            Width           =   1695
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
         Height          =   2115
         Left            =   -74850
         TabIndex        =   12
         Top             =   480
         Width           =   14985
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
            Left            =   13170
            Picture         =   "frmCancelacionAcreencias.frx":142E8
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Enviar a BackOffice"
            Top             =   1230
            Width           =   1200
         End
         Begin VB.ComboBox cboEstado 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   1200
            Width           =   4785
         End
         Begin VB.ComboBox cboTipoInstrumento 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   780
            Width           =   4785
         End
         Begin VB.ComboBox cboFondo 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   360
            Width           =   4785
         End
         Begin MSComCtl2.DTPicker dtpFechaOrdenDesde 
            Height          =   315
            Left            =   10140
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
            CheckBox        =   -1  'True
            Format          =   181403649
            CurrentDate     =   38785
         End
         Begin MSComCtl2.DTPicker dtpFechaOrdenHasta 
            Height          =   315
            Left            =   12945
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
            Format          =   181403649
            CurrentDate     =   38785
         End
         Begin MSComCtl2.DTPicker dtpFechaLiquidacionDesde 
            Height          =   315
            Left            =   10140
            TabIndex        =   19
            Top             =   780
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
            Format          =   181403649
            CurrentDate     =   38785
         End
         Begin MSComCtl2.DTPicker dtpFechaLiquidacionHasta 
            Height          =   315
            Left            =   12945
            TabIndex        =   20
            Top             =   780
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
            Format          =   181403649
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
            Left            =   360
            TabIndex        =   29
            Top             =   1335
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
            Left            =   360
            TabIndex        =   28
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
            Left            =   12120
            TabIndex        =   27
            Top             =   405
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
            Left            =   9240
            TabIndex        =   26
            Top             =   435
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
            Left            =   390
            TabIndex        =   25
            Top             =   375
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
            Left            =   7320
            TabIndex        =   24
            Top             =   435
            Width           =   1110
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Liquidación"
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
            Index           =   44
            Left            =   7320
            TabIndex        =   23
            Top             =   825
            Width           =   1560
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
            Index           =   45
            Left            =   9240
            TabIndex        =   22
            Top             =   825
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
            Index           =   46
            Left            =   12120
            TabIndex        =   21
            Top             =   825
            Width           =   510
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
         Height          =   1905
         Left            =   150
         TabIndex        =   1
         Top             =   480
         Width           =   14985
         Begin VB.TextBox txtLinea 
            Height          =   285
            Left            =   9330
            Locked          =   -1  'True
            TabIndex        =   30
            Top             =   1470
            Width           =   4215
         End
         Begin VB.TextBox txtGestor 
            Height          =   285
            Left            =   9330
            Locked          =   -1  'True
            TabIndex        =   35
            Top             =   1110
            Width           =   4215
         End
         Begin VB.TextBox txtEmisor 
            Height          =   285
            Left            =   9330
            Locked          =   -1  'True
            TabIndex        =   34
            Top             =   750
            Width           =   4215
         End
         Begin VB.TextBox txtSubClase 
            Height          =   285
            Left            =   2490
            Locked          =   -1  'True
            TabIndex        =   33
            Top             =   1470
            Width           =   4185
         End
         Begin VB.TextBox txtClase 
            Height          =   285
            Left            =   2490
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   1110
            Width           =   4185
         End
         Begin VB.TextBox txtInstrumento 
            Height          =   285
            Left            =   2490
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   750
            Width           =   4185
         End
         Begin VB.CommandButton cmdBuscarOrig 
            Caption         =   "..."
            Height          =   285
            Left            =   11640
            TabIndex        =   4
            Top             =   345
            Width           =   315
         End
         Begin VB.TextBox txtNumOrig 
            Height          =   285
            Left            =   9330
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   360
            Width           =   2295
         End
         Begin VB.ComboBox cboFondoOrden 
            Height          =   315
            Left            =   2490
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   360
            Width           =   4185
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
            Index           =   13
            Left            =   360
            TabIndex        =   58
            Top             =   1155
            Width           =   480
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Nro. Anexo"
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
            Index           =   179
            Left            =   7140
            TabIndex        =   11
            Top             =   435
            Width           =   960
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Línea"
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
            Left            =   7140
            TabIndex        =   10
            Top             =   1500
            Width           =   495
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
            Left            =   7140
            TabIndex        =   9
            Top             =   1155
            Width           =   570
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
            TabIndex        =   8
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
            TabIndex        =   7
            Top             =   810
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
            Left            =   7140
            TabIndex        =   6
            Top             =   810
            Width           =   570
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
            Left            =   360
            TabIndex        =   5
            Top             =   1500
            Width           =   810
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmCancelacionAcreencias.frx":14843
         Height          =   5445
         Left            =   -74850
         OleObjectBlob   =   "frmCancelacionAcreencias.frx":14861
         TabIndex        =   148
         Top             =   2700
         Width           =   14985
      End
   End
End
Attribute VB_Name = "frmCancelacionAcreencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSQL                       As String
'Recordsets de grillas del formulario
Private adoDetalleAnexo          As ADODB.Recordset
Private adoDetallePago           As ADODB.Recordset 'desconectado
Dim adoListaOrdenes              As ADODB.Recordset

'Variable de busqueda y carga de operaciones.
'Private strNumOperacionAnexo     As String
Private indAnexo                 As Boolean
Private blnValSeleccionado       As Boolean

'Arrays de los combos del formulario
Dim arrFondo()                   As String
Dim arrFondoOrden()              As String
Dim arrEstado()                  As String
'Dim arrEmisor()                  As String
'Dim arrMoneda()                  As String
Dim arrTipoInstrumento()         As String

Dim strCodTipoInstrumento        As String
Dim strCodigosFile               As String

Dim strCodFondo                  As String
Dim strCodEstado                 As String
'Dim strCodNegociacion            As String
Dim strCodEmisor                 As String
Dim strCodMoneda                 As String
'Dim strCodMonedaDocumento        As String
'Dim strIndTipoCambio             As String
'Dim strCodObligado               As String
Dim strCodGestor                 As String
Dim strCodFile                   As String
Dim strCodDetalleFile            As String
Dim strCodSubDetalleFile         As String
Dim strCodTitulo                 As String
Dim strEstado                    As String
'Dim strResponsablePago           As String
Dim strResponsablePagoCancel     As String

Dim datFechaEmision              As Date
Dim datFechaVencimiento          As Date
'Dim datFechaPago                 As Date

Dim strCodAnalitica              As String
'Dim strEstadoOrden               As String
'Dim strCodCobroInteres           As String
'Dim dblTipoCambioOperacion       As Double
Dim dblComisionDesembolso        As Double

Dim strNumAnexo                  As String
Dim strNumSolicitud              As String

Dim strNumOperacion              As String
Dim strNumContrato               As String
Dim intCantDocAnexo              As Integer
Dim intCantDocPendientes         As Integer
Dim dblTasaInteres               As Double
Dim dblMontoTotalAnexo           As Double
Dim dblMontoTotalAnexoDesc       As Double
'Dim dblDeudaTotalAnexo           As Double

Dim dblMontoInteresAFavor               As Double
Dim dblMontoInteresReconocido           As Double
Dim dblMontoInteresAjusteCobroMinimo    As Double

Dim dblMontoInteresMoratorio     As Double
Dim dblMontoImptoInteresMoratorio As Double

Dim dblInteresAjuste             As Double
Dim dblImptoInteresAjuste        As Double
Dim dblInteresAdicional          As Double
Dim dblImptoInteresAdicional     As Double
Dim dblAmortizacionPrincipal     As Double
Dim dblNuevoPrincipal            As Double

Dim dblTotalPrincipalAdeudado    As Double
Dim dblTotalInteresAFavor        As Double
Dim dblTotalIGVInteresAFavor     As Double
Dim dblTotalInteresAdic          As Double
Dim dblTotalImptoInteresAdic     As Double
Dim dblTotalInteresMor           As Double
Dim dblTotalImptoInteresMor      As Double
Dim dblTotalDeuda                As Double
Dim dblMargenDevolucion          As Double
Dim dblTotalInteresAjuste        As Double
Dim dblImptoTotalInteresAjuste   As Double
'Dim dblInteresDevuelto           As Double
Dim dblImptoInteresDevuelto      As Double
'Dim dblTotalInteresDevuelto      As Double
'Dim dblImptoTotalInteresDevuelto As Double

Dim strTipoTasa                  As String
Dim strPeriodoTasa               As String
Dim strPeriodoCapitalizacion     As String
Dim strBaseAnual                 As String

'Dim dblValorNominal              As Double
'Dim dblValorNominalDesc          As Double
'Dim dblDeuda                     As Double

'Dim intDiasAdicionales           As Integer
Dim strCodMonedaComision         As String
'Dim strPersonalizaComision       As String
Dim dblPorcenDescuento           As Double

Dim dblMontoRecibido             As Double
Dim dblMontoPagado               As Double
Dim strCodModalidadCalculoInteres As String

Private Sub InicializarValores()
    
    Dim adoRegistro As ADODB.Recordset
    
    '*** Valores Iniciales ***
    strEstado = Reg_Defecto
    tabPagosRFCP.Tab = 0
    
    dtpFechaOrdenDesde.Value = gdatFechaActual
    dtpFechaOrdenHasta.Value = dtpFechaOrdenDesde.Value
    dtpFechaLiquidacionDesde.Value = Null
    dtpFechaLiquidacionHasta.Value = dtpFechaLiquidacionDesde.Value
    
    strCodTitulo = Valor_Caracter
    strResponsablePagoCancel = Valor_Caracter
    
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
    
    
End Sub

Private Sub CargarListas()
    'Dim adoRecord   As ADODB.Recordset
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
    
End Sub

Public Sub Buscar()

    Dim strFechaOrdenDesde       As String, strFechaOrdenHasta        As String
    Dim strFechaLiquidacionDesde As String, strFechaLiquidacionHasta  As String
    Dim datFechaSiguiente        As Date
    Dim adoAuxiliar              As ADODB.Recordset
    
    Me.MousePointer = vbHourglass
    Set adoListaOrdenes = New ADODB.Recordset
    
    '*** Fecha Vigente, Moneda ***
    adoComm.CommandText = "{ call up_ACSelDatosParametro(23,'" & strCodFondo & "','" & gstrCodAdministradora & "','000') }"
    Set adoAuxiliar = adoComm.Execute
    
    If Not adoAuxiliar.EOF Then
        gdatFechaActual = CVDate(adoAuxiliar("FechaCuota"))
        strCodMoneda = Trim$(adoAuxiliar("CodMoneda"))
        frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
    End If

    adoAuxiliar.Close: Set adoAuxiliar = Nothing
    
    If Not IsNull(dtpFechaOrdenDesde.Value) And Not IsNull(dtpFechaOrdenHasta.Value) Then
        strFechaOrdenDesde = Convertyyyymmdd(dtpFechaOrdenDesde.Value)
        datFechaSiguiente = DateAdd("d", 1, dtpFechaOrdenHasta.Value)
        strFechaOrdenHasta = Convertyyyymmdd(datFechaSiguiente)
    End If
    
    If Not IsNull(dtpFechaLiquidacionDesde.Value) And Not IsNull(dtpFechaLiquidacionHasta.Value) Then
        strFechaLiquidacionDesde = Convertyyyymmdd(dtpFechaLiquidacionDesde.Value)
        datFechaSiguiente = DateAdd("d", 1, dtpFechaLiquidacionHasta.Value)
        strFechaLiquidacionHasta = Convertyyyymmdd(datFechaSiguiente)
    End If
            
    strSQL = "SELECT IOR.NumOrden, IOR.CodTitulo, FechaOrden,FechaLiquidacion,CodTitulo,Nemotecnico,EstadoOrden,IOR.CodFile,CodAnalitica,TipoOrden,IOR.CodMoneda," & _
       "(RTRIM(DescripParametro) + SPACE(1) + DescripOrden) DescripOrden,CantOrden,ValorNominal,IOR.PorcenDsctoValorNominal,MontoTotalMFL1, " & _
       "CodSigno DescripMoneda, IOR.NumAnexo, NumDocumentoFisico,IOR.CodDetalleFile, IOR.CodSubDetalleFile, IOR.CodFondo, IOR.CodGirador, " & _
       "IP1.DescripPersona DesGirador, IOR.CodObligado, IP2.DescripPersona DesObligado, IOR.CodGestor, IP3.DescripPersona DesGestor " & _
       "FROM InversionOrden IOR JOIN AuxiliarParametro AUX ON(AUX.CodParametro=IOR.TipoOrden AND AUX.CodTipoParametro = 'OPECAJ') " & _
       "JOIN Moneda MON ON(MON.CodMoneda=IOR.CodMoneda) " & _
       "LEFT JOIN InstitucionPersona IP1 ON (IP1.CodPersona = IOR.CodGirador AND IP1.TipoPersona = '" & Codigo_Tipo_Persona_Emisor & "') " & _
       "LEFT JOIN InstitucionPersona IP2 ON (IP2.CodPersona = IOR.CodObligado AND IP2.TipoPersona = '" & Codigo_Tipo_Persona_Obligado & "') " & _
       "LEFT JOIN InstitucionPersona IP3 ON (IP3.CodPersona = IOR.CodGestor AND IP3.TipoPersona = '" & Codigo_Tipo_Persona_Emisor & "') " & _
       "WHERE (IOR.TipoOrden = '" & Codigo_Orden_PagoCancelacion & "' OR IOR.TipoOrden = '" & Codigo_Orden_Prepago & "') AND IOR.CodAdministradora='" & gstrCodAdministradora & "' AND IOR.CodFondo='" & strCodFondo & "' "
        
    If strCodTipoInstrumento <> Valor_Caracter Then
        strSQL = strSQL & "AND IOR.CodFile='" & strCodTipoInstrumento & "' "
    Else
        strSQL = strSQL & "AND IOR.CodFile IN " & strCodigosFile & " "
    End If

    If Not IsNull(dtpFechaOrdenDesde.Value) And Not IsNull(dtpFechaOrdenHasta.Value) Then
        strSQL = strSQL & "AND (FechaOrden >='" & strFechaOrdenDesde & "' AND FechaOrden <'" & strFechaOrdenHasta & "') "
    End If
    
    If Not IsNull(dtpFechaLiquidacionDesde.Value) And Not IsNull(dtpFechaLiquidacionHasta.Value) Then
        strSQL = strSQL & "AND (FechaLiquidacion >='" & strFechaLiquidacionDesde & "' AND FechaLiquidacion <'" & strFechaLiquidacionHasta & "') "
    End If
    
    If strCodEstado <> Valor_Caracter Then
        strSQL = strSQL & "AND EstadoOrden='" & strCodEstado & "' "
    End If
    
    strSQL = strSQL & "ORDER BY IOR.NumOrden"
    
    With adoListaOrdenes
        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open strSQL
    End With
    
    tdgConsulta.DataSource = adoListaOrdenes
    tdgConsulta.Refresh

    '  If adoListaOrdenes.Recordset.RecordCount > 0 Then strEstado = Reg_Consulta

    Me.MousePointer = vbDefault
    
End Sub

Public Sub Cancelar()

    Call Buscar

End Sub

Private Sub CargarDesdeCarteraFacturas()
    Dim adoConsulta As ADODB.Recordset
    
    tdgCuotas.Visible = False
    tdgAnexo.Visible = True
    
    With adoComm
        
        .CommandText = "SELECT IO.NumAnexo, IO.CodFile, IO.CodDetalleFile, IO.CodSubDetalleFile, " & _
            "IFL.DescripFile, IDFL.DescripDetalleFile, ISDFL.DescripSubDetalleFile, " & _
            "IO.CodEmisor, IP1.DescripPersona as DescEmisor, IO.CodGestor, IP3.DescripPersona as DescGestor, " & _
            "LRED.DescripLimite as DescripLinea, IO.NumContrato, IO.CantDocumAnexo, IO.TasaInteres, IO.CodMoneda, IO.MontoTotalAnexo,  " & _
            "ROUND((IO.MontoTotalAnexo * IO.PorcenDsctoValorNominal / 100),2) as MontoTotalAnexoDesc, IO.PorcenDsctoValorNominal, IO.MontoAgenteMFL1 as MontoComisionDesembolso, " & _
            "IO.FechaEmision, MON.DescripMoneda   " & _
            "from InversionOperacion IO  " & _
            "join InversionKardex IK on (IK.CodFile = IO.CodFile AND IK.CodAnalitica = IO.CodAnalitica AND IK.SaldoFinal <> 0 AND IK.CodFondo = IO.CodFondo  " & _
            "       AND IK.IndUltimoMovimiento ='X')  " & _
            "join InstitucionPersona IP1 on (IO.CodEmisor = IP1.CodPersona and IP1.TipoPersona = '" & Codigo_Tipo_Persona_Emisor & "') " & _
            "join InstitucionPersona IP2 on (IO.CodObligado = IP2.CodPersona and IP2.TipoPersona = '" & Codigo_Tipo_Persona_Obligado & "')  " & _
            "join InstitucionPersona IP3 on (IO.CodGestor = IP3.CodPersona and IP3.TipoPersona = '" & Codigo_Tipo_Persona_Emisor & "' and IP3.IndBanco = 'X') " & _
            "join InstrumentoInversion II on (IO.CodFile = II.CodFile  and IO.CodAnalitica = II.CodAnalitica and IO.CodFondo = II.CodFondo)  " & _
            "join InversionFile IFL on (IO.CodFile = IFL.CodFile) " & _
            "join InversionDetalleFile IDFL on (IO.CodDetalleFile = IDFL.CodDetalleFile and IO.CodFile = IDFL.CodFile) " & _
            "join InversionSubDetalleFile ISDFL on (IO.CodFile = ISDFL.CodFile and IO.CodDetalleFile = ISDFL.CodDetalleFile and IO.CodSubDetalleFile = ISDFL.CodSubDetalleFile) " & _
            "join Moneda MON on (IO.CodMoneda = MON.CodMoneda) " & _
            "join LimiteReglamentoEstructuraDetalle LRED on (IO.CodLimiteCli = LRED.CodLimite  and IO.CodEstructura = LRED.CodEstructura) " & _
            "where IO.CodFondo = '" & strCodFondo & "' AND IO.CodAdministradora = '" & gstrCodAdministradora & "' AND IO.CodFile in ('" & CodFile_Descuento_Comprobantes_Pago & "','" & CodFile_Descuento_Documentos_Cambiario & "') " & _
            "AND IO.TipoOperacion = '" & Codigo_Orden_Compra & "' and IO.CodEmisor = '" & strCodEmisor & "' "
        
        If indAnexo Then
            .CommandText = .CommandText + " AND IO.NumAnexo = '" & strNumAnexo & "'"
        Else
            .CommandText = .CommandText + " AND IO.NumOperacion = '" & strNumOperacion & "'"
        End If
        
        .CommandText = .CommandText + "" & _
            "Group By IO.NumAnexo, IO.CodFile, IO.CodDetalleFile, IO.CodSubDetalleFile, IFL.DescripFile, IDFL.DescripDetalleFile, " & _
            "ISDFL.DescripSubDetalleFile, IO.CodEmisor, IP1.DescripPersona, IO.CodGestor, IP3.DescripPersona, LRED.DescripLimite, " & _
            "IO.NumContrato, IO.CantDocumAnexo, IO.TasaInteres, IO.CodMoneda, IO.MontoTotalAnexo,IO.MontoTotalAnexo , IO.PorcenDsctoValorNominal, " & _
            "IO.PorcenDsctoValorNominal,IO.FechaEmision , MON.DescripMoneda, MontoAgenteMFL1"

        Set adoConsulta = .Execute
        
    End With
    
    If Not adoConsulta.EOF Then
        '-/** setear valores de controles
        txtNumOrig = strNumAnexo
        txtInstrumento.Text = adoConsulta("DescripFile")
        txtClase.Text = adoConsulta("DescripDetalleFile")
        txtSubClase.Text = adoConsulta("DescripSubDetalleFile")
        txtEmisor.Text = adoConsulta("DescEmisor")
        txtGestor.Text = adoConsulta("DescGestor")
        txtLinea.Text = adoConsulta("DescripLinea")
        txtNumeroContrato.Text = adoConsulta("NumContrato")
        txtCantidadDocumentos.Text = adoConsulta("CantDocumAnexo")
        txtFechaInicio.Text = adoConsulta("FechaEmision")
        txtTasaFacial.Text = adoConsulta("TasaInteres")
        txtMoneda.Text = adoConsulta("DescripMoneda")
        txtTotalAnexo.Text = adoConsulta("MontoTotalAnexo")
        txtTotalAnexoDesc.Text = adoConsulta("MontoTotalAnexoDesc")
        txtPorcenDesc.Text = adoConsulta("PorcenDsctoValorNominal")
        txtComisiones.Text = adoConsulta("MontoComisionDesembolso")
        
        lblDescrip(7).Caption = "Cantidad de Documentos"
        lblDescrip(3).Caption = "Total (Sin Dsc | Con Dsc)"
        fraDatosAnexo.Caption = "Datos del Anexo"
        fraDetalle.Caption = "Detalle del Anexo"
        lblDescrip(179).Caption = "Nro. Anexo"
        lblDescrip(6).Caption = "Emisor"

        '-/** seteo de variables del anexo
        strNumAnexo = adoConsulta("NumAnexo")
        strCodFile = adoConsulta("CodFile")
        strCodDetalleFile = adoConsulta("CodDetalleFile")
        strCodSubDetalleFile = adoConsulta("CodSubDetalleFile")
        strCodEmisor = adoConsulta("CodEmisor")
        strCodGestor = adoConsulta("CodGestor")
        strNumContrato = adoConsulta("NumContrato")
        intCantDocAnexo = adoConsulta("CantDocumAnexo")
        dblTasaInteres = adoConsulta("TasaInteres")
        strCodMoneda = adoConsulta("CodMoneda")
        strCodMonedaComision = adoConsulta("CodMoneda")
        dblMontoTotalAnexo = adoConsulta("MontoTotalAnexo")
        dblMontoTotalAnexoDesc = adoConsulta("MontoTotalAnexoDesc")
        dblPorcenDescuento = adoConsulta("PorcenDsctoValorNominal")
        dblComisionDesembolso = adoConsulta("MontoComisionDesembolso")
        datFechaEmision = adoConsulta("FechaEmision")
    End If
        
    'Obteniendo Datos Adicionales
    
    With adoComm
        .CommandText = "SELECT SUM( IO.ValorNominal) as TotalAnexo, SUM( IO.ValorNominalDscto) as TotalAnexoDesc, count(*) as CantDocumPendientes, SUM(MontoAgenteMFL1) as TotalComision " & _
                        "from InversionOperacion IO    " & _
                        "join InversionKardex IK on (IK.CodFile = IO.CodFile AND IK.CodAnalitica = IO.CodAnalitica AND IK.SaldoFinal <> 0 AND IK.CodFondo = IO.CodFondo AND IK.IndUltimoMovimiento ='X')    " & _
                        "join InstitucionPersona IP1 on (IO.CodObligado = IP1.CodPersona and IP1.TipoPersona = '09')  " & _
                        "where   IO.CodFondo = '" & strCodFondo & "'  AND IO.CodAdministradora = '" & gstrCodAdministradora & _
                        "'   AND IO.CodFile in ('014','015')   AND IO.TipoOperacion = '01'   and IO.CodEmisor = '" & strCodEmisor & "'    " & _
                        "and IO.NumAnexo = '" & strNumAnexo & "' group by IO.CodEmisor, IO.NumAnexo,CantDocumAnexo  order by 1, 2 "
        
        Set adoConsulta = .Execute
        
    End With
    
    If Not adoConsulta.EOF Then
        txtTotalAnexo.Text = adoConsulta("TotalAnexo")
        txtTotalAnexoDesc.Text = adoConsulta("TotalAnexoDesc")
        txtCantidadDocumentosPendientes.Text = adoConsulta("CantDocumPendientes")
        txtComisiones.Text = adoConsulta("TotalComision")
        
        dblMontoTotalAnexo = adoConsulta("TotalAnexo")
        dblMontoTotalAnexoDesc = adoConsulta("TotalAnexoDesc")
        intCantDocPendientes = adoConsulta("CantDocumPendientes")
        dblComisionDesembolso = adoConsulta("TotalComision")
    End If
    
    'Rellenar Grilla de Detalle Anexo
    
    Set adoDetalleAnexo = New ADODB.Recordset
    Me.MousePointer = vbHourglass
    strSQL = "SELECT IO.NumOperacion, IO.CodFile, IO.CodAnalitica, IO.CodTitulo, II.Nemotecnico, " & _
       " IO.CodObligado, IO.CodComisionista, IO.NumSecuencialComisionistaCondicion, IP1.DescripPersona, IO.DescripOperacion, IO.NumDocumentoFisico, IO.FechaVencimiento, " & _
       " IO.ValorNominal, IO.ValorNominalDscto, IO.MontoInteres, IO.PorcenImptoInteres, IO.MontoImptoInteres, IO.PorcenComision, " & _
       " IO.CantDiasPlazo, IO.ResponsablePago, IO.MontoInteresCobroMinimo, IO.DiasCobroMinimo, " & _
       " dbo.uf_ACCalcularDeudaTotal('" & strCodFondo & "','" & gstrCodAdministradora & "',IO.NumOperacion,'" & gstrFechaActual & "') as Deuda, " & _
       " MontoPrincipalAdeudado, InteresAdicAdeudado +(case when IO.FechaVencimiento < '" & gstrFechaActual & "'" & _
       " then case when IOCC.FechaVencimientoCuota < '29990101' then " & _
       " dbo.uf_ACCalcularInteres(IO.TasaInteres,IO.TipoTasa,IO.PeriodoTasa,IO.PeriodoCapitalizacion,IO.BaseAnual,MontoPrincipalAdeudado,IO.FechaVencimiento,'" & gstrFechaActual & "') Else " & _
       " dbo.uf_ACCalcularInteres(IO.TasaInteres,IO.TipoTasa,IO.PeriodoTasa,IO.PeriodoCapitalizacion,IO.BaseAnual,MontoPrincipalAdeudado,IOCC.FechaInicioCuota,'" & gstrFechaActual & "') end" & _
       " else 0 end) as InteresAdicional," & _
       " (case when IO.FechaVencimiento < '" & gstrFechaActual & "'" & _
       " then cast(round((InteresAdicAdeudado + case when IOCC.FechaVencimientoCuota < '29990101' then " & _
       " dbo.uf_ACCalcularInteres(IO.TasaInteres,IO.TipoTasa,IO.PeriodoTasa,IO.PeriodoCapitalizacion,IO.BaseAnual,MontoPrincipalAdeudado,IO.FechaVencimiento,'" & gstrFechaActual & "') Else " & _
       " dbo.uf_ACCalcularInteres(IO.TasaInteres,IO.TipoTasa,IO.PeriodoTasa,IO.PeriodoCapitalizacion,IO.BaseAnual,MontoPrincipalAdeudado,IOCC.FechaInicioCuota,'" & gstrFechaActual & "') " & _
       " end) *(IO.PorcenImptoInteres/100),2)as Decimal(19,2))" & _
       " else 0 end) as IGVInteresAdicional, dbo.uf_ACCalcularInteresMoratorio(IO.CodFondo, IO.CodAdministradora, IO.NumOperacion,'" & gstrFechaActual & "') as InteresMoratorio, " & _
       " cast(round(dbo.uf_ACCalcularInteresMoratorio(IO.CodFondo, IO.CodAdministradora, IO.NumOperacion,'" & gstrFechaActual & "')*(IO.PorcenImptoInteres/100),2)as Decimal(19,2)) as IGVInteresMoratorio, " & _
       " IO.CodMonedaDocumento, IO.ValorNominalDocumento, IO.IndTipoCambio,IO.TipoTasa,IO.BaseAnual,IO.PeriodoTasa,IO.IndCapitalizable,IO.PeriodoCapitalizacion,IO.IndGeneraLetra, IO.ModoCobroInteres" & _
       " from InversionOperacion IO   " & _
       " join InversionKardex IK on (IK.CodFile = IO.CodFile AND IK.CodAnalitica = IO.CodAnalitica AND IK.SaldoFinal <> 0 AND IK.CodFondo = IO.CodFondo AND IK.IndUltimoMovimiento ='X')   " & _
       " join InstitucionPersona IP1 on (IO.CodObligado = IP1.CodPersona and IP1.TipoPersona = '" & Codigo_Tipo_Persona_Obligado & "') " & _
       " join InversionOperacionCalendarioCuota IOCC on (IO.CodFondo = IOCC.CodFondo and IO.NumOperacion = IOCC.NumOperacionOrig and IOCC.EstadoCuotaCalendario = 1 and IO.CodFile = IOCC.CodFile and IO.CodAnalitica = IOCC.CodAnalitica ) " & _
       " join InstrumentoInversion II on (IO.CodFile = II.CodFile and IO.CodAnalitica = II.CodAnalitica and IO.CodFondo = II.CodFondo and IO.CodAdministradora = II.CodAdministradora) " & _
       " where IO.CodFondo = '" & strCodFondo & "' AND IO.CodAdministradora = '" & gstrCodAdministradora & "' AND IO.CodFile in ('" & CodFile_Descuento_Comprobantes_Pago & "','" & CodFile_Descuento_Documentos_Cambiario & "')" & _
       " AND IO.TipoOperacion = '" & Codigo_Orden_Compra & "' and IO.CodEmisor = '" & strCodEmisor & "' and IO.NumAnexo = '" & strNumAnexo & "'"

    If Not indAnexo Then
        strSQL = strSQL & " AND IO.NumOperacion = '" & strNumOperacion & "'"
    End If
    
    strSQL = strSQL & " AND IO.CodTitulo NOT IN (SELECT CodTitulo FROM InversionOrden WHERE CodFondo = '" & strCodFondo & _
                        "' AND CodAdministradora = '" & gstrCodAdministradora & "' AND TipoOrden in ('" & Codigo_Orden_Prepago & "','" & _
                        Codigo_Orden_PagoCancelacion & "') AND EstadoOrden in ('" & Estado_Orden_Enviada & "','" & Estado_Orden_Ingresada & "','" & Estado_Orden_PorAutorizar & "'))"
    strSQL = strSQL & " ORDER BY 1"
                                
    With adoDetalleAnexo
        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open strSQL
    End With
    
    tdgAnexo.DataSource = adoDetalleAnexo
    tdgAnexo.Refresh
    Me.MousePointer = vbDefault
    
    If adoDetalleAnexo.EOF Then
        If indAnexo Then
            MsgBox "El Anexo no tiene operaciones pendientes de pago o éstas no han sido confirmadas.", vbInformation, "Error"
        Else
            MsgBox "La operación tiene una orden de pago que no ha sido confirmada.", vbExclamation, "Error"
        End If
    End If
    
    If Not adoDetalleAnexo.EOF Then
        CalcularTotalesAnexoFacturas
    End If
    
End Sub

Private Sub CalcularTotalesAnexoFacturas()
    Dim dblTotalValorNominal        As Double
    Dim dblTotalValorNominalDscto   As Double
    Dim dblTotalInteres             As Double
    Dim dblTotalInteresAdicional    As Double
    Dim dblTotalInteresMoratorio    As Double
    Dim dblTotalDeudaAnexo          As Double
    Dim dblTotalCobroMinimo         As Double
    
    dblTotalValorNominal = 0
    dblTotalValorNominalDscto = 0
    dblTotalInteres = 0
    dblTotalInteresAdicional = 0
    dblTotalInteresMoratorio = 0
    dblTotalDeudaAnexo = 0
    dblTotalCobroMinimo = 0
    
    adoDetalleAnexo.MoveFirst
    While Not adoDetalleAnexo.EOF
        dblTotalValorNominal = dblTotalValorNominal + adoDetalleAnexo("ValorNominal")
        dblTotalValorNominalDscto = dblTotalValorNominalDscto + adoDetalleAnexo("ValorNominalDscto")
        dblTotalInteres = dblTotalInteres + adoDetalleAnexo("MontoInteres")
        dblTotalInteresAdicional = dblTotalInteresAdicional + adoDetalleAnexo("InteresAdicional")
        dblTotalInteresMoratorio = dblTotalInteresMoratorio + adoDetalleAnexo("InteresMoratorio")
        dblTotalCobroMinimo = dblTotalCobroMinimo + adoDetalleAnexo("MontoInteresCobroMinimo")
        dblTotalDeudaAnexo = dblTotalDeudaAnexo + adoDetalleAnexo("Deuda")
        adoDetalleAnexo.MoveNext
    Wend
    
    tdgAnexo.Columns(5).FooterText = dblTotalValorNominal
    tdgAnexo.Columns(6).FooterText = dblTotalValorNominalDscto
    tdgAnexo.Columns(7).FooterText = dblTotalInteres
    tdgAnexo.Columns(8).FooterText = dblTotalCobroMinimo
    tdgAnexo.Columns(9).FooterText = dblTotalInteresAdicional
    tdgAnexo.Columns(10).FooterText = dblTotalInteresMoratorio
    tdgAnexo.Columns(11).FooterText = dblTotalDeudaAnexo
    
    adoDetalleAnexo.MoveFirst
    
    txtDeudaRestante.Text = dblTotalDeudaAnexo
    
End Sub

Private Sub CalcularTotalesDetallePago()
    Dim dblTotalAdeudado        As Double
    Dim dblTotalAjusteCobroMin  As Double
    Dim dblTotalInteresAdic     As Double
    Dim dblTotalIGVInteresAdic  As Double
    Dim dblTotalDeuda           As Double
    Dim dblTotalMontoPagado     As Double
    
    dblTotalAdeudado = 0
    dblTotalAjusteCobroMin = 0
    dblTotalInteresAdic = 0
    dblTotalIGVInteresAdic = 0
    dblTotalDeuda = 0
    dblTotalMontoPagado = 0
    
    adoDetallePago.MoveFirst
    While Not adoDetallePago.EOF
        dblTotalAdeudado = dblTotalAdeudado + adoDetallePago("PrincipalAdeudado")
        dblTotalAjusteCobroMin = dblTotalAjusteCobroMin + adoDetallePago("AjusteCobroMinimo")
        dblTotalInteresAdic = dblTotalInteresAdic + adoDetallePago("InteresAdicional")
        dblTotalIGVInteresAdic = dblTotalIGVInteresAdic + adoDetallePago("IGVInteresAdicional")
        dblTotalDeuda = dblTotalDeuda + adoDetallePago("Deuda")
        dblTotalMontoPagado = dblTotalMontoPagado + adoDetallePago("MontoPago")
        adoDetallePago.MoveNext
    Wend
    
    tdgDetallePago.Columns(3).FooterText = dblTotalAdeudado
    tdgDetallePago.Columns(4).FooterText = dblTotalAjusteCobroMin
    tdgDetallePago.Columns(5).FooterText = dblTotalInteresAdic
    tdgDetallePago.Columns(6).FooterText = dblTotalIGVInteresAdic
    tdgDetallePago.Columns(7).FooterText = dblTotalInteresMor
    tdgDetallePago.Columns(8).FooterText = dblTotalImptoInteresMor
    tdgDetallePago.Columns(9).FooterText = dblTotalDeuda
    tdgDetallePago.Columns(10).FooterText = dblTotalMontoPagado
    
    adoDetallePago.MoveFirst
    
End Sub

Private Sub CargarDesdeCarteraFlujos()
    Dim adoConsulta As ADODB.Recordset
    
    tdgAnexo.Visible = False
    tdgCuotas.Visible = True
    
    With adoComm
        .CommandText = "select ISL.CodFondo, NumSolicitud,IO.CodFile,IO.CodDetalleFile,IO.CodSubDetalleFile , IFL.DescripFile,IDFL.DescripDetalleFile,ISDFL.DescripSubDetalleFile, " & _
           "IO.CodEmisor,IP1.DescripPersona as DescEmisor, IO.CodGestor, IP3.DescripPersona as DescGestor, 'STUB' as DescripLinea, IO.NumContrato, " & _
           "(select COUNT(*) from InversionSolicitudCalendario where NumSolicitud = ISL.NumSolicitud and CodFondo = ISL.CodFondo) as CantCuotas, " & _
           "(select COUNT(*) from InversionSolicitudCalendario where NumSolicitud = ISL.NumSolicitud and CodFondo = ISL.CodFondo and EstadoCupon = 'P') as CantCuotasPendientes, " & _
           "ISL.TasaInteres,ISL.CodMoneda, ISL.MontoAprobado, ISL.MontoSolicitud, IO.MontoAgenteMFL1 as MontoComisionDesembolso, ISL.FechaEmision, MON.DescripMoneda " & _
           "from InversionOperacion IO " & _
           "join InversionSolicitud ISL on (IO.CodFondo = ISL.CodFondo and IO.CodFile = ISL.CodFile and IO.CodAnalitica = ISL.CodAnalitica and IO.TipoOperacion ='01') " & _
           "join InversionKardex IK on (ISL.CodFondo = IK.CodFondo and ISL.CodAnalitica = IK.CodAnalitica and ISL.CodFile = IK.CodFile and IK.SaldoFinal > 0 and IK.IndUltimoMovimiento = 'X') " & _
           "join InstitucionPersona IP1 on (IO.CodEmisor = IP1.CodPersona and IP1.TipoPersona = '02') " & _
           "join InstitucionPersona IP3 on (IO.CodGestor = IP3.CodPersona and IP3.TipoPersona = '02' and IP3.IndBanco = 'X') " & _
           "join InversionFile IFL on (IO.CodFile = IFL.CodFile) " & _
           "join InversionDetalleFile IDFL on (IO.CodDetalleFile = IDFL.CodDetalleFile and IO.CodFile = IDFL.CodFile) " & _
           "join InversionSubDetalleFile ISDFL on (IO.CodFile = ISDFL.CodFile and IO.CodDetalleFile = ISDFL.CodDetalleFile and IO.CodSubDetalleFile = ISDFL.CodSubDetalleFile) " & _
           "join Moneda MON on (IO.CodMoneda = MON.CodMoneda) " & _
           "Where NumSolicitud = " & strNumSolicitud & " And IO.CodFondo = '" & strCodFondo & "' AND IO.CodAdministradora = '" & gstrCodAdministradora & "' AND IO.CodFile = '016' "

        Set adoConsulta = .Execute
        
    End With
    
    If Not adoConsulta.EOF Then
        '-/** setear valores de controles
        txtNumOrig = adoConsulta("NumSolicitud")
        txtInstrumento.Text = adoConsulta("DescripFile")
        txtClase.Text = adoConsulta("DescripDetalleFile")
        txtSubClase.Text = adoConsulta("DescripSubDetalleFile")
        txtEmisor.Text = adoConsulta("DescEmisor")
        txtGestor.Text = adoConsulta("DescGestor")
        txtLinea.Text = adoConsulta("DescripLinea")
        txtNumeroContrato.Text = adoConsulta("NumContrato")
        txtCantidadDocumentos.Text = adoConsulta("CantCuotas")
        txtCantidadDocumentosPendientes.Text = adoConsulta("CantCuotasPendientes")
        txtFechaInicio.Text = adoConsulta("FechaEmision")
        txtTasaFacial.Text = adoConsulta("TasaInteres")
        txtMoneda.Text = adoConsulta("DescripMoneda")
        txtTotalAnexo.Text = adoConsulta("MontoSolicitud")
        txtTotalAnexoDesc.Text = adoConsulta("MontoAprobado")
        txtPorcenDesc.Text = "N/A"
        txtComisiones.Text = adoConsulta("MontoComisionDesembolso")
        
        lblDescrip(7).Caption = "Cantidad de Cuotas"
        lblDescrip(3).Caption = "Solicitado / Aprobado"
        lblDescrip(179).Caption = "Nro. Solicitud"
        lblDescrip(6).Caption = "Contratante"
        fraDatosAnexo.Caption = "Datos del Flujo"
        fraDetalle.Caption = "Detalle de Cuotas Pendientes"

        strCodFile = adoConsulta("CodFile")
        strCodDetalleFile = adoConsulta("CodDetalleFile")
        strCodSubDetalleFile = adoConsulta("CodSubDetalleFile")
        strCodEmisor = adoConsulta("CodEmisor")
        strCodGestor = adoConsulta("CodGestor")
        strNumContrato = adoConsulta("NumContrato")
        intCantDocAnexo = adoConsulta("CantCuotas")
        dblTasaInteres = adoConsulta("TasaInteres")
        strCodMoneda = adoConsulta("CodMoneda")
        strCodMonedaComision = adoConsulta("CodMoneda")
        dblMontoTotalAnexo = adoConsulta("MontoSolicitud")
        dblMontoTotalAnexoDesc = adoConsulta("MontoAprobado")
        dblPorcenDescuento = 100
        dblComisionDesembolso = adoConsulta("MontoComisionDesembolso")
        datFechaEmision = adoConsulta("FechaEmision")
    End If
    
    'Rellenar Grilla de Detalle de cuotas
    
    Set adoDetalleAnexo = New ADODB.Recordset
    Me.MousePointer = vbHourglass
    strSQL = "select NumCupon as NumCuota,NumSecuencial,FechaVencimiento as FechaPago,Principal,Intereses as Interes,IGVIntereses as ImptoInteres,InteresAdicional,IGVInteresAdicional as ImptoInteresAdicional, " & "Principal+Intereses+IGVIntereses+InteresAdicional+IGVInteresAdicional as Deuda " & "from InversionSolicitudCalendario where CodFondo = '" & strCodFondo & "' and NumSolicitud = " & strNumSolicitud & " and EstadoCupon = 'P' order by NumCupon,NumSecuencial"
                                
    With adoDetalleAnexo
        .ActiveConnection = gstrConnectConsulta
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open strSQL
    End With
    
    tdgCuotas.DataSource = adoDetalleAnexo
    tdgCuotas.Refresh
    Me.MousePointer = vbDefault
    
End Sub

Private Sub cboEstado_Click()
    strCodEstado = Valor_Caracter

    If cboEstado.ListIndex < 0 Then Exit Sub
    
    strCodEstado = Trim$(arrEstado(cboEstado.ListIndex))
    
    Call Buscar

End Sub

Private Sub cboFondo_Click()
    On Error GoTo cboFondo_Click_Err
    
    Dim adoRegistro As ADODB.Recordset
    
    strCodFondo = Valor_Caracter
    cboFondoOrden.ListIndex = cboFondo.ListIndex

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
    strSQL = "SELECT FIF.CodFile CODIGO,DescripFile DESCRIP " & "FROM FondoInversionFile FIF JOIN InversionFile IVF ON(IVF.CodFile=FIF.CodFile) " & "WHERE TipoValor='" & Codigo_Valor_RentaFija & "' AND TipoPlazo='" & Codigo_Valor_CortoPlazo & "' AND IndInstrumento='X' AND IndVigente='X' AND " & "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND FIF.CodFile <> '" & CodFile_Descuento_Flujos_Dinerarios & "' ORDER BY DescripFile"
    CargarControlLista strSQL, cboTipoInstrumento, arrTipoInstrumento(), Sel_Todos
    
    If cboTipoInstrumento.ListCount > 0 Then cboTipoInstrumento.ListIndex = 0
    
    Exit Sub

cboFondo_Click_Err:
    MsgBox err.Description & vbCrLf & "in Inversion.frmCancelacionAcreencias.cboFondo_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"
    Resume Next
End Sub

Private Sub cboTipoInstrumento_Click()
    strCodTipoInstrumento = Valor_Caracter

    If cboTipoInstrumento.ListIndex < 0 Then Exit Sub
    
    strCodTipoInstrumento = Trim$(arrTipoInstrumento(cboTipoInstrumento.ListIndex))
       
    Call Buscar

End Sub

Private Sub cmdAnterior_Click()

    With tabPagosRFCP
        .TabEnabled(0) = False
        .TabEnabled(1) = True
        .TabEnabled(2) = False
        .Tab = 1
    End With
    
End Sub

Private Sub cmdAnular_Click()

    Dim strNumOrden  As String
    Dim strCodTitulo As String
    Dim intRegistro As Integer

    Dim strMensaje As String
    
    For intRegistro = 0 To tdgConsulta.SelBookmarks.Count - 1
        'verificar si la orden no está ya anulada
        adoListaOrdenes.MoveFirst
        adoListaOrdenes.Move CLng(tdgConsulta.SelBookmarks(intRegistro) - 1), 0
        tdgConsulta.Refresh
        
        strNumOrden = Trim$(adoListaOrdenes("NumOrden"))
        strCodTitulo = Trim$(adoListaOrdenes("CodTitulo"))
        strCodEstado = Trim$(adoListaOrdenes("EstadoOrden"))
        
        If strCodEstado <> Estado_Orden_Anulada And strCodEstado <> Estado_Orden_Procesada Then
        
            strMensaje = "Se procederá a eliminar la ORDEN " & strNumOrden & " por la " & tdgConsulta.Columns(3) & vbNewLine & vbNewLine & "¿ Seguro de continuar ?"
            
            If MsgBox(strMensaje, vbQuestion + vbYesNo + vbDefaultButton2, "Anular Orden") = vbYes Then
        
                '*** Anular Orden ***
                adoComm.CommandText = "UPDATE InversionOrden SET EstadoOrden='" & Estado_Orden_Anulada & "' WHERE CodFondo='" & strCodFondo & _
                                        "' AND CodAdministradora='" & gstrCodAdministradora & "' AND CodTitulo='" & strCodTitulo & _
                                        "' AND NumOrden='" & strNumOrden & "'"
                adoConn.Execute adoComm.CommandText
                
                MsgBox "Se anuló la orden correctamente.", vbInformation, "Anular Orden"
              
            End If
            
        Else
            If strCodEstado = Estado_Orden_Anulada Then
                MsgBox "La orden " & strNumOrden & " ya ha sido anulada.", vbExclamation, "Anular Orden"
            Else
                MsgBox "La orden " & strNumOrden & " ya ha sido procesada." & vbNewLine & "No se puede anular.", vbCritical, "Anular Orden"
            End If
        End If
    Next
    
    Call Buscar
    
End Sub

Private Sub cmdBuscar_Click()
    Call Buscar
End Sub

Private Sub cmdBuscarOrig_Click()
    Dim frmBuscarOp As frmBuscarOperacionAnexo
    Set frmBuscarOp = New frmBuscarOperacionAnexo
    
    frmBuscarOp.strCodFondo = strCodFondo
    frmBuscarOp.Show 1
    
    With frmBuscarOp
        blnValSeleccionado = .blnValSeleccionado

        If blnValSeleccionado Then
            strNumAnexo = .numAnexo
            strNumOperacion = .numOperacion
            indAnexo = .indAnexo
            strCodEmisor = .strCodEmisor
        End If

    End With
    
    Set frmBuscarOp = Nothing
    
    If blnValSeleccionado Then
        Call CargarDesdeCarteraFacturas
    End If
    
End Sub

Private Sub cmdCancelar_Click()
    
    With tabPagosRFCP
        .TabEnabled(0) = True
        .TabEnabled(1) = False
        .TabEnabled(2) = False
        .Tab = 0
    End With
    
    cmdCancelar.Visible = False
    cmdNuevo.Visible = True
    cmdAnular.Visible = True
    cmdBuscar.Visible = True

End Sub

Private Sub cmdEnviar_Click()
    Dim strFechaDesde As String
    Dim intRegistro   As Integer, intContador         As Integer
    
    If adoListaOrdenes.RecordCount = 0 Then Exit Sub
    
    strFechaDesde = Convertyyyymmdd(dtpFechaOrdenDesde.Value)
    
    intContador = tdgConsulta.SelBookmarks.Count - 1
    
    If intContador < 0 Then
        MsgBox "No se ha seleccionado ningún registro", vbCritical, Me.Caption
        Exit Sub
    End If
        
    For intRegistro = 0 To intContador
        tdgConsulta.Row = tdgConsulta.SelBookmarks(intRegistro) - 1
               
        If strCodEstado = Estado_Orden_Ingresada Then
            adoComm.CommandText = "UPDATE InversionOrden SET EstadoOrden='" & Estado_Orden_Enviada & "'," & "UsuarioEdicion='" & gstrLogin & "',FechaEdicion='" & strFechaDesde & Space$(1) & Format$(Time, "hh:mm") & "' " & "WHERE NumOrden='" & Trim$(tdgConsulta.Columns(0)) & "' AND CodFondo='" & strCodFondo & "' AND " & "CodAdministradora='" & gstrCodAdministradora & "' AND EstadoOrden='" & Estado_Orden_Ingresada & "'"
        ElseIf strCodEstado = Estado_Orden_Enviada Then
            adoComm.CommandText = "UPDATE InversionOrden SET EstadoOrden='" & Estado_Orden_Ingresada & "'," & "UsuarioEdicion='" & gstrLogin & "',FechaEdicion='" & strFechaDesde & Space$(1) & Format$(Time, "hh:mm") & "' " & "WHERE NumOrden='" & Trim$(tdgConsulta.Columns(0)) & "' AND CodFondo='" & strCodFondo & "' AND " & "CodAdministradora='" & gstrCodAdministradora & "' AND EstadoOrden='" & Estado_Orden_Enviada & "'"
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
    Call Grabar
    cmdCancelar.Visible = False
    cmdNuevo.Visible = True
    cmdAnular.Visible = True
    cmdBuscar.Visible = True
End Sub

Private Sub Grabar()
    Dim cantOrdenesReg As Integer
    'Dim strMsgError As String
    Dim cantOrden As Double
    Dim dblInteresAjusteGrabado As Double  'InteresAjuste del pago parcial anticipado
    Dim dblImptoInteresAjusteGrabado As Double
    
    Dim dblAjusteInteresPagoAdelantado As Double
    Dim dblImptoAjusteInteresPagoAdelantado As Double
    Dim dblAjusteCobroMinimo As Double
    Dim dblImptoAjusteCobroMinimo As Double
    
    Dim dblMontoPagadoOrden As Double
    
    Me.MousePointer = vbHourglass
    
    adoDetallePago.MoveFirst
    adoDetalleAnexo.MoveFirst
    cantOrdenesReg = 0
    
    While Not adoDetallePago.EOF
        
        If adoDetallePago("TipoOperacion") = Codigo_Orden_PagoCancelacion Then
            cantOrden = adoDetallePago("PrincipalAdeudado")
            dblInteresAjusteGrabado = 0
        ElseIf adoDetallePago("TipoOperacion") = Codigo_Orden_Prepago Then
            cantOrden = txtAmortizacionPrincipal.Value
            dblInteresAjusteGrabado = CDec(txtInteresAjuste.Value)
        End If
        dblImptoInteresAjusteGrabado = Round(dblInteresAjusteGrabado * gdblTasaIgv, 2)
        
        dblAjusteInteresPagoAdelantado = adoDetallePago("AjusteInteres")
        dblImptoAjusteInteresPagoAdelantado = Round(adoDetallePago("AjusteInteres") * gdblTasaIgv, 2)
        
        dblAjusteCobroMinimo = adoDetallePago("AjusteCobroMinimo")
        dblImptoAjusteCobroMinimo = Round(dblAjusteCobroMinimo * gdblTasaIgv, 2)
            
        If cantOrdenesReg = adoDetallePago.RecordCount - 1 Then
            dblMontoPagadoOrden = adoDetallePago("MontoPago") + txtMargenDevolucion.Value
        Else
            dblMontoPagadoOrden = adoDetallePago("MontoPago")
        End If
        
        '*** Guardar Orden de pago por cada pago ***
        With adoComm
            'txtMontoRecibido.Value
            .CommandText = "{ call up_IVAdicInversionOrden('" & strCodFondo & "','" & gstrCodAdministradora & "','','" & gstrFechaActual & _
                "','" & adoDetalleAnexo("CodTitulo") & "','" & adoDetalleAnexo("Nemotecnico") & "','" & gstrPeriodoActual & _
               "','" & gstrMesActual & "','','03','" & strCodAnalitica & "','" & strCodFile & "','" & adoDetalleAnexo("CodAnalitica") & _
               "','" & strCodDetalleFile & "','" & strCodSubDetalleFile & _
               "','" & adoDetallePago("TipoOperacion") & "','01','','01','" & adoDetalleAnexo("DescripOperacion") & "','" & strCodEmisor & "','','','" & adoDetalleAnexo("CodComisionista") & "'," & adoDetalleAnexo("NumSecuencialComisionistaCondicion") & ",'" & gstrFechaActual & _
               "','" & Convertyyyymmdd(adoDetallePago("FechaVencimiento")) & "','" & Convertyyyymmdd(Valor_Fecha) & _
               "','" & Convertyyyymmdd(CDate(txtFechaInicio.Text)) & "','" & adoDetalleAnexo("CodMonedaDocumento") & "'," & adoDetalleAnexo("ValorNominalDocumento") & _
               ",'" & Trim$(adoDetalleAnexo("IndTipoCambio")) & "','" & strCodMoneda & _
               "','" & strCodMoneda & "'," & cantOrden & "," & gdblTipoCambio & "," & gdblTipoCambio & "," & adoDetalleAnexo("ValorNominal") & "," & dblPorcenDescuento & _
               "," & adoDetalleAnexo("ValorNominalDscto") & ",0,0,0," & _
               adoDetalleAnexo("MontoInteres") & ",0,0,0,0,0,0,0,0," & adoDetalleAnexo("MontoImptoInteres") & "," & dblMontoPagadoOrden & _
               ",0,0,0," & dblInteresAjusteGrabado & ",0,0,0,0,0,0,1,0," & dblImptoInteresAjusteGrabado & ",0," & adoDetallePago("PrincipalAdeudado") & "," & adoDetalleAnexo("CantDiasPlazo") & _
               ",'X','','','','','','" & strCodEmisor & "','" & adoDetalleAnexo("CodObligado") & "','" & adoDetalleAnexo("CodObligado") & "','" & _
               strCodGestor & "','',0,'','X','X','" & adoDetalleAnexo("TipoTasa") & _
               "','" & adoDetalleAnexo("BaseAnual") & "'," & CDec(dblTasaInteres) & ",'" & adoDetalleAnexo("PeriodoTasa") & "','" & _
               adoDetalleAnexo("IndCapitalizable") & "','" & adoDetalleAnexo("PeriodoCapitalizacion") & _
               "','" & adoDetalleAnexo("IndGeneraLetra") & "'," & CDec(dblTasaInteres) & "," & CDec(dblTasaInteres) & "," & CDec(dblTasaInteres) & _
               ",'','','" & Trim$(txtObservacion.Text) & "','" & gstrLogin & "','" & gstrFechaActual & _
               "','" & gstrLogin & "','" & gstrFechaActual & "','" & adoDetalleAnexo("CodTitulo") & "','" & adoDetalleAnexo("ModoCobroInteres") & _
               "'," & adoDetalleAnexo("MontoInteres") & "," & adoDetalleAnexo("MontoInteresCobroMinimo") & "," & adoDetallePago("InteresAdicional") & "," & dblMontoInteresMoratorio & _
               "," & adoDetalleAnexo("DiasCobroMinimo") & "," & adoDetallePago("CantDiasMora") & ",'01'," & adoDetallePago("PorcenIGV") & "," & adoDetalleAnexo("MontoImptoInteres") & "," & _
               adoDetallePago("IGVInteresAdicional") & "," & dblMontoImptoInteresMoratorio & "," & adoDetallePago("PorcenIGV") & _
               ",0," & dblInteresAjusteGrabado & "," & dblImptoInteresAjusteGrabado & "," & dblAjusteInteresPagoAdelantado & "," & _
               dblImptoAjusteInteresPagoAdelantado & "," & dblAjusteCobroMinimo & "," & dblImptoAjusteCobroMinimo & ",'" & strNumAnexo & "','" & Trim$(strNumContrato) & "','" & Trim$(adoDetalleAnexo("NumDocumentoFisico")) & "','','42','09','" & strCodEmisor & "','" & Codigo_Tipo_Persona_Emisor & "','" & adoDetalleAnexo("ResponsablePago") & _
               "',''," & dblMontoTotalAnexo & "," & intCantDocAnexo & "," & adoDetalleAnexo("PorcenComision") & "," & adoDetallePago("Deuda") & ")}"

            adoConn.Execute .CommandText
            
        End With
        
        Me.MousePointer = vbDefault
        
        
        frmMainMdi.stbMdi.Panels(3).Text = "Acción"
        
        '        cmdOpcion.Visible = True
        With tabPagosRFCP
            .TabEnabled(0) = True
            .Tab = 0
        End With
        
        adoDetallePago.MoveNext
        adoDetalleAnexo.MoveNext
        cantOrdenesReg = cantOrdenesReg + 1
    Wend
    
    MsgBox "Se registraron " & cantOrdenesReg & " órdenes de pago.", vbInformation
    
    Call Buscar

End Sub

Private Sub cmdNuevo_Click()

    If Not EsDiaUtil(gdatFechaActual) Then
        MsgBox "No se puede negociar en un día no útil!", vbCritical, Me.Caption
        Exit Sub
    End If
  
    frmMainMdi.stbMdi.Panels(3).Text = "Adicionar Orden..."
                
    strEstado = Reg_Adicion
      
    cmdNuevo.Visible = False
    cmdAnular.Visible = False
    cmdBuscar.Visible = False
    
    With tabPagosRFCP
        .TabEnabled(0) = False
        .TabEnabled(1) = True
        .TabEnabled(2) = False
        .Tab = 1
    End With
    
    cmdCancelar.Visible = True
    
    Call ClearDatosAnexo
End Sub

Private Sub cmdPagoTotal_Click()
    txtMontoRecibido.Text = tdgAnexo.Columns(11).FooterText
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub GenerarDetallePagoFactura()
    
    Dim strNaturalezaPago As String
    Dim strTipoOperacion  As String
    Dim adoConsulta       As ADODB.Recordset
    Dim strUltimaOrden    As String
    Dim CantDiasMora      As Integer
        
    'se generan los datos del recordset desconectado
    Set adoDetallePago = New ADODB.Recordset

    With adoDetallePago.Fields
        .Append "NumOperacion", adVarChar, 10
        .Append "DescripPersona", adVarChar, 200
        .Append "PorcenIGV", adDecimal, 19
        .Append "PrincipalAdeudado", adDecimal, 19
        .Append "MontoInteresCobroMinimo", adDecimal, 19
        .Append "AjusteCobroMinimo", adDecimal, 19  'ajuste por cobro de intereses minimos
        .Append "InteresAdicional", adDecimal, 19
        .Append "IGVInteresAdicional", adDecimal, 19
        .Append "InteresMoratorio", adDecimal, 19
        .Append "IGVInteresMoratorio", adDecimal, 19
        .Append "Deuda", adDecimal, 19
        .Append "MontoPago", adDecimal, 19
        .Append "TipoPago", adVarChar, 30
        .Append "TipoOperacion", adVarChar, 2
        .Append "FechaVencimiento", adDate
        .Append "AjusteInteres", adDecimal, 19 'si es pago total anticipado
        .Append "CantDiasMora", adInteger
        
        .Item("PorcenIGV").Precision = 19
        .Item("PorcenIGV").NumericScale = 6
        .Item("PrincipalAdeudado").Precision = 19
        .Item("PrincipalAdeudado").NumericScale = 2
        .Item("MontoInteresCobroMinimo").Precision = 19
        .Item("MontoInteresCobroMinimo").NumericScale = 2
        .Item("AjusteCobroMinimo").Precision = 19
        .Item("AjusteCobroMinimo").NumericScale = 2
        .Item("InteresAdicional").Precision = 19
        .Item("InteresAdicional").NumericScale = 2
        .Item("IGVInteresAdicional").Precision = 19
        .Item("IGVInteresAdicional").NumericScale = 2
        .Item("InteresMoratorio").Precision = 19
        .Item("InteresMoratorio").NumericScale = 2
        .Item("IGVInteresMoratorio").Precision = 19
        .Item("IGVInteresMoratorio").NumericScale = 2
        .Item("Deuda").Precision = 19
        .Item("Deuda").NumericScale = 2
        .Item("MontoPago").Precision = 19
        .Item("MontoPago").NumericScale = 2
        .Item("AjusteInteres").Precision = 19
        .Item("AjusteInteres").NumericScale = 2

    End With
    
    adoDetalleAnexo.MoveFirst
    adoDetallePago.Open
    dblMontoRecibido = CDec(txtMontoRecibido.Text)
    
    While (Not adoDetalleAnexo.EOF And dblMontoRecibido > 0)

        If dblMontoRecibido >= adoDetalleAnexo("Deuda") Then
            dblMontoPagado = adoDetalleAnexo("Deuda")
            strNaturalezaPago = "Total"
            strTipoOperacion = Codigo_Orden_PagoCancelacion
        Else
            dblMontoPagado = dblMontoRecibido
            strNaturalezaPago = "Parcial"
            strTipoOperacion = Codigo_Orden_Prepago
        End If
        
        dblMontoRecibido = dblMontoRecibido - dblMontoPagado
        CantDiasMora = DateDiff("d", adoDetalleAnexo("FechaVencimiento"), Convertddmmyyyy(gstrFechaActual))

        If CantDiasMora < 0 Then
            CantDiasMora = 0
        End If
                
        adoDetallePago.AddNew Array("NumOperacion", "DescripPersona", "PorcenIGV", "PrincipalAdeudado", "MontoInteresCobroMinimo", "AjusteCobroMinimo", "InteresAdicional", "IGVInteresAdicional", "InteresMoratorio", "IGVInteresMoratorio", "Deuda", "MontoPago", "TipoPago", "TipoOperacion", "FechaVencimiento", "AjusteInteres", "CantDiasMora"), Array(adoDetalleAnexo("NumOperacion"), adoDetalleAnexo("DescripPersona"), adoDetalleAnexo("PorcenImptoInteres"), adoDetalleAnexo("MontoPrincipalAdeudado"), adoDetalleAnexo("MontoInteresCobroMinimo"), 0, adoDetalleAnexo("InteresAdicional"), adoDetalleAnexo("IGVInteresAdicional"), 0, 0, adoDetalleAnexo("Deuda"), dblMontoPagado, strNaturalezaPago, strTipoOperacion, adoDetalleAnexo("FechaVencimiento"), 0, CantDiasMora)
        
        adoDetalleAnexo.MoveNext
    Wend
    
    'Si hay pago parcial, rellenar los datos
    adoDetallePago.MoveLast
    
    If Not adoDetallePago.EOF Then
        'Consultar el Interes Moratorio provisionado
         With adoComm
            .CommandText = "select dbo.uf_ACCalcularInteresMoratorio('" & gstrCodFondoContable & "','" & gstrCodAdministradora & "','" & adoDetallePago("NumOperacion") & "','" & gstrFechaActual & "') as InteresMoratorio "
            Set adoConsulta = .Execute
        End With
    
        If adoConsulta.EOF Then
            MsgBox ("Data de Valorizacion diaria Inconsistente!")
            Exit Sub
        End If
        
        dblMontoInteresMoratorio = adoConsulta("InteresMoratorio")
        dblMontoImptoInteresMoratorio = 0 ' Round(dblMontoInteresMoratorio * gdblTasaIgv, 2)
        adoDetallePago("InteresMoratorio") = dblMontoInteresMoratorio
        adoDetallePago("IGVInteresMoratorio") = dblMontoImptoInteresMoratorio
        
        ' Seleccionar modalidad de calculo de interes
         With adoComm
            .CommandText = "select dbo.uf_IVObtenerModalidadCalculoInteres('" & gstrCodFondoContable & "','" & gstrCodAdministradora & "') as ValorParametro"
            Set adoConsulta = .Execute
        End With
        If adoConsulta.EOF Then
            strCodModalidadCalculoInteres = "01"
        Else
            strCodModalidadCalculoInteres = Trim$(adoConsulta("ValorParametro"))
        End If
        
        With adoComm
            .CommandText = "select FechaVencimiento, TasaInteres,TipoTasa, PeriodoTasa, PeriodoCapitalizacion, BaseAnual from InversionOperacion " & " Where CodFondo = '" & gstrCodFondoContable & "' AND NumOperacion = '" & adoDetallePago("NumOperacion") & "'"
            Set adoConsulta = .Execute
        End With

        If adoConsulta.EOF Then
            MsgBox ("Data operativa Inconsistente!")
            Exit Sub
        End If
        
        datFechaVencimiento = adoConsulta("FechaVencimiento")
        strTipoTasa = adoConsulta("TipoTasa")
        strPeriodoTasa = adoConsulta("PeriodoTasa")
        strPeriodoCapitalizacion = adoConsulta("PeriodoCapitalizacion")
        strBaseAnual = adoConsulta("BaseAnual")
        
        If adoDetallePago("TipoOperacion") = Codigo_Orden_Prepago Then
   
            If Convertddmmyyyy(gstrFechaActual) < datFechaVencimiento Then

                With adoComm
                    .CommandText = "select ValorInteresDiferido as InteresAFavor from InversionValorizacionDiaria IVD join InversionOperacion IO on " & _
                                    "(IVD.CodFondo = IO.CodFondo and IVD.CodAdministradora = IO.CodAdministradora and IO.CodFile = IVD.CodFile " & _
                                    "and IO.CodAnalitica = IVD.CodAnalitica and IO.CodTitulo = IVD.CodTitulo) where IO.NumOperacion = '" & adoDetallePago("NumOperacion") & "' " & _
                                    "and IO.CodFondo = '" & gstrCodFondoContable & "' and FechaValorizacion = '" & gstrFechaActual & "'"
                    Set adoConsulta = .Execute
                End With

                If adoConsulta.EOF And Convertddmmyyyy(gstrFechaActual) < datFechaVencimiento Then
                    MsgBox ("Data de Valorizacion diaria Inconsistente!")
                    Exit Sub
                End If
                
                dblMontoInteresAFavor = adoConsulta("InteresAFavor") ' * (-1)
                
                If strCodModalidadCalculoInteres = "01" Then
                     With adoComm
                        .CommandText = "select dbo.uf_ACCalcularInteres(" & dblTasaInteres & ", '" & strTipoTasa & "', '" & strPeriodoTasa & "', '" & strPeriodoCapitalizacion & "', '" & strBaseAnual & "'," & adoDetallePago("PrincipalAdeudado") - (dblMontoPagado + dblMontoInteresAFavor) & ", '" & gstrFechaActual & "','" & Convertyyyymmdd(datFechaVencimiento) & "') as InteresAjuste"
                        Set adoConsulta = .Execute
                    End With
    
                    If adoConsulta.EOF Then
                        MsgBox ("Data Inconsistente!")
                        Exit Sub
                    End If
                    dblInteresAjuste = dblMontoInteresAFavor - adoConsulta("InteresAjuste")

                ElseIf strCodModalidadCalculoInteres = "02" Then
                    With adoComm
                        .CommandText = "select dbo.uf_ACCalcularInteres(" & dblTasaInteres & ", '" & strTipoTasa & "', '" & strPeriodoTasa & "', '" & strPeriodoCapitalizacion & "', '" & strBaseAnual & "'," & dblMontoPagado & ", '" & gstrFechaActual & "','" & Convertyyyymmdd(datFechaVencimiento) & "') as InteresAjuste"
                        Set adoConsulta = .Execute
                    End With
    
                    If adoConsulta.EOF Then
                        MsgBox ("Data Inconsistente!")
                        Exit Sub
                    End If
                    dblInteresAjuste = adoConsulta("InteresAjuste")
                End If
                
            Else
                dblMontoInteresAFavor = 0
                dblInteresAjuste = 0
            End If
            
            dblImptoInteresAjuste = dblInteresAjuste * adoDetallePago("PorcenIGV") / 100
            dblInteresAdicional = adoDetallePago("InteresAdicional")
            dblImptoInteresAdicional = adoDetallePago("IGVInteresAdicional")
            
            If strCodModalidadCalculoInteres = "01" Then
                dblAmortizacionPrincipal = adoDetallePago("MontoPago") + dblMontoInteresAFavor
                txtNuevaDeudaParcial.Text = dblNuevoPrincipal - (adoConsulta("InteresAjuste") * (1 + gdblTasaIgv))
            ElseIf strCodModalidadCalculoInteres = "02" Then
                dblAmortizacionPrincipal = dblInteresAjuste + adoDetallePago("MontoPago") - dblMontoInteresMoratorio - dblMontoImptoInteresMoratorio - dblInteresAdicional - dblImptoInteresAdicional
                
                If dblAmortizacionPrincipal < 0 Then
                    dblAmortizacionPrincipal = 0
                End If
                
                txtNuevaDeudaParcial.Text = adoDetallePago("Deuda") - adoDetallePago("MontoPago")
            End If
            
            dblNuevoPrincipal = adoDetallePago("PrincipalAdeudado") - dblAmortizacionPrincipal
           
            fraDetallePagoParcial.Enabled = True
            txtPrincipalAdeudado.Text = adoDetallePago("PrincipalAdeudado")
            txtInteresAFavor.Text = dblMontoInteresAFavor  ' Data contable
            txtPorcenImptoInteresAFavor.Text = adoDetallePago("PorcenIGV")
            txtImptoInteresAFavor.Text = dblMontoInteresAFavor * adoDetallePago("PorcenIGV") / 100
            
            txtInteresAdicional.Text = dblInteresAdicional
            txtPorcenImptoInteresAdicional.Text = adoDetallePago("PorcenIGV")
            txtImptoInteresAdicional.Text = dblImptoInteresAdicional
            
            txtDeudaTotal.Text = adoDetallePago("Deuda")
            txtMontoPagoParcial.Text = adoDetallePago("MontoPago")
            
            txtInteresAjuste.Text = dblInteresAjuste
            txtPorcenImptoInteresAjuste.Text = adoDetallePago("PorcenIGV")
            txtImptoInteresAjuste.Text = dblImptoInteresAjuste
            txtAmortizacionPrincipal.Text = dblAmortizacionPrincipal
            txtTotalNotaCredito.Text = dblInteresAjuste + (dblInteresAjuste * adoDetallePago("PorcenIGV") / 100)
            
            txtNuevoPrincipal.Text = dblNuevoPrincipal
            txtNuevoInteresAFavor.Text = dblMontoInteresAFavor - dblInteresAjuste
            txtPorcenImptoNuevoInteresAFavor.Text = adoDetallePago("PorcenIGV")
            txtImptoNuevoInteresAFavor.Text = (dblMontoInteresAFavor - dblInteresAjuste) * adoDetallePago("PorcenIGV") / 100

        End If
        
        adoDetallePago.MoveNext
    End If
    
    tdgDetallePago.DataSource = adoDetallePago
    
    'Calculando sumario:
    dblTotalPrincipalAdeudado = 0
    dblTotalInteresAFavor = 0
    dblTotalIGVInteresAFavor = 0
    dblTotalInteresAdic = 0
    dblTotalImptoInteresAdic = 0
    dblTotalInteresMor = 0
    dblTotalImptoInteresMor = 0
    dblTotalDeuda = 0
    dblTotalInteresAjuste = 0
    dblImptoTotalInteresAjuste = 0
    
    adoDetallePago.MoveFirst
    
    While Not adoDetallePago.EOF
        dblTotalPrincipalAdeudado = dblTotalPrincipalAdeudado + adoDetallePago("PrincipalAdeudado")
        dblTotalInteresAdic = dblTotalInteresAdic + adoDetallePago("InteresAdicional")
        dblTotalImptoInteresAdic = dblTotalImptoInteresAdic + adoDetallePago("IGVInteresAdicional")
        dblTotalInteresMor = dblTotalInteresMor + adoDetallePago("InteresMoratorio")
        dblTotalImptoInteresMor = dblTotalImptoInteresMor + adoDetallePago("IGVInteresMoratorio")
        dblTotalDeuda = dblTotalDeuda + adoDetallePago("Deuda")
        
        With adoComm
            .CommandText = "select ValorInteresAcumuladoInicial as InteresReconocido, ValorInteresDiferido as InteresAFavor from InversionValorizacionDiaria IVD join InversionOperacion IO on " & _
                            "(IVD.CodFondo = IO.CodFondo and IVD.CodAdministradora = IO.CodAdministradora and IO.CodFile = IVD.CodFile " & _
                            "and IO.CodAnalitica = IVD.CodAnalitica and IO.CodTitulo = IVD.CodTitulo) where IO.NumOperacion = '" & adoDetallePago("NumOperacion") & "' " & _
                            "and IO.CodFondo = '" & gstrCodFondoContable & "' and FechaValorizacion = '" & gstrFechaActual & "'"
            Set adoConsulta = .Execute
        End With
        
        If adoConsulta.EOF Then
            MsgBox ("Data de Valorizacion Diaria Inconsistente!")
            'Exit Sub
        Else
            dblMontoInteresReconocido = adoConsulta("InteresReconocido")
            dblMontoInteresAFavor = adoConsulta("InteresAFavor") ' * (-1)
        End If
        
        '----------------
        With adoComm
            .CommandText = "select FechaVencimiento, TasaInteres,TipoTasa, PeriodoTasa, BaseAnual from InversionOperacion " & _
                            " Where CodFondo = '" & gstrCodFondoContable & "' AND NumOperacion = '" & adoDetallePago("NumOperacion") & "'"
            Set adoConsulta = .Execute
        End With

        If adoConsulta.EOF Then
            MsgBox ("Data operativa Inconsistente!")
            Exit Sub
        End If
        
        datFechaVencimiento = adoConsulta("FechaVencimiento")
        strTipoTasa = adoConsulta("TipoTasa")
        strPeriodoTasa = adoConsulta("PeriodoTasa")
        strBaseAnual = adoConsulta("BaseAnual")
          
        'Si es pago total adelantado:
        If Convertddmmyyyy(gstrFechaActual) < datFechaVencimiento And adoDetallePago("TipoOperacion") = Codigo_Orden_PagoCancelacion And strCodModalidadCalculoInteres = "02" Then
            With adoComm
                .CommandText = "select dbo.uf_ACCalcularInteres(" & dblTasaInteres & ", '" & strTipoTasa & "', '" & strPeriodoTasa & "','" & strPeriodoCapitalizacion & "',  '" & strBaseAnual & "'," & adoDetallePago("PrincipalAdeudado") & ", '" & gstrFechaActual & "','" & Convertyyyymmdd(datFechaVencimiento) & "') as InteresAjuste"
                Set adoConsulta = .Execute
            End With
            
            If adoConsulta.EOF Then
                MsgBox ("Data Inconsistente!")
                Exit Sub
            End If
            
            dblInteresAjuste = dblMontoInteresAFavor - adoConsulta("InteresAjuste")
        Else
            dblInteresAjuste = 0
        End If
        
        'Solo se aplica el ajuste por cobro minimo de intereses si es pago total
        If adoDetallePago("TipoOperacion") = Codigo_Orden_PagoCancelacion Then
            dblMontoInteresAjusteCobroMinimo = adoDetallePago("MontoInteresCobroMinimo") - dblMontoInteresReconocido
            If dblMontoInteresAjusteCobroMinimo < 0 Then
                dblMontoInteresAjusteCobroMinimo = 0
            End If
            
            'Si existe ajuste por cobro minimo de intereses, ya no hay ajuste por pago total adelantado
            If dblMontoInteresAjusteCobroMinimo > 0 Then
                dblInteresAjuste = 0
            End If
        Else
            dblMontoInteresAjusteCobroMinimo = 0
        End If
        
        adoDetallePago("AjusteInteres") = dblInteresAjuste
        adoDetallePago("AjusteCobroMinimo") = dblMontoInteresAjusteCobroMinimo
        
        dblTotalInteresAFavor = dblTotalInteresAFavor + dblMontoInteresAFavor
        dblTotalIGVInteresAFavor = dblTotalIGVInteresAFavor + Round(dblMontoInteresAFavor * adoDetallePago("PorcenIGV") / 100, 2)

        dblTotalInteresAjuste = dblTotalInteresAjuste + dblInteresAjuste
        dblImptoTotalInteresAjuste = dblImptoTotalInteresAjuste + Round(dblInteresAjuste * adoDetallePago("PorcenIGV") / 100, 2)
               
        strUltimaOrden = adoDetallePago("TipoOperacion")
        adoDetallePago.MoveNext
        
    Wend
    
    If adoDetallePago.EOF And strUltimaOrden = Codigo_Orden_PagoCancelacion Then
        fraDetallePagoParcial.Enabled = False
        ClearDetallePagoParcial
    End If
    
    'dblInteresDevuelto = dblInteresAjuste
    dblImptoInteresDevuelto = CDec(txtImptoInteresAjuste.Text)
    
    txtTotalPrincipalAdeudado.Text = dblTotalPrincipalAdeudado
    txtTotalInteresAFavor.Text = dblTotalInteresAFavor
    txtTotalIGVInteresAFavor.Text = dblTotalIGVInteresAFavor
    
    txtAjusteDev.Text = dblTotalInteresAjuste
    txtIGVAjusteDev.Text = dblImptoTotalInteresAjuste
    
    txtTotalInteresAdic.Text = dblTotalInteresAdic
    txtTotalIGVInteresAdic.Text = dblTotalImptoInteresAdic
    txtTotalDeuda.Text = dblTotalDeuda
    dblMontoRecibido = CDec(txtMontoRecibido.Text)
    txtMontoPagado.Text = dblMontoRecibido
    
    txtPorcenImptoAjusteDev.Text = txtPorcenImptoInteresAFavor.Text
    txtPorcenTotalImptoInteresAdic.Text = txtPorcenImptoInteresAFavor.Text
    txtPorcenTotalImptoInteresAFavor.Text = txtPorcenImptoInteresAFavor.Text
    
    If dblTotalDeuda < dblMontoRecibido Then
        dblMargenDevolucion = dblMontoRecibido - dblTotalDeuda
    Else
        dblMargenDevolucion = 0
    End If
    
    txtMargenDevolucion.Text = dblMargenDevolucion
    txtTotalNoConsumido.Text = dblTotalInteresAFavor + dblTotalIGVInteresAFavor - (dblTotalInteresAjuste + dblImptoTotalInteresAjuste)
        
End Sub

Private Sub ClearDetallePagoParcial()
    txtPrincipalAdeudado.Text = 0
    txtInteresAFavor.Text = 0  ' Data contable
    txtPorcenImptoInteresAFavor.Text = 0
    txtImptoInteresAFavor.Text = 0
            
    txtInteresAdicional.Text = 0
    txtPorcenImptoInteresAdicional.Text = 0
    txtImptoInteresAdicional.Text = 0

    txtDeudaTotal.Text = 0
    txtMontoPagoParcial.Text = 0

    txtInteresAjuste.Text = 0
    txtPorcenImptoInteresAjuste.Text = 0
    txtImptoInteresAjuste.Text = 0
    txtAmortizacionPrincipal.Text = 0
    txtTotalNotaCredito.Text = 0
            
    txtNuevoPrincipal.Text = 0
    txtNuevoInteresAFavor.Text = 0
    txtPorcenImptoNuevoInteresAFavor.Text = 0
    txtImptoNuevoInteresAFavor.Text = 0
    txtNuevaDeudaParcial.Text = 0

End Sub

Private Sub ClearDatosAnexo()

    tdgAnexo.Columns(5).FooterText = 0
    tdgAnexo.Columns(6).FooterText = 0
    tdgAnexo.Columns(7).FooterText = 0
    tdgAnexo.Columns(8).FooterText = 0
    tdgAnexo.Columns(9).FooterText = 0

    txtEmisor.Text = Valor_Caracter
    txtNumOrig.Text = Valor_Caracter
    txtInstrumento.Text = Valor_Caracter
    txtClase.Text = Valor_Caracter
    txtSubClase.Text = Valor_Caracter
    txtGestor.Text = Valor_Caracter
    txtLinea.Text = Valor_Caracter
    txtNumeroContrato.Text = Valor_Caracter
    txtCantidadDocumentos.Text = 0
    txtCantidadDocumentosPendientes.Text = 0
    txtFechaInicio.Text = Valor_Caracter
    txtTasaFacial.Text = 0
    txtMoneda.Text = Valor_Caracter
    txtTotalAnexo.Text = 0
    txtTotalAnexoDesc.Text = 0
    txtPorcenDesc.Text = 0
    txtComisiones.Text = 0
    txtMontoRecibido.Text = 0
    txtDeudaRestante.Text = 0
    
    Set adoDetalleAnexo = Nothing
    
    tdgAnexo.DataSource = adoDetalleAnexo
    tdgAnexo.Refresh
    
End Sub

Private Sub cmdSiguiente_Click()
    On Error GoTo CtrlError
    
    If txtMontoRecibido.Value <= 0 Then
        MsgBox "¡El monto pagado debe ser mayor a 0!", vbCritical, "Faltan Datos"
        Exit Sub
    End If
    
    adoDetalleAnexo.MoveFirst
    If adoDetalleAnexo.EOF = True Then
        MsgBox "¡No hay operacion o anexo cargado!", vbCritical, "Faltan Datos"
        Exit Sub
    End If
    
    ClearDetallePagoParcial

'    If strCodFile = "016" Then
'        GenerarDetallePagoFlujos
'    Else
        GenerarDetallePagoFactura
        CalcularTotalesDetallePago
'    End If
    
    With tabPagosRFCP
        .TabEnabled(0) = False
        .TabEnabled(1) = False
        .TabEnabled(2) = True
        .Tab = 2
    End With
    
    Exit Sub
        
CtrlError:
    
    MsgBox "¡No hay operacion o anexo cargado!", vbCritical, "Faltan Datos"
    Exit Sub
  

End Sub

Private Sub Form_Load()
    
    Call InicializarValores
    Call CargarListas
    Call Buscar
 
    With tabPagosRFCP
        .TabEnabled(0) = True
        .TabEnabled(1) = False
        .TabEnabled(2) = False
        .Tab = 0
    End With

    cmdCancelar.Visible = False
    CentrarForm Me
        
    Call ValidaExisteTipoCambio(Codigo_TipoCambio_SBS, gstrFechaActual)
    
End Sub

Private Sub txtMontoRecibido_Change()
    If txtMontoRecibido.Text = "" Then
        txtMontoRecibido.Text = "0.00"
    End If
    txtDeudaRestante.Text = CDec(tdgAnexo.Columns(11).FooterText) - CDec(txtMontoRecibido.Text)
    If CDec(txtDeudaRestante.Text) < 0 Then
        txtDeudaRestante.Text = 0
        'txtMontoRecibido.Text = CDec(tdgAnexo.Columns(9).FooterText)
    End If
End Sub

Private Sub txtMontoRecibido_LostFocus()
    If txtMontoRecibido.Text = "" Then
        txtMontoRecibido.Text = "0.00"
    End If
    txtDeudaRestante.Text = CDec(tdgAnexo.Columns(11).FooterText) - CDec(txtMontoRecibido.Text)
    If CDec(txtDeudaRestante.Text) < 0 Then
        txtDeudaRestante.Text = 0
        'txtMontoRecibido.Text = CDec(tdgAnexo.Columns(9).FooterText)
    End If
End Sub

Private Sub txtNumOrig_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        strNumSolicitud = txtNumOrig.Text
        CargarDesdeCarteraFacturas
    End If
    
End Sub
