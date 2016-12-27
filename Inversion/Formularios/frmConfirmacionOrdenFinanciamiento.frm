VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{830D5A06-9B70-4F8C-98B6-7A19C4E7760B}#1.0#0"; "TAMControls.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmConfirmacionOrdenFinanciamiento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Confirmaciones - Ordenes de Financiamiento"
   ClientHeight    =   9540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13920
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9540
   ScaleWidth      =   13920
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   11490
      TabIndex        =   96
      Top             =   8760
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
      Left            =   540
      TabIndex        =   95
      Top             =   8760
      Width           =   4200
      _ExtentX        =   7408
      _ExtentY        =   1296
      Buttons         =   3
      Caption0        =   "&Confirmar"
      Tag0            =   "3"
      ToolTipText0    =   "Confirmar"
      Caption1        =   "&Eliminar"
      Tag1            =   "4"
      ToolTipText1    =   "Eliminar"
      Caption2        =   "&Buscar"
      Tag2            =   "5"
      ToolTipText2    =   "Buscar"
      UserControlWidth=   4200
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Procesar"
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
      Left            =   10050
      Picture         =   "frmConfirmacionOrdenFinanciamiento.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   94
      Top             =   8760
      Width           =   1200
   End
   Begin TabDlg.SSTab tabConfirmacionOrden 
      Height          =   8625
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   13845
      _ExtentX        =   24421
      _ExtentY        =   15214
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
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
      TabPicture(0)   =   "frmConfirmacionOrdenFinanciamiento.frx":0568
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraCriterio"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "tdgConsulta"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Confirmación"
      TabPicture(1)   =   "frmConfirmacionOrdenFinanciamiento.frx":0584
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblTirNeta(1)"
      Tab(1).Control(1)=   "lblDescrip(16)"
      Tab(1).Control(2)=   "lblTirNeta(0)"
      Tab(1).Control(3)=   "lblDescrip(3)"
      Tab(1).Control(4)=   "lblPrecio(1)"
      Tab(1).Control(5)=   "lblDescrip(2)"
      Tab(1).Control(6)=   "lblPrecio(0)"
      Tab(1).Control(7)=   "lblDescrip(7)"
      Tab(1).Control(8)=   "lblDescrip(45)"
      Tab(1).Control(9)=   "fraDatos"
      Tab(1).Control(10)=   "fraDatosFL1"
      Tab(1).Control(11)=   "txtObservacion"
      Tab(1).Control(12)=   "cmdAccion"
      Tab(1).Control(13)=   "fraAcreencias"
      Tab(1).Control(14)=   "fraNotaCredito"
      Tab(1).ControlCount=   15
      Begin VB.Frame fraNotaCredito 
         Caption         =   "Nota de Crédito"
         Height          =   1335
         Left            =   -74640
         TabIndex        =   120
         Top             =   4410
         Width           =   5775
         Begin TAMControls.TAMTextBox txtInteresAFavor 
            Height          =   285
            Left            =   3030
            TabIndex        =   121
            Top             =   240
            Width           =   2025
            _ExtentX        =   3572
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
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Locked          =   -1  'True
            Container       =   "frmConfirmacionOrdenFinanciamiento.frx":05A0
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            Apariencia      =   1
            Borde           =   1
            MaximoValor     =   0
         End
         Begin TAMControls.TAMTextBox txtIGVAFavor 
            Height          =   285
            Left            =   3030
            TabIndex        =   122
            Top             =   585
            Width           =   2025
            _ExtentX        =   3572
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
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Locked          =   -1  'True
            Container       =   "frmConfirmacionOrdenFinanciamiento.frx":05BC
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            Apariencia      =   1
            Borde           =   1
            MaximoValor     =   0
         End
         Begin TAMControls.TAMTextBox txtTotalNotaCredito 
            Height          =   285
            Left            =   3030
            TabIndex        =   123
            Top             =   945
            Width           =   2025
            _ExtentX        =   3572
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
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Locked          =   -1  'True
            Container       =   "frmConfirmacionOrdenFinanciamiento.frx":05D8
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            Apariencia      =   1
            Borde           =   1
            MaximoValor     =   0
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Index           =   64
            Left            =   705
            TabIndex        =   126
            Top             =   660
            Width           =   330
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Interés a Favor"
            BeginProperty Font 
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
            Index           =   58
            Left            =   705
            TabIndex        =   125
            Top             =   315
            Width           =   1290
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Nota de crédito"
            BeginProperty Font 
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
            Left            =   705
            TabIndex        =   124
            Top             =   1005
            Width           =   1815
         End
      End
      Begin VB.Frame fraAcreencias 
         Caption         =   "Intereses y Comisiones"
         Height          =   4365
         Left            =   -74640
         TabIndex        =   99
         Top             =   2880
         Width           =   5775
         Begin VB.TextBox txtVacCorrido 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   1
            Left            =   2910
            MaxLength       =   45
            TabIndex        =   101
            Top             =   5160
            Width           =   2025
         End
         Begin VB.TextBox txtInteresCorrido 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   2
            Left            =   2910
            MaxLength       =   45
            TabIndex        =   100
            Top             =   4845
            Width           =   2025
         End
         Begin TAMControls.TAMTextBox txtIGVComisionDesembolso 
            Height          =   285
            Index           =   0
            Left            =   3030
            TabIndex        =   102
            TabStop         =   0   'False
            Top             =   3305
            Width           =   2025
            _ExtentX        =   3572
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
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Locked          =   -1  'True
            Container       =   "frmConfirmacionOrdenFinanciamiento.frx":05F4
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            Apariencia      =   1
            Borde           =   1
            MaximoValor     =   0
         End
         Begin TAMControls.TAMTextBox txtCapital 
            Height          =   285
            Left            =   3030
            TabIndex        =   103
            TabStop         =   0   'False
            Top             =   330
            Width           =   2025
            _ExtentX        =   3572
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
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Locked          =   -1  'True
            Container       =   "frmConfirmacionOrdenFinanciamiento.frx":0610
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            Apariencia      =   1
            Borde           =   1
            MaximoValor     =   0
         End
         Begin TAMControls.TAMTextBox txtInteres 
            Height          =   285
            Left            =   3030
            TabIndex        =   104
            TabStop         =   0   'False
            Top             =   855
            Width           =   2025
            _ExtentX        =   3572
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
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Locked          =   -1  'True
            Container       =   "frmConfirmacionOrdenFinanciamiento.frx":062C
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            Apariencia      =   1
            Borde           =   1
            MaximoValor     =   0
         End
         Begin TAMControls.TAMTextBox txtIGVInteres 
            Height          =   285
            Left            =   3030
            TabIndex        =   105
            TabStop         =   0   'False
            Top             =   1230
            Width           =   2025
            _ExtentX        =   3572
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
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Locked          =   -1  'True
            Container       =   "frmConfirmacionOrdenFinanciamiento.frx":0648
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            Apariencia      =   1
            Borde           =   1
            MaximoValor     =   0
         End
         Begin TAMControls.TAMTextBox txtComisionDesembolso 
            Height          =   285
            Left            =   3030
            TabIndex        =   106
            TabStop         =   0   'False
            Top             =   2985
            Width           =   2025
            _ExtentX        =   3572
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
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Locked          =   -1  'True
            Container       =   "frmConfirmacionOrdenFinanciamiento.frx":0664
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            Apariencia      =   1
            Borde           =   1
            MaximoValor     =   0
         End
         Begin TAMControls.TAMTextBox txtMontoTotal 
            Height          =   285
            Left            =   3030
            TabIndex        =   107
            TabStop         =   0   'False
            Top             =   3870
            Width           =   2025
            _ExtentX        =   3572
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
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Locked          =   -1  'True
            Container       =   "frmConfirmacionOrdenFinanciamiento.frx":0680
            Text            =   "0.00"
            Decimales       =   2
            Estilo          =   4
            Apariencia      =   1
            Borde           =   1
            MaximoValor     =   0
         End
         Begin VB.Label lblDescripMoneda 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nuevos Soles (S/.)"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   5
            Left            =   1500
            TabIndex        =   119
            Top             =   4590
            Width           =   1545
         End
         Begin VB.Label lblMontoTotal 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   3
            Left            =   3090
            TabIndex        =   118
            Tag             =   "0.00"
            ToolTipText     =   "Monto Total de la Orden"
            Top             =   4590
            Width           =   2025
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Monto Total"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   55
            Left            =   540
            TabIndex        =   117
            Top             =   4605
            Width           =   855
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Comisión de Desembolso"
            BeginProperty Font 
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
            Left            =   705
            TabIndex        =   116
            Top             =   3060
            Width           =   2100
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   0
            X1              =   300
            X2              =   4920
            Y1              =   5520
            Y2              =   5520
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "VAC Corrido"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   59
            Left            =   330
            TabIndex        =   115
            Top             =   5190
            Width           =   855
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Interés"
            BeginProperty Font 
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
            Index           =   61
            Left            =   705
            TabIndex        =   114
            Top             =   960
            Width           =   585
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Index           =   62
            Left            =   705
            TabIndex        =   113
            Top             =   1305
            Width           =   330
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Index           =   65
            Left            =   705
            TabIndex        =   112
            Top             =   3405
            Width           =   330
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Capital"
            BeginProperty Font 
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
            Index           =   66
            Left            =   705
            TabIndex        =   111
            Top             =   435
            Width           =   600
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "InterÃ©s Corrido"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   67
            Left            =   360
            TabIndex        =   110
            Top             =   4890
            Width           =   1020
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Monto Total"
            BeginProperty Font 
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
            Index           =   68
            Left            =   705
            TabIndex        =   109
            Top             =   3930
            Width           =   1035
         End
         Begin VB.Label lblDescripMoneda 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Height          =   195
            Index           =   8
            Left            =   1335
            TabIndex        =   108
            Top             =   435
            Width           =   1665
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   1
            X1              =   480
            X2              =   5040
            Y1              =   750
            Y2              =   750
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   8
            X1              =   420
            X2              =   5040
            Y1              =   3720
            Y2              =   3720
         End
      End
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   -65010
         TabIndex        =   97
         Top             =   7590
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   1296
         Buttons         =   2
         Caption0        =   "&Procesar"
         Tag0            =   "2"
         ToolTipText0    =   "Procesar"
         Caption1        =   "&Cancelar"
         Tag1            =   "8"
         ToolTipText1    =   "Cancelar"
         UserControlWidth=   2700
      End
      Begin VB.TextBox txtObservacion 
         Height          =   435
         Left            =   -74640
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   77
         Top             =   7620
         Width           =   8535
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmConfirmacionOrdenFinanciamiento.frx":069C
         Height          =   5385
         Left            =   360
         OleObjectBlob   =   "frmConfirmacionOrdenFinanciamiento.frx":06B6
         TabIndex        =   75
         Top             =   2700
         Width           =   13065
      End
      Begin VB.Frame fraDatosFL1 
         Caption         =   "Comisiones y Monto Total (FL1)"
         Height          =   4365
         Left            =   -74640
         TabIndex        =   39
         Top             =   2880
         Width           =   5775
         Begin VB.TextBox txtComisionGastoBancario 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   3090
            TabIndex        =   83
            Top             =   2880
            Width           =   2025
         End
         Begin VB.TextBox txtComisionEspecial 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   3090
            TabIndex        =   82
            Top             =   3210
            Width           =   2025
         End
         Begin VB.TextBox txtComisionFondoG 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   3090
            MaxLength       =   45
            TabIndex        =   81
            Top             =   2235
            Width           =   2025
         End
         Begin VB.TextBox txtVacCorrido 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   2910
            MaxLength       =   45
            TabIndex        =   60
            Top             =   5160
            Width           =   2025
         End
         Begin VB.TextBox txtComisionConasev 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   3090
            MaxLength       =   45
            TabIndex        =   45
            Top             =   2565
            Width           =   2025
         End
         Begin VB.TextBox txtComisionFondo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   3090
            MaxLength       =   45
            TabIndex        =   44
            Top             =   1905
            Width           =   2025
         End
         Begin VB.TextBox txtComisionCavali 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   3090
            MaxLength       =   45
            TabIndex        =   43
            Top             =   1590
            Width           =   2025
         End
         Begin VB.TextBox txtComisionBolsa 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   3090
            MaxLength       =   45
            TabIndex        =   42
            Top             =   1275
            Width           =   2025
         End
         Begin VB.TextBox txtComisionAgente 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   3090
            MaxLength       =   45
            TabIndex        =   41
            Top             =   975
            Width           =   2025
         End
         Begin VB.TextBox txtInteresCorrido 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   2910
            MaxLength       =   45
            TabIndex        =   40
            Top             =   4845
            Width           =   2025
         End
         Begin VB.Label lblDescripMoneda 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Nuevos Soles (S/.)"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   4
            Left            =   1500
            TabIndex        =   91
            Top             =   4560
            Width           =   1545
         End
         Begin VB.Label lblDescripMoneda 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Nuevos Soles (S/.)"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   1590
            TabIndex        =   90
            Top             =   3990
            Width           =   1365
         End
         Begin VB.Label lblMontoTotal 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   2
            Left            =   3090
            TabIndex        =   89
            Tag             =   "0.00"
            ToolTipText     =   "Monto Total de la Orden"
            Top             =   4560
            Width           =   2025
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto Total"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   53
            Left            =   540
            TabIndex        =   88
            Top             =   4575
            Width           =   855
         End
         Begin VB.Label lblSubTotal 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   2
            Left            =   3090
            TabIndex        =   87
            Tag             =   "0.00"
            ToolTipText     =   "SubTotal de la Orden"
            Top             =   600
            Width           =   2025
         End
         Begin VB.Label lblDescripMoneda 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Nuevos Soles (S/.)"
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   2
            Left            =   1800
            TabIndex        =   86
            Top             =   630
            Width           =   765
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Gastos Bancarios"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   51
            Left            =   510
            TabIndex        =   85
            Top             =   2880
            Width           =   1245
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "ComisiÃ³n Especial"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   52
            Left            =   510
            TabIndex        =   84
            Top             =   3210
            Width           =   1275
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "ComisiÃ³n Fondo GarantÃ­a"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   47
            Left            =   510
            TabIndex        =   80
            Top             =   2250
            Width           =   1800
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   4
            X1              =   300
            X2              =   4920
            Y1              =   5490
            Y2              =   5490
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "VAC Corrido"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   17
            Left            =   330
            TabIndex        =   59
            Top             =   5160
            Width           =   855
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "ComisiÃ³n SAB"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   24
            Left            =   510
            TabIndex        =   58
            Top             =   990
            Width           =   990
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "ComisiÃ³n BVL"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   26
            Left            =   510
            TabIndex        =   57
            Top             =   1290
            Width           =   975
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "ComisiÃ³n Cavali"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   27
            Left            =   510
            TabIndex        =   56
            Top             =   1605
            Width           =   1110
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "ComisiÃ³n Fondo LiquidaciÃ³n"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   32
            Left            =   510
            TabIndex        =   55
            Top             =   1935
            Width           =   1980
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "ComisiÃ³n Conasev"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   33
            Left            =   510
            TabIndex        =   54
            Top             =   2580
            Width           =   1305
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "IGV"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   34
            Left            =   540
            TabIndex        =   53
            Top             =   3540
            Width           =   270
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "SubTotal"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   25
            Left            =   540
            TabIndex        =   52
            Top             =   360
            Width           =   645
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "InterÃ©s Corrido"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   35
            Left            =   360
            TabIndex        =   51
            Top             =   4860
            Width           =   1020
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto Total"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   36
            Left            =   540
            TabIndex        =   50
            Top             =   3990
            Width           =   855
         End
         Begin VB.Label lblDescripMoneda 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Nuevos Soles (S/.)"
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   0
            Left            =   1800
            TabIndex        =   49
            Top             =   330
            Width           =   765
         End
         Begin VB.Label lblSubTotal 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   0
            Left            =   3090
            TabIndex        =   48
            Tag             =   "0.00"
            ToolTipText     =   "SubTotal de la Orden"
            Top             =   300
            Width           =   2025
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   3
            X1              =   540
            X2              =   5100
            Y1              =   930
            Y2              =   930
         End
         Begin VB.Label lblComisionIgv 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   0
            Left            =   3090
            TabIndex        =   47
            Tag             =   "0.00"
            ToolTipText     =   "Monto de ComisiÃ³n IGV"
            Top             =   3510
            Width           =   2025
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   2
            X1              =   510
            X2              =   5130
            Y1              =   3855
            Y2              =   3855
         End
         Begin VB.Label lblMontoTotal 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   0
            Left            =   3090
            TabIndex        =   46
            Tag             =   "0.00"
            ToolTipText     =   "Monto Total de la Orden"
            Top             =   3945
            Width           =   2025
         End
      End
      Begin VB.Frame fraDatos 
         Caption         =   "Datos Operación"
         Height          =   1935
         Left            =   -74640
         TabIndex        =   24
         Top             =   480
         Width           =   12765
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Nro. Poliza"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   54
            Left            =   240
            TabIndex        =   93
            Top             =   1230
            Width           =   765
         End
         Begin VB.Label lblNroPoliza 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1305
            TabIndex        =   92
            Top             =   1200
            Width           =   4185
         End
         Begin VB.Label lblNemonico 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   10140
            TabIndex        =   79
            Top             =   1185
            Width           =   1575
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Nemónico"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   46
            Left            =   8580
            TabIndex        =   78
            Top             =   1215
            Width           =   720
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Base Anual"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   9
            Left            =   5820
            TabIndex        =   74
            Top             =   1215
            Width           =   990
         End
         Begin VB.Label lblBaseAnual 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7125
            TabIndex        =   73
            Top             =   1185
            Width           =   1155
         End
         Begin VB.Label lblFechaEmision 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7125
            TabIndex        =   72
            Top             =   270
            Width           =   1155
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Emisión"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   14
            Left            =   5820
            TabIndex        =   71
            Top             =   285
            Width           =   750
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Vencimiento"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   13
            Left            =   5820
            TabIndex        =   70
            Top             =   615
            Width           =   1035
         End
         Begin VB.Label lblFechaVencimiento 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7125
            TabIndex        =   69
            Top             =   585
            Width           =   1155
         End
         Begin VB.Label lblTasa 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7125
            TabIndex        =   68
            Top             =   885
            Width           =   1155
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Tasa"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   10
            Left            =   5820
            TabIndex        =   67
            Top             =   915
            Width           =   555
         End
         Begin VB.Label lblEmisor 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1305
            TabIndex        =   66
            Top             =   885
            Width           =   4185
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Emisor"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   11
            Left            =   240
            TabIndex        =   65
            Top             =   900
            Width           =   990
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Liquidación"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   5
            Left            =   3150
            TabIndex        =   38
            Top             =   1560
            Width           =   810
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Cambio"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   8
            Left            =   8580
            TabIndex        =   37
            Top             =   870
            Width           =   1110
         End
         Begin VB.Label lblFechaLiquidacion 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4335
            TabIndex        =   36
            Top             =   1515
            Width           =   1155
         End
         Begin VB.Label lblTipoCambio 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   10140
            TabIndex        =   35
            Top             =   870
            Width           =   1575
         End
         Begin VB.Label lblValorNominal 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   10140
            TabIndex        =   34
            Top             =   270
            Width           =   1575
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Valor Nominal"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   12
            Left            =   8580
            TabIndex        =   33
            Top             =   270
            Width           =   1230
         End
         Begin VB.Label lblCantOrden 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   10140
            TabIndex        =   32
            Top             =   570
            Width           =   1575
         End
         Begin VB.Label lblFechaOperacion 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1305
            TabIndex        =   31
            Top             =   1515
            Width           =   1155
         End
         Begin VB.Label lblAgente 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1305
            TabIndex        =   30
            Top             =   570
            Width           =   4185
         End
         Begin VB.Label lblFondo 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1305
            TabIndex        =   29
            Top             =   270
            Width           =   4185
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad Nominal"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   6
            Left            =   8580
            TabIndex        =   28
            Top             =   570
            Width           =   1245
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Operación"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   27
            Top             =   1530
            Width           =   735
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Agente"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   1
            Left            =   240
            TabIndex        =   26
            Top             =   600
            Width           =   990
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Fondo"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   240
            TabIndex        =   25
            Top             =   270
            Width           =   990
         End
      End
      Begin VB.Frame fraCriterio 
         Caption         =   "Criterios de Búsqueda"
         Height          =   2055
         Left            =   360
         TabIndex        =   9
         Top             =   480
         Width           =   13065
         Begin VB.CommandButton cmdExportarExcel 
            Caption         =   "&Excel"
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
            Left            =   11640
            Picture         =   "frmConfirmacionOrdenFinanciamiento.frx":5F0A
            Style           =   1  'Graphical
            TabIndex        =   98
            Top             =   1200
            Width           =   1200
         End
         Begin VB.ComboBox cboTipoOrden 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   61
            Top             =   1080
            Width           =   4995
         End
         Begin VB.ComboBox cboEstado 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   1440
            Width           =   4995
         End
         Begin VB.ComboBox cboTipoInstrumento 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   720
            Width           =   4995
         End
         Begin VB.ComboBox cboFondo 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   360
            Width           =   4995
         End
         Begin MSComCtl2.DTPicker dtpFechaOrdenDesde 
            Height          =   285
            Left            =   9120
            TabIndex        =   13
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
            Left            =   11400
            TabIndex        =   14
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
         Begin MSComCtl2.DTPicker dtpFechaLiquidacionDesde 
            Height          =   285
            Left            =   9120
            TabIndex        =   15
            Top             =   840
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
         Begin MSComCtl2.DTPicker dtpFechaLiquidacionHasta 
            Height          =   285
            Left            =   11400
            TabIndex        =   16
            Top             =   840
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
            Caption         =   "Desde"
            Height          =   195
            Index           =   42
            Left            =   8520
            TabIndex        =   64
            Top             =   840
            Width           =   465
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            Height          =   195
            Index           =   18
            Left            =   10800
            TabIndex        =   63
            Top             =   840
            Width           =   420
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Orden de"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   15
            Left            =   360
            TabIndex        =   62
            Top             =   1100
            Width           =   660
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Estado"
            Height          =   195
            Index           =   23
            Left            =   360
            TabIndex        =   23
            Top             =   1460
            Width           =   495
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Instrumento"
            Height          =   195
            Index           =   22
            Left            =   360
            TabIndex        =   22
            Top             =   740
            Width           =   825
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            Height          =   195
            Index           =   21
            Left            =   10800
            TabIndex        =   21
            Top             =   360
            Width           =   420
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            Height          =   195
            Index           =   20
            Left            =   8520
            TabIndex        =   20
            Top             =   360
            Width           =   465
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fondo"
            Height          =   195
            Index           =   19
            Left            =   360
            TabIndex        =   19
            Top             =   380
            Width           =   450
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Orden"
            Height          =   195
            Index           =   43
            Left            =   6960
            TabIndex        =   18
            Top             =   345
            Width           =   930
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Liquidación"
            Height          =   195
            Index           =   44
            Left            =   6960
            TabIndex        =   17
            Top             =   840
            Width           =   1305
         End
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   45
         Left            =   -74610
         TabIndex        =   76
         Top             =   7290
         Width           =   1065
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Precio (FL1)"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   7
         Left            =   -74580
         TabIndex        =   8
         Top             =   2595
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label lblPrecio 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   0
         Left            =   -73335
         TabIndex        =   7
         Top             =   2565
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Precio (FL2)"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   2
         Left            =   -68070
         TabIndex        =   6
         Top             =   2580
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label lblPrecio 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   1
         Left            =   -66930
         TabIndex        =   5
         Top             =   2565
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Tir Neta (FL1)"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   3
         Left            =   -71625
         TabIndex        =   4
         Top             =   2595
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label lblTirNeta 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   0
         Left            =   -70230
         TabIndex        =   3
         Top             =   2565
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Tir Neta (FL2)"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   16
         Left            =   -65220
         TabIndex        =   2
         Top             =   2565
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label lblTirNeta 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   1
         Left            =   -63765
         TabIndex        =   1
         Top             =   2565
         Visible         =   0   'False
         Width           =   1350
      End
   End
End
Attribute VB_Name = "frmConfirmacionOrdenFinanciamiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Ficha de ConfirmaciÃ³n de Ordenes y GeneraciÃ³n de Operaciones"
Option Explicit

Dim arrFondo()              As String, arrEstado()              As String
Dim arrTipoInstrumento()    As String, arrTipoOrden()           As String

Dim strCodFondo             As String, strCodMoneda             As String
Dim strCodTipoInstrumento   As String, strCodEstado             As String
Dim strCodTipoOrden         As String
Dim strEstado               As String, strSQL                   As String
Dim strFechaOperacion       As String, strFechaSiguiente        As String
Dim strCodFile              As String, strCodAnalitica          As String
Dim strCodDetalleFile       As String, strCodTipoOrdenBusqueda  As String
Dim strCodSubDetalleFile    As String
Dim adoRegistroAux          As ADODB.Recordset
Dim adoConsulta             As ADODB.Recordset
Dim indSortAsc              As Boolean, indSortDesc             As Boolean

Public oExportacion         As clsExportacion
Public indOk                As Boolean
Dim adoExportacion          As ADODB.Recordset

Public Sub Adicionar()

End Sub

Private Sub CalculoTotal(Index As Integer)

    Dim curComImp As Currency, curMonTotal As Currency

    If Not IsNumeric(txtComisionAgente(Index).Text) And Not IsNumeric(txtComisionBolsa(Index).Text) And Not IsNumeric(txtComisionConasev(Index).Text) And Not IsNumeric(txtComisionCavali(Index).Text) And Not IsNumeric(txtComisionFondo(Index).Text) And Not IsNumeric(txtComisionFondoG(Index).Text) Then Exit Sub
    
    curComImp = CCur(CCur(txtComisionAgente(Index).Text) + CCur(txtComisionBolsa(Index).Text) + CCur(txtComisionConasev(Index).Text) + CCur(txtComisionCavali(Index).Text) + CCur(txtComisionFondo(Index).Text) + CCur(txtComisionFondoG(Index).Text)) * gdblTasaIgv
    lblComisionIgv(Index).Caption = CStr(curComImp)

    curComImp = CCur(txtComisionAgente(Index).Text) + CCur(txtComisionBolsa(Index).Text) + CCur(txtComisionConasev(Index).Text) + CCur(txtComisionCavali(Index).Text) + CCur(txtComisionFondo(Index).Text) + CCur(txtComisionFondoG(Index).Text) + CCur(txtComisionGastoBancario(Index).Text) + CCur(txtComisionEspecial(Index).Text) + CCur(lblComisionIgv(Index).Caption)

    If tdgConsulta.Columns(7) = Codigo_Orden_Compra Then   '*** Compra ***
        curMonTotal = CCur(lblSubTotal(Index).Caption) + curComImp + CCur(txtVacCorrido(Index).Text)
    ElseIf strCodTipoOrden = Codigo_Orden_Venta Or strCodTipoOrden = Codigo_Orden_Quiebre Then  '*** Venta ***
        curMonTotal = CCur(lblSubTotal(Index).Caption) - curComImp
    End If

    curMonTotal = curMonTotal + CCur(txtInteresCorrido(Index).Text)
    
    lblMontoTotal(Index).Caption = CStr(curMonTotal)
    
End Sub

Public Sub Eliminar()

    If strEstado = Reg_Consulta Or strEstado = Reg_Edicion Then
        Dim strMensaje  As String
        
        'verificar si la orden no está¡ ya anulada
        
        If strCodEstado <> Estado_Orden_Anulada And strCodEstado <> Estado_Orden_Procesada Then
        
            strMensaje = "Se procederá a eliminar la ORDEN" & vbNewLine & vbNewLine & _
                "Número" & Space(6) & ":" & Space(1) & tdgConsulta.Columns(1) & vbNewLine & _
                "Descripción" & ":" & Space(1) & Trim(tdgConsulta.Columns(5)) & vbNewLine & vbNewLine & vbNewLine & _
                "¿ Seguro de continuar ?"
            
            If MsgBox(strMensaje, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
        
                '*** Anular Orden ***
                adoComm.CommandText = "UPDATE FinanciamientoOrden SET EstadoOrden='" & Estado_Orden_Anulada & "' WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                    "NumOrden='" & Trim(tdgConsulta.Columns(1)) & "'"
                adoConn.Execute adoComm.CommandText
                
                MsgBox Mensaje_Eliminacion_Exitosa, vbExclamation, Me.Caption
                
                tabConfirmacionOrden.Tab = 0
                Call Buscar
                
                Exit Sub
            End If
            
        Else

            If strCodEstado = Estado_Orden_Anulada Then
                MsgBox "La orden " & Trim(tdgConsulta.Columns(1)) & " ya ha sido anulada.", vbExclamation, "Anular Orden"
            Else
                MsgBox "La orden " & Trim(tdgConsulta.Columns(1)) & " ya ha sido procesada." & vbNewLine & "No se puede anular.", vbCritical, "Anular Orden"
            End If
        End If
        
        
    End If
    
End Sub
Public Function TodoOK() As Boolean

    Dim adoTemporal As ADODB.Recordset
        
    TodoOK = False
    
    Set adoTemporal = New ADODB.Recordset
                
    '*** Verificar Dinamica Contable ***
'    adoComm.CommandText = "SELECT * FROM DinamicaContable " & _
'        "WHERE TipoOperacion='" & strCodTipoOrden & "' AND CodFile='" & strCodFile & "' AND " & _
'        "(CodDetalleFile = '" & strCodDetalleFile & "' OR CodDetalleFile='000') AND " & _
'        "(CodSubDetalleFile = '" & strCodSubDetalleFile & "' OR CodSubDetalleFile = '000') AND " & _
'        "CodAdministradora='" & gstrCodAdministradora & "' AND CodMoneda = '" & _
'        IIf((strCodMoneda) = Codigo_Moneda_Local, Codigo_Moneda_Local, Codigo_Moneda_Extranjero) & "'"
'    Set adoTemporal = adoComm.Execute
'
'    If adoTemporal.EOF Then
'        MsgBox "NO EXISTE Dinámica Contable para la operación", vbCritical, Me.Caption
'        adoTemporal.Close: Set adoTemporal = Nothing
'        Exit Function
'    End If
    
    'adoTemporal.Close
    Set adoTemporal = Nothing

    If tdgConsulta.SelBookmarks.Count - 1 < 0 Then
        MsgBox "No se ha seleccionado ningún registro", vbCritical, Me.Caption
        Exit Function
    End If
    
    If adoConsulta.RecordCount = 0 Then
        MsgBox "No existen registros para procesar", vbCritical, Me.Caption
        Exit Function
    End If
    
    TodoOK = True

End Function

Public Sub Grabar()

    If strEstado = Reg_Defecto Then Exit Sub
    
    Dim adoRegistro     As ADODB.Recordset
    Dim strNumCobertura As String
    Dim adoError        As ADODB.Error
    Dim strErrMsg       As String
    Dim intAccion       As Integer
    Dim lngNumError     As Long
    
    On Error GoTo CtrlError
    
'    Set adoRegistro = New ADODB.Recordset
'    strNumCobertura = Valor_Caracter
'    With adoComm
'        .CommandText = "SELECT NumCobertura FROM InversionCobertura WHERE CodTitulo='" & tdgConsulta.Columns(2) & "' AND " & _
'            "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
'        Set adoRegistro = .Execute
'
'        If Not adoRegistro.EOF Then
'            strNumCobertura = adoRegistro("NumCobertura")
'        End If
'
'        adoRegistro.Close: Set adoRegistro = Nothing
'    End With
        
    Call cmdProcesar_Click
    
    Call Cancelar
    
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

Public Sub Imprimir()

End Sub

Public Sub Modificar()

    If strEstado = Reg_Defecto Then Exit Sub
    
    If Trim(tdgConsulta.Columns(7)) <> Estado_Orden_Enviada Then
        MsgBox "Solo se pueden confirmar las Ordenes con estado ENVIADAS A BACKOFFICE.", vbCritical, Me.Caption
        Exit Sub
    End If
    
    If strEstado = Reg_Consulta Then
        strEstado = Reg_Edicion
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabConfirmacionOrden
            .TabEnabled(0) = False
            .Tab = 1
        End With
        'Call Habilita
    End If
    
End Sub

Private Sub LlenarFormulario(strModo As String)

    Dim adoRegistro     As ADODB.Recordset, adoTemporal     As ADODB.Recordset
    Dim strSQL          As String, strCodBaseAnual          As String
    Dim intRegistro     As Integer
    Dim InteresNC       As Double
    Dim IGVNC           As Double
    Dim TotalNC         As Double
    
    Select Case strModo
        Case Reg_Edicion
            Set adoRegistro = New ADODB.Recordset
            Set adoTemporal = New ADODB.Recordset

            With adoComm
            
            .CommandText = "SELECT * FROM FinanciamientoOrden WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "NumOrden='" & tdgConsulta.Columns(1) & "'"
            Set adoRegistro = .Execute

            If Not adoRegistro.EOF Then
                fraDatos.Caption = "Datos Operación : " & Trim(tdgConsulta.Columns(5))
            
                lblEmisor.Caption = Valor_Caracter
                .CommandText = "SELECT DescripPersona FROM InstitucionPersona WHERE CodPersona='" & adoRegistro("CodAcreedor") & "' AND " & _
                    "TipoPersona='" & Codigo_Tipo_Persona_Emisor & "'"
                Set adoTemporal = .Execute
                
                If Not adoTemporal.EOF Then
                    lblEmisor.Caption = Trim(adoTemporal("DescripPersona"))
                End If
                adoTemporal.Close: Set adoTemporal = Nothing
                
                strCodFile = Trim(adoRegistro("CodFile"))
                strCodAnalitica = Trim(adoRegistro("CodAnalitica"))
                strCodDetalleFile = Trim(adoRegistro("CodDetalleFile"))
                strCodTipoOrden = Trim(adoRegistro("TipoOrden"))
                
                strCodBaseAnual = Trim(adoRegistro("BaseAnual"))
                If strCodBaseAnual = Codigo_Base_Actual_Actual Or strCodBaseAnual = Codigo_Base_Actual_365 Or strCodBaseAnual = Codigo_Base_30_365 Then
                    lblBaseAnual.Caption = "365"
                Else
                    lblBaseAnual.Caption = "360"
                End If
                lblNemonico.Caption = Trim(adoRegistro("Nemotecnico"))
                                                                        
                lblNroPoliza.Caption = adoRegistro("NumDocumento")
                
                strCodMoneda = Trim(adoRegistro("CodMoneda"))
                            
                lblDescripMoneda(0).Caption = ObtenerSignoMoneda(adoRegistro("CodMoneda"))
                lblDescripMoneda(2).Caption = ObtenerSignoMoneda(adoRegistro("CodMoneda"))
                lblDescripMoneda(3).Caption = ObtenerSignoMoneda(adoRegistro("CodMoneda"))
                lblDescripMoneda(4).Caption = ObtenerSignoMoneda(adoRegistro("CodMoneda"))
                                
                'lblDescripMoneda(1).Caption = ObtenerSignoMoneda(adoRegistro("CodMoneda"))
                
                lblFondo.Caption = Trim(cboFondo.Text)
                lblFechaOperacion.Caption = CStr(adoRegistro("FechaOrden"))
                lblFechaLiquidacion.Caption = CStr(adoRegistro("FechaLiquidacion"))
                lblFechaEmision.Caption = CStr(adoRegistro("FechaEmision"))
                lblFechaVencimiento.Caption = CStr(adoRegistro("FechaVencimiento"))
                strFechaOperacion = Convertyyyymmdd(CVDate(adoRegistro("FechaOrden")))
                strFechaSiguiente = Convertyyyymmdd(DateAdd("d", 1, CVDate(adoRegistro("FechaOrden"))))
                
                lblTasa.Caption = CStr(adoRegistro("TasaInteres"))
                lblValorNominal.Caption = CStr(adoRegistro("ValorNominal"))
                lblTipoCambio.Caption = CStr(adoRegistro("ValorTipoCambio"))
                lblCantOrden.Caption = CStr(adoRegistro("CantOrden"))
                                
                '*** Préstamos ***
                If adoRegistro("CodFile") = CodFile_Financiamiento_Prestamos Then
                    fraAcreencias.Visible = True
                    fraNotaCredito.Visible = True
                    txtCapital.Text = CStr(adoRegistro("CantOrden"))
                    txtInteres.Text = CStr(adoRegistro("MontoInteres"))
                    txtIGVInteres.Text = CStr(adoRegistro("MontoImptoInteres"))
                    txtInteresAFavor.Text = 0
                    txtIGVAFavor.Text = 0
                    
                    txtComisionDesembolso.Text = CStr(adoRegistro("MontoComision"))
                    txtIGVComisionDesembolso(0).Text = CStr(adoRegistro("MontoImptoComision"))
                    txtMontoTotal.Text = CStr(adoRegistro("MontoVencimiento"))
                Else
                    fraAcreencias.Visible = False
                    fraNotaCredito.Visible = False
                End If
                
                txtObservacion.Text = Trim(adoRegistro("Observacion"))
           
                lblPrecio(1).Visible = True: lblTirNeta(1).Visible = True
                lblDescrip(2).Visible = True: lblDescrip(16).Visible = True
            
            End If
            adoRegistro.Close: Set adoRegistro = Nothing
    
            End With
    End Select
    
End Sub
Public Sub SubImprimir(Index As Integer)

    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    Dim strFechaDesde           As String, strFechaHasta        As String
    Dim strSeleccionRegistro    As String

    If tabConfirmacionOrden.Tab = 1 Then Exit Sub
    
    Select Case Index
        Case 1
            gstrNameRepo = "InversionOrden"
            
            strSeleccionRegistro = "{InversionOrden.FechaOrden} IN 'Fch1' TO 'Fch2'"
            gstrSelFrml = strSeleccionRegistro
            frmRangoFecha.Show vbModal
                        
            If gstrSelFrml <> "0" Then
                Set frmReporte = New frmVisorReporte
    
                ReDim aReportParamS(5)
                ReDim aReportParamFn(5)
                ReDim aReportParamF(5)
                            
                aReportParamFn(0) = "Usuario"
                aReportParamFn(1) = "FechaDesde"
                aReportParamFn(2) = "FechaHasta"
                aReportParamFn(3) = "Hora"
                aReportParamFn(4) = "Fondo"
                aReportParamFn(5) = "NombreEmpresa"
                            
                aReportParamF(0) = gstrLogin
                aReportParamF(1) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10)
                aReportParamF(2) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10)
                aReportParamF(3) = Format(Time(), "hh:mm:ss")
                aReportParamF(4) = Trim(cboFondo.Text)
                aReportParamF(5) = gstrNombreEmpresa & Space(1)
                            
                aReportParamS(0) = strCodFondo
                aReportParamS(1) = gstrCodAdministradora
                aReportParamS(2) = Convertyyyymmdd(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10))
                aReportParamS(3) = Convertyyyymmdd(DateAdd("d", 1, CVDate(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10))))
                aReportParamS(4) = strCodMoneda
                aReportParamS(5) = strCodTipoInstrumento
            End If
        Case 2
            gstrNameRepo = "InversionOperacion"
            
            strSeleccionRegistro = "{InversionOperacion.FechaOperacion} IN 'Fch1' TO 'Fch2'"
            gstrSelFrml = strSeleccionRegistro
            frmRangoFecha.Show vbModal
                        
            If gstrSelFrml <> "0" Then
                Set frmReporte = New frmVisorReporte
    
                ReDim aReportParamS(5)
                ReDim aReportParamFn(5)
                ReDim aReportParamF(5)
                            
                aReportParamFn(0) = "Usuario"
                aReportParamFn(1) = "FechaDesde"
                aReportParamFn(2) = "FechaHasta"
                aReportParamFn(3) = "Hora"
                aReportParamFn(4) = "Fondo"
                aReportParamFn(5) = "NombreEmpresa"
                            
                aReportParamF(0) = gstrLogin
                aReportParamF(1) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10)
                aReportParamF(2) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10)
                aReportParamF(3) = Format(Time(), "hh:mm:ss")
                aReportParamF(4) = Trim(cboFondo.Text)
                aReportParamF(5) = gstrNombreEmpresa & Space(1)
                            
                aReportParamS(0) = strCodFondo
                aReportParamS(1) = gstrCodAdministradora
                aReportParamS(2) = Convertyyyymmdd(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10))
                aReportParamS(3) = Convertyyyymmdd(DateAdd("d", 1, CVDate(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10))))
                aReportParamS(4) = strCodMoneda
                aReportParamS(5) = strCodTipoInstrumento
            End If
            
    End Select

    If gstrSelFrml = "0" Then Exit Sub
    
    gstrSelFrml = Valor_Caracter
    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"

    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())

    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
    frmReporte.Show vbModal

    Set frmReporte = Nothing

    Screen.MousePointer = vbNormal
    
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
        '*** Fecha Vigente, Moneda ***
        .CommandText = "{ call up_ACSelDatosParametro(23,'" & strCodFondo & "','" & gstrCodAdministradora & "') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            gdatFechaActual = CVDate(adoRegistro("FechaCuota"))
            gstrFechaActual = Convertyyyymmdd(gdatFechaActual)
            dtpFechaOrdenDesde.Value = gdatFechaActual
            dtpFechaOrdenHasta.Value = dtpFechaOrdenDesde.Value

                        
            frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
            
            gstrPeriodoActual = Format(Year(gdatFechaActual), "0000")
            gstrMesActual = Format(Month(gdatFechaActual), "00")
            
            'ACTUALIZA PARAMETROS GLOBALES POR FONDO
            If Not CargarParametrosGlobales(strCodFondo) Then Exit Sub

        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
End Sub


Private Sub cboTipoInstrumento_Click()

    strCodTipoInstrumento = Valor_Caracter
    If cboTipoInstrumento.ListIndex < 0 Then Exit Sub
    
    strCodTipoInstrumento = Trim(arrTipoInstrumento(cboTipoInstrumento.ListIndex))
    
    '*** Tipo de Orden segun el tipo de instrumento ***
    strSQL = "SELECT IFTON.CodTipoOperacion CODIGO,DescripParametro DESCRIP " & _
        "FROM InversionFileTipoOperacionNegociacion IFTON JOIN AuxiliarParametro TON ON(TON.CodParametro=IFTON.CodTipoOperacion AND TON.CodTipoParametro='OPECAJ' AND TON.ValorParametro = 'I')" & _
        "WHERE IFTON.CodFile='" & strCodTipoInstrumento & "' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboTipoOrden, arrTipoOrden(), Valor_Caracter
    
End Sub


Private Sub cboTipoOrden_Click()

    strCodTipoOrdenBusqueda = ""
    If cboTipoOrden.ListIndex < 0 Then Exit Sub

    strCodTipoOrdenBusqueda = Trim(arrTipoOrden(cboTipoOrden.ListIndex))

End Sub


Private Sub cmdExportarExcel_Click()
    Call ExportarExcel
End Sub

Private Sub ExportarExcel()
    
    Dim adoRegistro As ADODB.Recordset
    Dim execSQL As String
    Dim rutaExportacion As String

    Dim datFechaSiguiente As Date
    Dim strFechaLiquidacionHasta As String

    Set frmFormulario = frmConfirmacionOrdenFinanciamiento

    Set adoRegistro = New ADODB.Recordset

    'If TodoOK() Then

        Dim strNameProc As String

        gstrNameRepo = "ConfirmacionOrden"

        strNameProc = ObtenerBaseReporte(gstrNameRepo)

        Dim arrParmS(7)

        arrParmS(0) = Trim(strCodFondo)
        arrParmS(1) = Trim(gstrCodAdministradora)

        If strCodTipoInstrumento <> Valor_Caracter Then
            arrParmS(2) = Trim(strCodTipoInstrumento)
        Else
            arrParmS(2) = "%"
        End If
        
        If strCodTipoOrdenBusqueda <> Valor_Caracter Then
            arrParmS(3) = Trim(strCodTipoOrdenBusqueda)
        Else
            arrParmS(3) = "%"
        End If

        If IsNull(dtpFechaOrdenDesde.Value) And IsNull(dtpFechaOrdenHasta.Value) Then
            arrParmS(4) = Convertyyyymmdd(dtpFechaLiquidacionDesde.Value)
            datFechaSiguiente = DateAdd("d", 1, dtpFechaLiquidacionHasta.Value)
            strFechaLiquidacionHasta = Convertyyyymmdd(datFechaSiguiente)
            arrParmS(5) = strFechaLiquidacionHasta
            arrParmS(6) = "L"
        Else
            arrParmS(4) = Convertyyyymmdd(dtpFechaOrdenDesde.Value)
            arrParmS(5) = Convertyyyymmdd(dtpFechaOrdenHasta.Value)
            arrParmS(6) = "O"
        End If
        
        If strCodEstado <> Valor_Caracter Then
            arrParmS(7) = strCodEstado
        Else
            MsgBox "Debe seleccionar un Estado.", vbCritical, Me.Caption
            If cboEstado.Enabled Then cboEstado.SetFocus
            Exit Sub
        End If

        execSQL = ObtenerCommandText(strNameProc, arrParmS())

        With adoComm

            .CommandText = execSQL

            Set adoRegistro = .Execute

        End With

        Set oExportacion = New clsExportacion

        Call ConfiguraRecordsetExportacion

        Call LlenarRecordsetExportacion(adoRegistro)

        If adoExportacion.RecordCount > 0 Then

            frmRutaGrabar.Show vbModal

            If indOk = True Then

                Screen.MousePointer = vbHourglass

                rutaExportacion = gs_FormName

                If oExportacion.ExportaRecordSetExcel(adoExportacion, gstrNameRepo, rutaExportacion) Then
                    MsgBox "Exportacion realizada", vbInformation
                Else
                    MsgBox "Fallo en exportacion", vbCritical
                End If

                Set oExportacion = Nothing

            End If

        Else
            MsgBox "No existen registros, exportacion a excel cancelada", vbExclamation
        End If

        Screen.MousePointer = vbDefault

    'End If
        
End Sub

Private Function ObtenerBaseReporte(ByVal strNombreReporte As String) As String
    
    ObtenerBaseReporte = Valor_Caracter
    
    Dim crxAplicacion As CRAXDRT.Application
    Dim crxReporte As CRAXDRT.Report
    Dim strReportPath As String
    Dim strBase As String
    Dim intIndex As Integer
        
    strReportPath = gstrRptPath & strNombreReporte & ".RPT"
    
    On Error GoTo Ctrl_Error
    
    Set crxAplicacion = New CRAXDRT.Application

    Set crxReporte = crxAplicacion.OpenReport(strReportPath)

    strBase = crxReporte.Database.Tables(1).Name

    intIndex = InStr(1, strBase, ";", vbBinaryCompare)
        
    strBase = Mid(strBase, 1, intIndex - 1)
    
    ObtenerBaseReporte = strBase

    Set crxReporte = Nothing
    Set crxAplicacion = Nothing
    
    Exit Function
    
Ctrl_Error:
MsgBox "Error al obtener la base del Reporte", vbCritical
Exit Function

End Function

Private Function ObtenerCommandText(ByVal strCadena As String, ByRef arrParametros()) As String
    
    Dim strParametros As String
    Dim i As Integer
    
    strParametros = "{ call " & strCadena & " ("
    
    For i = 0 To UBound(arrParametros)
    
        strParametros = strParametros & "'" & arrParametros(i) & "'" & ","
    
    Next
    
    strParametros = Mid(strParametros, 1, Len(strParametros) - 1)
    
    strParametros = strParametros & ") }"
    
    ObtenerCommandText = strParametros

End Function

Private Sub ConfiguraRecordsetExportacion()

    Set adoExportacion = New ADODB.Recordset

    With adoExportacion
       .CursorLocation = adUseClient
       .Fields.Append "NumOrden", adChar, 10
       .Fields.Append "FechaOrden", adDate
       .Fields.Append "FechaLiquidacion", adDate
       .Fields.Append "CodTitulo", adChar, 15
       .Fields.Append "Nemotecnico", adChar, 15
       .Fields.Append "EstadoOrden", adChar, 2
       
       .Fields.Append "CodFile", adChar, 3
       .Fields.Append "CodAnalitica", adChar, 8
       .Fields.Append "TipoOrden", adChar, 2
       
       .Fields.Append "CodMoneda", adChar, 2
       .Fields.Append "DescripOrden", adVarChar, 100
       .Fields.Append "CantOrden", adDecimal
       .Fields.Append "ValorNominal", adDecimal
       .Fields.Append "PrecioUnitarioMFL1", adDecimal
       .Fields.Append "MontoTotalMFL1", adDecimal
       .Fields.Append "DescripMoneda", adChar, 3
'       .CursorType = adOpenStatic

       .LockType = adLockBatchOptimistic
    End With
    
    adoExportacion.Open
    
End Sub

Private Sub LlenarRecordsetExportacion(ByRef adoRecords As ADODB.Recordset)
        
    Dim dblTipoCambio As Double, dblBruto As Double, dblSAB As Double, dblBVL As Double, dblFondo As Double
    Dim dblCavali As Double, dblFdoCavali As Double, dblConasev As Double, dblTotCom As Double, dblComBroker As Double
    Dim dblCotiza As Double, dblIGV As Double, dblIGVCavali As Double, dblNeto As Double
        
    'dblTipoCambio = CDbl(txtTipoCambio.Text)
        
    If Not adoRecords.EOF Then
    
        Do Until adoRecords.EOF
                
                adoExportacion.AddNew
                
                adoExportacion.Fields("NumOrden") = Trim(adoRecords.Fields("NumOrden"))
                adoExportacion.Fields("FechaOrden") = Trim(adoRecords.Fields("FechaOrden"))
                adoExportacion.Fields("FechaLiquidacion") = adoRecords.Fields("FechaLiquidacion")
                adoExportacion.Fields("CodTitulo") = Trim(adoRecords.Fields("CodTitulo"))
                adoExportacion.Fields("Nemotecnico") = Trim(adoRecords.Fields("Nemotecnico"))
                adoExportacion.Fields("EstadoOrden") = Trim(adoRecords.Fields("EstadoOrden"))
                
                adoExportacion.Fields("CodFile") = Trim(adoRecords.Fields("CodFile"))
                adoExportacion.Fields("CodAnalitica") = Trim(adoRecords.Fields("CodAnalitica"))
                adoExportacion.Fields("TipoOrden") = adoRecords.Fields("TipoOrden")
                
                adoExportacion.Fields("CodMoneda") = adoRecords.Fields("CodMoneda")
                adoExportacion.Fields("DescripOrden") = adoRecords.Fields("DescripOrden")
                adoExportacion.Fields("CantOrden") = adoRecords.Fields("CantOrden")
                adoExportacion.Fields("ValorNominal") = adoRecords.Fields("ValorNominal")
                adoExportacion.Fields("PrecioUnitarioMFL1") = adoRecords.Fields("PrecioUnitarioMFL1")
                adoExportacion.Fields("MontoTotalMFL1") = adoRecords.Fields("MontoTotalMFL1")
                adoExportacion.Fields("DescripMoneda") = adoRecords.Fields("DescripMoneda")
                
                adoExportacion.Update
    
                adoRecords.MoveNext
                
        Loop
        
        adoRecords.Close: Set adoRecords = Nothing
    
    End If

End Sub


Private Sub cmdProcesar_Click()

    Dim strFechaProceso             As String
    Dim intRegistro                 As Integer, intContador         As Integer
    Dim strFinanciamientoOrdenXML   As String
    Dim objFinanciamientoOrdenXML   As DOMDocument60
    Dim strMsgError                 As String
    
    
    If Not TodoOK() Then Exit Sub
    
    Me.MousePointer = vbHourglass
    
    strFechaProceso = Convertyyyymmdd(gdatFechaActual) & Space(1) & Format(Time, "hh:mm")

    intContador = tdgConsulta.SelBookmarks.Count - 1
        
    Call ConfiguraRecordsetAuxiliar
    
    For intRegistro = 0 To intContador
               
        adoConsulta.MoveFirst
        
        adoConsulta.Move CLng(tdgConsulta.SelBookmarks(intRegistro) - 1), 0
                        
        tdgConsulta.Refresh
                                
        If (tdgConsulta.Columns("EstadoOrden") = Estado_Orden_Anulada) Then
            MsgBox "La orden " & tdgConsulta.Columns("NumOrden") & " ha sido Anulada, no se puede procesar.", vbCritical, Me.Caption
        ElseIf (tdgConsulta.Columns("EstadoOrden") = Estado_Orden_Procesada) Then
            MsgBox "La orden " & tdgConsulta.Columns("NumOrden") & " ya ha sido Procesada.", vbCritical, Me.Caption
        Else
            adoRegistroAux.AddNew
            
            adoRegistroAux.Fields("CodFondo") = strCodFondo
            adoRegistroAux.Fields("CodAdministradora") = gstrCodAdministradora
            adoRegistroAux.Fields("NumOrden") = tdgConsulta.Columns("NumOrden")
        End If
    Next
    
    If adoRegistroAux.RecordCount > 0 Then
        Call XMLADORecordset(objFinanciamientoOrdenXML, "FinanciamientoOrden", "Orden", adoRegistroAux, strMsgError)
        strFinanciamientoOrdenXML = objFinanciamientoOrdenXML.xml
    
        adoComm.CommandText = "{ call up_FIProcFinanciamientoOrden('" & _
                                strCodFondo & "','" & gstrCodAdministradora & "','" & _
                                strFechaProceso & "','" & _
                                strFinanciamientoOrdenXML & "') }"
    
        adoComm.Execute adoComm.CommandText
        
        Me.MousePointer = vbDefault
        
        MsgBox Mensaje_Confirmacion_Exitoso, vbExclamation, gstrNombreEmpresa
    Else
        MsgBox "No se confirmó ninguna orden.", vbInformation
    End If
    Call Buscar


End Sub

Private Sub dtpFechaLiquidacionDesde_Click()

    If IsNull(dtpFechaLiquidacionDesde.Value) Then
        dtpFechaLiquidacionHasta.Value = Null
        dtpFechaOrdenDesde.Value = gdatFechaActual
        dtpFechaOrdenHasta.Value = dtpFechaOrdenDesde.Value
    Else
        dtpFechaLiquidacionDesde.Value = gdatFechaActual
        dtpFechaLiquidacionHasta.Value = dtpFechaLiquidacionDesde.Value
        dtpFechaOrdenDesde.Value = Null
        dtpFechaOrdenHasta.Value = dtpFechaOrdenDesde.Value
    End If
    
End Sub


Private Sub dtpFechaLiquidacionHasta_Click()

    If IsNull(dtpFechaLiquidacionHasta.Value) Then
        dtpFechaLiquidacionDesde.Value = Null
        dtpFechaOrdenDesde.Value = gdatFechaActual
        dtpFechaOrdenHasta.Value = dtpFechaOrdenDesde.Value
    Else
        dtpFechaLiquidacionDesde.Value = gdatFechaActual
        dtpFechaLiquidacionHasta.Value = dtpFechaLiquidacionDesde.Value
        dtpFechaOrdenDesde.Value = Null
        dtpFechaOrdenHasta.Value = dtpFechaOrdenDesde.Value
    End If
    
End Sub


Private Sub dtpFechaOrdenDesde_Click()

    If IsNull(dtpFechaOrdenDesde.Value) Then
        dtpFechaOrdenHasta.Value = Null
        dtpFechaLiquidacionDesde.Value = gdatFechaActual
        dtpFechaLiquidacionHasta.Value = dtpFechaLiquidacionDesde.Value
    Else
        dtpFechaOrdenDesde.Value = gdatFechaActual
        dtpFechaOrdenHasta.Value = dtpFechaOrdenDesde.Value
        dtpFechaLiquidacionDesde.Value = Null
        dtpFechaLiquidacionHasta.Value = dtpFechaLiquidacionDesde.Value
    End If
    
End Sub


Private Sub dtpFechaOrdenHasta_Click()

    If IsNull(dtpFechaOrdenHasta.Value) Then
        dtpFechaOrdenDesde.Value = Null
        dtpFechaLiquidacionDesde.Value = gdatFechaActual
        dtpFechaLiquidacionHasta.Value = dtpFechaLiquidacionDesde.Value
    Else
        dtpFechaOrdenDesde.Value = gdatFechaActual
        dtpFechaOrdenHasta.Value = dtpFechaOrdenDesde.Value
        dtpFechaLiquidacionDesde.Value = Null
        dtpFechaLiquidacionHasta.Value = dtpFechaLiquidacionDesde.Value
    End If
    
End Sub


Private Sub Form_Activate()

    frmMainMdi.stbMdi.Panels(3).Text = Me.Caption
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

Public Sub Buscar()

    Set adoConsulta = New ADODB.Recordset
     
    Dim strFechaOrdenDesde          As String, strFechaOrdenHasta        As String
    Dim strFechaLiquidacionDesde    As String, strFechaLiquidacionHasta  As String
    Dim datFechaSiguiente           As Date

    Me.MousePointer = vbHourglass
    
    If Not IsNull(dtpFechaOrdenDesde.Value) Or Not IsNull(dtpFechaOrdenHasta.Value) Then
        strFechaOrdenDesde = Convertyyyymmdd(dtpFechaOrdenDesde.Value)
        datFechaSiguiente = DateAdd("d", 1, dtpFechaOrdenHasta.Value)
        strFechaOrdenHasta = Convertyyyymmdd(datFechaSiguiente)
    End If
    
    If Not IsNull(dtpFechaLiquidacionDesde.Value) Or Not IsNull(dtpFechaLiquidacionHasta.Value) Then
        strFechaLiquidacionDesde = Convertyyyymmdd(dtpFechaLiquidacionDesde.Value)
        datFechaSiguiente = DateAdd("d", 1, dtpFechaLiquidacionHasta.Value)
        strFechaLiquidacionHasta = Convertyyyymmdd(datFechaSiguiente)
    End If
    
    strSQL = "SELECT NumOrden,FechaOrden,FechaLiquidacion,'',Nemotecnico,EstadoOrden,CodFile,CodAnalitica,TipoOrden,IOR.CodMoneda," & _
        "DescripOrden,CantOrden,ValorNominal,MontoInteres,ValorNominal as MontoTotalMFL1, CodSigno DescripMoneda " & _
        "FROM FinanciamientoOrden IOR JOIN AuxiliarParametro TON ON(TON.CodParametro=IOR.TipoOrden AND TON.CodTipoParametro='OPECAJ' AND TON.ValorParametro = 'I') " & _
        "JOIN Moneda MON ON(MON.CodMoneda=IOR.CodMoneda) " & _
        "WHERE CodAdministradora='" & gstrCodAdministradora & "' AND CodFondo='" & strCodFondo & "' "
        
    If strCodTipoInstrumento <> Valor_Caracter Then
        strSQL = strSQL & "AND CodFile='" & strCodTipoInstrumento & "' "
    End If

    If Not IsNull(dtpFechaOrdenDesde.Value) Or Not IsNull(dtpFechaOrdenHasta.Value) Then
        strSQL = strSQL & "AND (FechaOrden >='" & strFechaOrdenDesde & "' AND FechaOrden <'" & strFechaOrdenHasta & "') "
    End If
    
    If Not IsNull(dtpFechaLiquidacionDesde.Value) Or Not IsNull(dtpFechaLiquidacionHasta.Value) Then
        strSQL = strSQL & "AND (FechaLiquidacion >='" & strFechaLiquidacionDesde & "' AND FechaLiquidacion <'" & strFechaLiquidacionHasta & "') "
    End If
        
    If strCodTipoOrdenBusqueda <> Valor_Caracter Then
        strSQL = strSQL & " AND TipoOrden='" & strCodTipoOrdenBusqueda & "' "
    End If
    
    If strCodEstado <> Valor_Caracter Then
        strSQL = strSQL & " AND EstadoOrden='" & strCodEstado & "' "
    End If
    strSQL = strSQL & "ORDER BY NumOrden"
    
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

    Me.MousePointer = vbDefault
    
End Sub
Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Ordenes de Inversión"
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Text = "Operaciones de Inversión"
    
End Sub
Private Sub CargarListas()

    Dim adoRegistro As ADODB.Recordset
    Dim intRegistro As Integer
    
    '*** Fondos ***
    strSQL = "{ call up_ACSelDatosParametro(74,'" & gstrCodAdministradora & "','" & gstrCodFondoContable & "') }"
    CargarControlLista strSQL, cboFondo, arrFondo(), Valor_Caracter
        
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
    
    '*** Tipo de Instrumento ***
    strSQL = "SELECT CodFile CODIGO,DescripFile DESCRIP FROM InversionFile WHERE CodFile = '" & CodFile_Financiamiento_Prestamos & "' AND IndVigente='X' ORDER BY DescripFile"
    CargarControlLista strSQL, cboTipoInstrumento, arrTipoInstrumento(), Sel_Todos
    
    If cboTipoInstrumento.ListCount > 0 Then cboTipoInstrumento.ListIndex = 0
    
    '*** Estados de la Orden ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='ESTORD' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboEstado, arrEstado(), Sel_Todos
    
    intRegistro = ObtenerItemLista(arrEstado(), Estado_Orden_Enviada)
    If intRegistro >= 0 Then cboEstado.ListIndex = intRegistro
        
    '*** Tipo de Orden ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='OPECAJ' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboTipoOrden, arrTipoOrden(), Sel_Todos
    
    If cboTipoOrden.ListCount > 0 Then cboTipoOrden.ListIndex = 0
                    
End Sub
Private Sub InicializarValores()
    
    '*** Valores Iniciales ***
    strEstado = Reg_Defecto
    tabConfirmacionOrden.Tab = 0

    dtpFechaOrdenDesde.Value = gdatFechaActual
    dtpFechaOrdenHasta.Value = dtpFechaOrdenDesde.Value
    dtpFechaLiquidacionDesde.Value = Null
    dtpFechaLiquidacionHasta.Value = Null
    
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgConsulta.Columns(1).Width = tdgConsulta.Width * 0.01 * 9
    tdgConsulta.Columns(9).Width = tdgConsulta.Width * 0.01 * 10
    tdgConsulta.Columns(3).Width = tdgConsulta.Width * 0.01 * 3
    tdgConsulta.Columns(4).Width = tdgConsulta.Width * 0.01 * 7
    tdgConsulta.Columns(5).Width = tdgConsulta.Width * 0.01 * 45
    tdgConsulta.Columns(0).Width = tdgConsulta.Width * 0.01 * 4
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
        
End Sub
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

Public Sub Salir()

    Unload Me
    
End Sub

Public Sub Cancelar()

    cmdOpcion.Visible = True
    With tabConfirmacionOrden
        .TabEnabled(0) = True
        .Tab = 0
    End With
    Call Buscar
    
End Sub
Private Sub Form_Unload(Cancel As Integer)

    Set frmConfirmacionOrdenFinanciamiento = Nothing
    Call OcultarReportes
    frmMainMdi.stbMdi.Panels(3).Text = "Acción"
    
End Sub


Private Sub lblCantOrden_Change()

    Call FormatoMillarEtiqueta(lblCantOrden, Decimales_Monto)
    
End Sub

Private Sub lblComisionIgv_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblComisionIgv(Index), Decimales_Monto)
    
End Sub



Private Sub lblMontoTotal_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblMontoTotal(Index), Decimales_Monto)
    
End Sub

Private Sub lblPrecio_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblPrecio(Index), Decimales_Precio)
    
End Sub

Private Sub lblSubTotal_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblSubTotal(Index), Decimales_Monto)
    
End Sub

Private Sub lblTasa_Change()
    
    Call FormatoMillarEtiqueta(lblTasa, Decimales_Tasa)
    
End Sub

Private Sub lblTipoCambio_Change()

    Call FormatoMillarEtiqueta(lblTipoCambio, Decimales_TipoCambio)
    
End Sub

Private Sub lblTirNeta_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblTirNeta(Index), Decimales_Tasa)
    
End Sub

Private Sub lblValorNominal_Change()

    Call FormatoMillarEtiqueta(lblValorNominal, Decimales_Monto)
    
End Sub

Private Sub tabConfirmacionOrden_Click(PreviousTab As Integer)

    Select Case tabConfirmacionOrden.Tab
        Case 1
            If PreviousTab = 0 And strEstado = Reg_Consulta Then Call Accion(vModify)
            If strEstado = Reg_Defecto Then tabConfirmacionOrden.Tab = 0
    End Select
    
End Sub

Private Sub tdgConsulta_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)

    If ColIndex = 6 Then
        Call DarFormatoValor(Value, Decimales_Monto)
    End If
    
End Sub


Private Sub txtComisionAgente_Change(Index As Integer)

    Call FormatoCajaTexto(txtComisionAgente(Index), Decimales_Monto)
    
End Sub


Private Sub txtComisionAgente_KeyPress(Index As Integer, KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtComisionAgente(Index), Decimales_Monto)
    
    If KeyAscii = vbKeyReturn Then
        Call CalculoTotal(Index)
    End If
    
End Sub


Private Sub txtComisionBolsa_Change(Index As Integer)

    Call FormatoCajaTexto(txtComisionBolsa(Index), Decimales_Monto)
    
End Sub


Private Sub txtComisionBolsa_KeyPress(Index As Integer, KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtComisionBolsa(Index), Decimales_Monto)
    
    If KeyAscii = vbKeyReturn Then
        Call CalculoTotal(Index)
    End If
    
End Sub

Private Sub txtComisionCavali_Change(Index As Integer)

    Call FormatoCajaTexto(txtComisionCavali(Index), Decimales_Monto)
    
End Sub

Private Sub txtComisionCavali_KeyPress(Index As Integer, KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtComisionCavali(Index), Decimales_Monto)
    
    If KeyAscii = vbKeyReturn Then
        Call CalculoTotal(Index)
    End If
    
End Sub

Private Sub txtComisionConasev_Change(Index As Integer)

    Call FormatoCajaTexto(txtComisionConasev(Index), Decimales_Monto)
    
End Sub

Private Sub txtComisionConasev_KeyPress(Index As Integer, KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtComisionConasev(Index), Decimales_Monto)
    
    If KeyAscii = vbKeyReturn Then
        Call CalculoTotal(Index)
    End If
    
End Sub


Private Sub txtComisionEspecial_Change(Index As Integer)

Call FormatoCajaTexto(txtComisionEspecial(Index), Decimales_Monto)

End Sub

Private Sub txtComisionEspecial_KeyPress(Index As Integer, KeyAscii As Integer)
    
    Call ValidaCajaTexto(KeyAscii, "M", txtComisionEspecial(Index), Decimales_Monto)
    
    If KeyAscii = vbKeyReturn Then
        Call CalculoTotal(Index)
    End If

End Sub

Private Sub txtComisionFondo_Change(Index As Integer)

    Call FormatoCajaTexto(txtComisionFondo(Index), Decimales_Monto)
    
End Sub


Private Sub txtComisionFondo_KeyPress(Index As Integer, KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtComisionFondo(Index), Decimales_Monto)
    
    If KeyAscii = vbKeyReturn Then
        Call CalculoTotal(Index)
    End If
    
End Sub


Private Sub txtComisionFondoG_Change(Index As Integer)

    Call FormatoCajaTexto(txtComisionFondoG(Index), Decimales_Monto)
    
End Sub

Private Sub txtComisionFondoG_KeyPress(Index As Integer, KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtComisionFondoG(Index), Decimales_Monto)
    
    If KeyAscii = vbKeyReturn Then
        Call CalculoTotal(Index)
    End If
    
End Sub

Private Sub txtComisionGastoBancario_Change(Index As Integer)

Call FormatoCajaTexto(txtComisionGastoBancario(Index), Decimales_Monto)

End Sub

Private Sub txtComisionGastoBancario_KeyPress(Index As Integer, KeyAscii As Integer)
    
    Call ValidaCajaTexto(KeyAscii, "M", txtComisionGastoBancario(Index), Decimales_Monto)
    
    If KeyAscii = vbKeyReturn Then
        Call CalculoTotal(Index)
    End If
End Sub

Private Sub txtInteresCorrido_Change(Index As Integer)

    Call FormatoCajaTexto(txtInteresCorrido(Index), Decimales_Monto)
    
End Sub

Private Sub txtInteresCorrido_KeyPress(Index As Integer, KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtInteresCorrido(Index), Decimales_Monto)
    
    If KeyAscii = vbKeyReturn Then
        Call CalculoTotal(Index)
    End If
    
End Sub

Private Sub txtVacCorrido_Change(Index As Integer)

    Call FormatoCajaTexto(txtVacCorrido(Index), Decimales_Monto)
    
End Sub

Private Sub txtVacCorrido_KeyPress(Index As Integer, KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtVacCorrido(Index), Decimales_Monto)
    
    If KeyAscii = vbKeyReturn Then
        Call CalculoTotal(Index)
    End If
    
End Sub
Private Sub ConfiguraRecordsetAuxiliar()

    Set adoRegistroAux = New ADODB.Recordset

    With adoRegistroAux
       .CursorLocation = adUseClient
       .Fields.Append "CodFondo", adVarChar, 3
       .Fields.Append "CodAdministradora", adVarChar, 3
       .Fields.Append "NumOrden", adVarChar, 10
       .LockType = adLockBatchOptimistic
    End With

    adoRegistroAux.Open

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
