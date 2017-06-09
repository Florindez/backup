VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{830D5A06-9B70-4F8C-98B6-7A19C4E7760B}#1.0#0"; "TAMControls.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmOrdenDescuentoContratos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Desembolso"
   ClientHeight    =   9345
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   14505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9345
   ScaleWidth      =   14505
   ShowInTaskbar   =   0   'False
   Begin TAMControls2.ucBotonEdicion2 cmdOpcion 
      Height          =   735
      Left            =   630
      TabIndex        =   293
      Top             =   8520
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
   Begin TabDlg.SSTab tabRFCortoPlazo 
      Height          =   8355
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   14737
      _Version        =   393216
      Style           =   1
      Tab             =   2
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
      TabPicture(0)   =   "frmOrdenDescuentoContratos.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "tdgConsulta"
      Tab(0).Control(1)=   "fraCriterio"
      Tab(0).Control(2)=   "cmdOpcion1"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Datos Orden Inversión"
      TabPicture(1)   =   "frmOrdenDescuentoContratos.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraDatosAnexo"
      Tab(1).Control(1)=   "fraResumen"
      Tab(1).Control(2)=   "fraDatosBasicos"
      Tab(1).Control(3)=   "fraDatosTitulo"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Negociación"
      TabPicture(2)   =   "frmOrdenDescuentoContratos.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "lblDescrip(35)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "fraComisionMontoFL2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "fraComisionMontoFL1"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "fraDatosNegociacion"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "fraPosicion"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "fraComisiones"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).ControlCount=   6
      Begin VB.Frame fraComisiones 
         Height          =   1125
         Left            =   240
         TabIndex        =   6
         Top             =   5640
         Width           =   6735
         Begin VB.TextBox txtComisionAgente 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   4290
            Locked          =   -1  'True
            MaxLength       =   45
            TabIndex        =   8
            Top             =   150
            Width           =   2025
         End
         Begin VB.TextBox txtComisionBolsa 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   4290
            Locked          =   -1  'True
            MaxLength       =   45
            TabIndex        =   7
            Top             =   450
            Width           =   2025
         End
         Begin TAMControls.TAMTextBox txtPorcenIgv 
            Height          =   315
            Index           =   0
            Left            =   2760
            TabIndex        =   9
            Top             =   750
            Width           =   1335
            _ExtentX        =   2355
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
            Container       =   "frmOrdenDescuentoContratos.frx":0054
            Text            =   "0.000000"
            Decimales       =   6
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            EnterTab        =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   6
         End
         Begin TAMControls.TAMTextBox txtPorcenAgente 
            Height          =   285
            Index           =   0
            Left            =   2760
            TabIndex        =   10
            Top             =   150
            Width           =   1335
            _ExtentX        =   2355
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
            Container       =   "frmOrdenDescuentoContratos.frx":0070
            Text            =   "0.000000"
            Decimales       =   6
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            EnterTab        =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   6
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisión Transf,ITF,desem."
            BeginProperty Font 
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
            Index           =   94
            Left            =   240
            TabIndex        =   11
            Top             =   180
            Width           =   2355
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Otras comisiones"
            BeginProperty Font 
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
            Index           =   95
            Left            =   240
            TabIndex        =   12
            Top             =   495
            Width           =   1455
         End
         Begin VB.Label lblPorcenBolsa 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   0
            Left            =   2760
            TabIndex        =   18
            Tag             =   "0"
            ToolTipText     =   "Porcentaje de Comisión BVL"
            Top             =   450
            Width           =   1335
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
            Index           =   91
            Left            =   660
            TabIndex        =   17
            Top             =   810
            Width           =   960
         End
         Begin VB.Label lblComisionIgv 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   0
            Left            =   4290
            TabIndex        =   16
            Tag             =   "0.00"
            ToolTipText     =   "Monto de Comisión IGV"
            Top             =   780
            Width           =   2025
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisión SAB"
            BeginProperty Font 
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
            Left            =   480
            TabIndex        =   15
            Top             =   240
            Visible         =   0   'False
            Width           =   1170
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisión BVL"
            BeginProperty Font 
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
            Left            =   360
            TabIndex        =   14
            Top             =   585
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.Label lblDescrip 
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
            Index           =   34
            Left            =   240
            TabIndex        =   13
            Top             =   810
            Width           =   330
         End
      End
      Begin VB.Frame fraDatosTituloCancel 
         Caption         =   "Datos de la Orden"
         Height          =   3255
         Left            =   -1.00000e5
         TabIndex        =   253
         Top             =   2280
         Visible         =   0   'False
         Width           =   13935
         Begin VB.TextBox txtNumAnexoCancel 
            BackColor       =   &H8000000F&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2040
            Locked          =   -1  'True
            MaxLength       =   45
            TabIndex        =   263
            Top             =   360
            Width           =   2010
         End
         Begin VB.TextBox txtNumDocDsctoCancel 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   10890
            Locked          =   -1  'True
            MaxLength       =   45
            TabIndex        =   262
            Top             =   360
            Width           =   2490
         End
         Begin VB.TextBox txtObservacionCancel 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   525
            Left            =   2040
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   261
            Top             =   2550
            Width           =   11370
         End
         Begin VB.TextBox txtNemonicoCancel 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   7800
            MaxLength       =   15
            TabIndex        =   260
            Top             =   2055
            Width           =   2655
         End
         Begin VB.TextBox txtDescripOrdenCancel 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   2040
            MaxLength       =   45
            TabIndex        =   259
            Top             =   2055
            Width           =   4170
         End
         Begin VB.ComboBox cboResponsablePagoCancel 
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   315
            Left            =   10800
            Style           =   2  'Dropdown List
            TabIndex        =   258
            Top             =   960
            Width           =   2625
         End
         Begin VB.ComboBox cboViaCobranza 
            Height          =   315
            Left            =   10800
            Style           =   2  'Dropdown List
            TabIndex        =   257
            Top             =   1320
            Width           =   2625
         End
         Begin VB.TextBox txtMonedaOrig 
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   6000
            Locked          =   -1  'True
            MaxLength       =   45
            TabIndex        =   256
            Top             =   1320
            Width           =   2610
         End
         Begin VB.CheckBox chkPrelacion 
            Caption         =   "Con Prelación"
            Enabled         =   0   'False
            Height          =   255
            Left            =   8280
            TabIndex        =   255
            Top             =   1680
            Width           =   1455
         End
         Begin VB.TextBox txtNumContratoCancel 
            BackColor       =   &H8000000F&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   6240
            Locked          =   -1  'True
            MaxLength       =   45
            TabIndex        =   254
            Top             =   360
            Width           =   2010
         End
         Begin MSComCtl2.DTPicker dtpFechaOrdenCancel 
            Height          =   315
            Left            =   2040
            TabIndex        =   264
            Top             =   960
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
            Format          =   48234497
            CurrentDate     =   38776
         End
         Begin MSComCtl2.DTPicker dtpFechaLiquidacionCancel 
            Height          =   315
            Left            =   6000
            TabIndex        =   265
            Top             =   960
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
            Format          =   48234497
            CurrentDate     =   38776
         End
         Begin TAMControls.TAMTextBox txtMontoTotalRecibido 
            Height          =   315
            Left            =   2040
            TabIndex        =   266
            Top             =   1680
            Width           =   2025
            _ExtentX        =   3572
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
            Container       =   "frmOrdenDescuentoContratos.frx":008C
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
         Begin TAMControls.TAMTextBox txtDeudaTotal 
            Height          =   315
            Left            =   2040
            TabIndex        =   267
            Top             =   1320
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   556
            BackColor       =   14737632
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
            Container       =   "frmOrdenDescuentoContratos.frx":00A8
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
         Begin TAMControls.TAMTextBox txtSaldoDeuda 
            Height          =   315
            Left            =   11400
            TabIndex        =   268
            Top             =   1680
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   556
            BackColor       =   14737632
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
            Container       =   "frmOrdenDescuentoContratos.frx":00C4
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
         Begin TAMControls.TAMTextBox txtMontoTotalCancel 
            Height          =   315
            Left            =   6000
            TabIndex        =   269
            Top             =   1680
            Width           =   2025
            _ExtentX        =   3572
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
            Container       =   "frmOrdenDescuentoContratos.frx":00E0
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
            Caption         =   "Número de Anexo"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   118
            Left            =   360
            TabIndex        =   284
            Top             =   405
            Width           =   1275
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Nro.  Documento descontado"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   119
            Left            =   8400
            TabIndex        =   283
            Top             =   405
            Width           =   2100
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000010&
            BorderStyle     =   6  'Inside Solid
            Index           =   1
            X1              =   0
            X2              =   13440
            Y1              =   840
            Y2              =   840
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Instrucciones"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   120
            Left            =   360
            TabIndex        =   282
            Top             =   2580
            Width           =   945
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Nemónico"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   121
            Left            =   6720
            TabIndex        =   281
            Top             =   2115
            Width           =   720
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Moneda Origen"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   122
            Left            =   4560
            TabIndex        =   280
            Top             =   1365
            Width           =   1095
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Descripción"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   123
            Left            =   360
            TabIndex        =   279
            Top             =   2160
            Width           =   840
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Liquidación"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   124
            Left            =   4860
            TabIndex        =   278
            Top             =   1005
            Width           =   810
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Orden"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   125
            Left            =   360
            TabIndex        =   277
            Top             =   1005
            Width           =   435
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Pagador"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   127
            Left            =   9720
            TabIndex        =   276
            Top             =   1000
            Width           =   600
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Via de Cobranza"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   128
            Left            =   9240
            TabIndex        =   275
            Top             =   1380
            Width           =   1170
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto Pago"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   135
            Left            =   4560
            TabIndex        =   274
            Top             =   1725
            Width           =   870
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Deuda a la fecha"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   156
            Left            =   360
            TabIndex        =   273
            Top             =   1360
            Width           =   1230
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Deuda"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   157
            Left            =   10200
            TabIndex        =   272
            Top             =   1725
            Width           =   930
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto recibido"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   139
            Left            =   360
            TabIndex        =   271
            Top             =   1750
            Width           =   1050
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Número de Contrato"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   162
            Left            =   4440
            TabIndex        =   270
            Top             =   405
            Width           =   1425
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
         Height          =   2115
         Left            =   -74760
         TabIndex        =   230
         Top             =   4275
         Width           =   13935
         Begin VB.TextBox txtNumAnexo 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   286
            Top             =   400
            Width           =   1455
         End
         Begin VB.TextBox txtDescripOrden 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1560
            MaxLength       =   45
            TabIndex        =   236
            Top             =   1620
            Width           =   4290
         End
         Begin VB.ComboBox cboMoneda 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   235
            Top             =   1225
            Width           =   2400
         End
         Begin VB.TextBox txtNemonico 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   5400
            MaxLength       =   15
            TabIndex        =   234
            Top             =   1225
            Width           =   1935
         End
         Begin VB.TextBox txtObservacion 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   7560
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   233
            Top             =   1620
            Width           =   6090
         End
         Begin VB.CheckBox chkDiasAdicional 
            Caption         =   "Adicionar días protesto"
            BeginProperty Font 
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
            Left            =   12480
            TabIndex        =   232
            Top             =   720
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.ComboBox cboResponsablePago 
            Height          =   315
            Left            =   11400
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   231
            Top             =   1225
            Width           =   2265
         End
         Begin MSComCtl2.DTPicker dtpFechaOrden 
            Height          =   315
            Left            =   1560
            TabIndex        =   237
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
            Format          =   48234497
            CurrentDate     =   38776
         End
         Begin MSComCtl2.DTPicker dtpFechaLiquidacion 
            Height          =   315
            Left            =   4560
            TabIndex        =   238
            Top             =   810
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
            Format          =   48234497
            CurrentDate     =   38776
         End
         Begin MSComCtl2.UpDown updDiasPlazo 
            Height          =   315
            Left            =   13395
            TabIndex        =   239
            Top             =   1200
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   556
            _Version        =   393216
            OrigLeft        =   13395
            OrigTop         =   1200
            OrigRight       =   13650
            OrigBottom      =   1515
            Max             =   360
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.DTPicker dtpFechaVencimiento 
            Height          =   315
            Left            =   8280
            TabIndex        =   240
            Top             =   825
            Width           =   1575
            _ExtentX        =   2778
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
            Format          =   48234497
            CurrentDate     =   38776
         End
         Begin MSComCtl2.DTPicker dtpFechaPago 
            Height          =   315
            Left            =   8280
            TabIndex        =   241
            Top             =   1225
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
            Format          =   48234497
            CurrentDate     =   38776
         End
         Begin TAMControls.TAMTextBox txtDiasPlazo 
            Height          =   315
            Left            =   11400
            TabIndex        =   242
            Top             =   840
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            BackColor       =   16777215
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
            Alignment       =   1
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "frmOrdenDescuentoContratos.frx":00FC
            Text            =   "0"
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            EnterTab        =   -1  'True
            Apariencia      =   1
            Borde           =   1
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "N° Anexo"
            BeginProperty Font 
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
            TabIndex        =   287
            Top             =   450
            Width           =   795
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
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
            Left            =   360
            TabIndex        =   252
            Top             =   855
            Width           =   525
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Liquidación"
            BeginProperty Font 
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
            Left            =   3300
            TabIndex        =   251
            Top             =   855
            Width           =   975
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
            Left            =   360
            TabIndex        =   250
            Top             =   1620
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
            Left            =   10080
            TabIndex        =   249
            Top             =   840
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
            Left            =   360
            TabIndex        =   248
            Top             =   1245
            Width           =   690
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
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
            Left            =   4200
            TabIndex        =   247
            Top             =   1245
            Width           =   840
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Pago"
            BeginProperty Font 
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
            Index           =   86
            Left            =   7560
            TabIndex        =   246
            Top             =   1260
            Width           =   450
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
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
            Left            =   6120
            TabIndex        =   245
            Top             =   1680
            Width           =   1155
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
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
            Left            =   6240
            TabIndex        =   244
            Top             =   855
            Width           =   1965
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Pagador"
            BeginProperty Font 
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
            Index           =   138
            Left            =   10440
            TabIndex        =   243
            Top             =   1260
            Width           =   720
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
         Height          =   2625
         Left            =   -74760
         TabIndex        =   186
         Top             =   390
         Width           =   13935
         Begin VB.CheckBox chkTitulo 
            Height          =   255
            Left            =   13560
            TabIndex        =   200
            ToolTipText     =   "Seleccionar Título"
            Top             =   360
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.ComboBox cboEmisor 
            Height          =   315
            Left            =   9300
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   199
            Top             =   360
            Width           =   4185
         End
         Begin VB.ComboBox cboClaseInstrumento 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2370
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   198
            Top             =   1095
            Width           =   4185
         End
         Begin VB.ComboBox cboTitulo 
            Height          =   315
            Left            =   9300
            Style           =   2  'Dropdown List
            TabIndex        =   197
            Top             =   360
            Width           =   4185
         End
         Begin VB.ComboBox cboTipoOrden 
            Height          =   315
            Left            =   2370
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   196
            Top             =   1830
            Width           =   4185
         End
         Begin VB.ComboBox cboTipoInstrumentoOrden 
            Height          =   315
            Left            =   2370
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   195
            Top             =   720
            Width           =   4185
         End
         Begin VB.ComboBox cboFondoOrden 
            Height          =   315
            Left            =   2370
            Style           =   2  'Dropdown List
            TabIndex        =   194
            Top             =   360
            Width           =   4185
         End
         Begin VB.ComboBox cboSubClaseInstrumento 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2370
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   193
            Top             =   1455
            Width           =   4185
         End
         Begin VB.ComboBox cboObligado 
            Height          =   315
            Left            =   9300
            Style           =   2  'Dropdown List
            TabIndex        =   192
            Top             =   720
            Visible         =   0   'False
            Width           =   4185
         End
         Begin VB.ComboBox cboGestor 
            Height          =   315
            Left            =   9300
            Style           =   2  'Dropdown List
            TabIndex        =   191
            Top             =   1080
            Width           =   4185
         End
         Begin VB.ComboBox cboOrigen 
            Height          =   315
            Left            =   9315
            Style           =   2  'Dropdown List
            TabIndex        =   190
            Top             =   1800
            Width           =   4185
         End
         Begin VB.ComboBox cboOperacion 
            Height          =   315
            Left            =   9300
            Style           =   2  'Dropdown List
            TabIndex        =   189
            Top             =   1440
            Width           =   4185
         End
         Begin VB.ComboBox cboLineaCliente 
            Height          =   315
            Left            =   9315
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   188
            Top             =   2160
            Width           =   4185
         End
         Begin VB.TextBox txtNumOperacionOrig 
            Height          =   285
            Left            =   2370
            Locked          =   -1  'True
            TabIndex        =   187
            Top             =   2160
            Width           =   2385
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
            TabIndex        =   212
            Top             =   1485
            Width           =   810
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
            Left            =   7170
            TabIndex        =   211
            Top             =   405
            Width           =   570
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Orden de"
            BeginProperty Font 
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
            TabIndex        =   210
            Top             =   1845
            Width           =   795
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
            TabIndex        =   209
            Top             =   750
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
            Left            =   360
            TabIndex        =   208
            Top             =   375
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
            Left            =   360
            TabIndex        =   207
            Top             =   1110
            Width           =   480
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Obligado"
            BeginProperty Font 
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
            Left            =   7170
            TabIndex        =   206
            Top             =   765
            Visible         =   0   'False
            Width           =   765
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
            Left            =   7170
            TabIndex        =   205
            Top             =   1125
            Width           =   570
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Mercado Negociación"
            BeginProperty Font 
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
            Left            =   7170
            TabIndex        =   204
            Top             =   1875
            Width           =   1860
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Liquidación Operación"
            BeginProperty Font 
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
            Left            =   7170
            TabIndex        =   203
            Top             =   1500
            Width           =   1905
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
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
            Left            =   7170
            TabIndex        =   202
            Top             =   2220
            Width           =   1500
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Nro. Solicitud"
            BeginProperty Font 
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
            Index           =   126
            Left            =   360
            TabIndex        =   201
            Top             =   2205
            Width           =   1170
         End
      End
      Begin VB.Frame fraPosicion 
         Caption         =   "Datos Posición"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   9480
         TabIndex        =   174
         Top             =   390
         Width           =   4695
         Begin VB.Label lblMoneda 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2160
            TabIndex        =   184
            Tag             =   "0.00"
            ToolTipText     =   "Moneda del Título"
            Top             =   1800
            Width           =   2025
         End
         Begin VB.Label lblStockNominal 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2160
            TabIndex        =   183
            Tag             =   "0.00"
            Top             =   1440
            Width           =   2025
         End
         Begin VB.Label lblBaseTasaCupon 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2160
            TabIndex        =   182
            Tag             =   "0.00"
            Top             =   1080
            Width           =   2025
         End
         Begin VB.Label lblClasificacion 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2160
            TabIndex        =   181
            Tag             =   "0.00"
            Top             =   720
            Width           =   2025
         End
         Begin VB.Label lblFechaCupon 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2160
            TabIndex        =   180
            Tag             =   "0.00"
            Top             =   360
            Width           =   2025
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
            Index           =   49
            Left            =   480
            TabIndex        =   179
            Top             =   1815
            Width           =   690
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Stock Nominal"
            BeginProperty Font 
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
            Left            =   480
            TabIndex        =   178
            Top             =   1455
            Width           =   1245
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Base - Tasa %"
            BeginProperty Font 
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
            Left            =   480
            TabIndex        =   177
            Top             =   1095
            Width           =   1230
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Clasificación"
            BeginProperty Font 
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
            Left            =   480
            TabIndex        =   176
            Top             =   735
            Width           =   1080
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Cupón Vigente"
            BeginProperty Font 
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
            Left            =   480
            TabIndex        =   175
            Top             =   375
            Width           =   1245
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
         Height          =   3165
         Left            =   180
         TabIndex        =   143
         Top             =   390
         Width           =   9135
         Begin VB.TextBox txtCantidad 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   2160
            MaxLength       =   45
            TabIndex        =   152
            Top             =   2760
            Width           =   1900
         End
         Begin VB.TextBox txtTipoCambio 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   6570
            MaxLength       =   45
            TabIndex        =   151
            Text            =   "0.00"
            Top             =   2460
            Visible         =   0   'False
            Width           =   1830
         End
         Begin VB.ComboBox cboTipoTasa 
            Height          =   315
            Left            =   2160
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   150
            Top             =   570
            Width           =   1900
         End
         Begin VB.ComboBox cboBaseAnual 
            Height          =   315
            Left            =   2160
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   149
            Top             =   975
            Width           =   1900
         End
         Begin VB.TextBox txtTasa 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   2160
            Locked          =   -1  'True
            MaxLength       =   45
            TabIndex        =   148
            Top             =   210
            Width           =   1900
         End
         Begin VB.ComboBox cboNegociacion 
            Height          =   315
            Left            =   6570
            Style           =   2  'Dropdown List
            TabIndex        =   147
            Top             =   300
            Width           =   2295
         End
         Begin VB.ComboBox cboConceptoCosto 
            Height          =   315
            Left            =   6570
            Style           =   2  'Dropdown List
            TabIndex        =   146
            Top             =   690
            Width           =   2295
         End
         Begin VB.TextBox txtValorNominalDcto 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   2160
            Locked          =   -1  'True
            MaxLength       =   45
            TabIndex        =   145
            Text            =   "0.00"
            Top             =   2400
            Width           =   1900
         End
         Begin VB.ComboBox cboCobroInteres 
            Height          =   315
            Left            =   2160
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   144
            Top             =   1335
            Width           =   1900
         End
         Begin TAMControls.TAMTextBox txtPorcenDctoValorNominal 
            Height          =   315
            Left            =   2160
            TabIndex        =   153
            Top             =   2040
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
            Container       =   "frmOrdenDescuentoContratos.frx":0118
            Text            =   "0.000000"
            Decimales       =   6
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            EnterTab        =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   6
         End
         Begin TAMControls.TAMTextBox txtValorNominal 
            Height          =   315
            Left            =   2160
            TabIndex        =   154
            Top             =   1680
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
            Container       =   "frmOrdenDescuentoContratos.frx":0134
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
         Begin VB.Label lblFechaLiquidacion 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6570
            TabIndex        =   173
            Tag             =   "0.00"
            ToolTipText     =   "Fecha de Liquidación"
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad Facial"
            BeginProperty Font 
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
            Left            =   360
            TabIndex        =   172
            Top             =   2760
            Width           =   1335
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Cambio"
            BeginProperty Font 
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
            Left            =   4830
            TabIndex        =   171
            Top             =   2490
            Visible         =   0   'False
            Width           =   1065
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
            Left            =   360
            TabIndex        =   170
            Top             =   615
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
            Left            =   360
            TabIndex        =   169
            Top             =   990
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
            Left            =   360
            TabIndex        =   168
            Top             =   300
            Width           =   1005
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000015&
            X1              =   4470
            X2              =   4470
            Y1              =   240
            Y2              =   3000
         End
         Begin VB.Label lblFechaEmision 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6570
            TabIndex        =   167
            Tag             =   "0.00"
            ToolTipText     =   "Fecha Emisión"
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label lblFechaVencimiento 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6570
            TabIndex        =   166
            Tag             =   "0.00"
            ToolTipText     =   "Fecha de Vencimiento del Título de la Orden"
            Top             =   1800
            Width           =   1815
         End
         Begin VB.Label lblDiasPlazo 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6570
            TabIndex        =   165
            Tag             =   "0.00"
            ToolTipText     =   "Días de Plazo del Título de la Orden"
            Top             =   2160
            Width           =   1815
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
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
            Index           =   12
            Left            =   360
            TabIndex        =   164
            Top             =   1680
            Width           =   1185
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Liquidación"
            BeginProperty Font 
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
            Left            =   4800
            TabIndex        =   163
            Top             =   1095
            Width           =   975
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Emisión"
            BeginProperty Font 
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
            Left            =   4800
            TabIndex        =   162
            Top             =   1455
            Width           =   645
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Vencimiento"
            BeginProperty Font 
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
            Left            =   4800
            TabIndex        =   161
            Top             =   1815
            Width           =   1050
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Plazo (Días)"
            BeginProperty Font 
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
            Index           =   83
            Left            =   4800
            TabIndex        =   160
            Top             =   2205
            Width           =   1050
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Mecanismo"
            BeginProperty Font 
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
            Left            =   4770
            TabIndex        =   159
            Top             =   390
            Width           =   960
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Concepto Costo Neg."
            BeginProperty Font 
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
            Index           =   82
            Left            =   4770
            TabIndex        =   158
            Top             =   720
            Width           =   1830
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "% V.Nominal Dcto."
            BeginProperty Font 
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
            Index           =   89
            Left            =   360
            TabIndex        =   157
            Top             =   2070
            Width           =   1590
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Valor Nominal Dcto."
            BeginProperty Font 
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
            Index           =   90
            Left            =   360
            TabIndex        =   156
            Top             =   2430
            Width           =   1710
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
            Left            =   360
            TabIndex        =   155
            Top             =   1335
            Width           =   1650
         End
      End
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
         Height          =   4605
         Left            =   240
         TabIndex        =   100
         Top             =   3600
         Width           =   6735
         Begin VB.CommandButton cmdCalculo 
            Caption         =   "#"
            Height          =   375
            Left            =   510
            TabIndex        =   111
            ToolTipText     =   "Calcular Valor al Vencimiento y TIRs de la orden"
            Top             =   5550
            Width           =   375
         End
         Begin VB.TextBox txtInteresCorrido 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   4290
            MaxLength       =   45
            TabIndex        =   110
            Top             =   3225
            Width           =   2025
         End
         Begin VB.TextBox txtPrecioUnitario 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   240
            MaxLength       =   45
            TabIndex        =   109
            Top             =   5400
            Width           =   1340
         End
         Begin VB.TextBox txtTirNeta 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   2610
            MaxLength       =   45
            TabIndex        =   108
            Top             =   5580
            Width           =   1365
         End
         Begin VB.TextBox txtComisionConasev 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   -1830
            MaxLength       =   45
            TabIndex        =   106
            Top             =   3360
            Visible         =   0   'False
            Width           =   2025
         End
         Begin VB.TextBox txtComisionFondo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   -1830
            MaxLength       =   45
            TabIndex        =   105
            Top             =   3600
            Visible         =   0   'False
            Width           =   2025
         End
         Begin VB.TextBox txtComisionCavali 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   -1830
            MaxLength       =   45
            TabIndex        =   104
            Top             =   3120
            Visible         =   0   'False
            Width           =   2025
         End
         Begin VB.TextBox txtIntAdicional 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   4260
            Locked          =   -1  'True
            MaxLength       =   45
            TabIndex        =   103
            Top             =   1200
            Width           =   2025
         End
         Begin VB.CheckBox chkAplicar 
            Caption         =   "Aplicar Comisiones"
            BeginProperty Font 
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
            Index           =   0
            Left            =   240
            TabIndex        =   102
            Top             =   1830
            Width           =   2115
         End
         Begin VB.CheckBox chkInteresCorrido 
            Caption         =   "Interés Corrido"
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
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   101
            Top             =   3240
            Width           =   1935
         End
         Begin TAMControls.TAMTextBox txtTirBruta1 
            Height          =   315
            Left            =   960
            TabIndex        =   107
            Top             =   4200
            Width           =   1335
            _ExtentX        =   2355
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
            Container       =   "frmOrdenDescuentoContratos.frx":0150
            Text            =   "0.000000"
            Decimales       =   6
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            EnterTab        =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   6
         End
         Begin TAMControls.TAMTextBox txtMontoVencimiento1 
            Height          =   315
            Left            =   4290
            TabIndex        =   112
            Top             =   4200
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   556
            BackColor       =   16777215
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
            Alignment       =   1
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   -2147483640
            Container       =   "frmOrdenDescuentoContratos.frx":016C
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
         Begin TAMControls.TAMTextBox txtPrecioUnitario1 
            Height          =   315
            Left            =   1320
            TabIndex        =   113
            Top             =   330
            Width           =   1335
            _ExtentX        =   2355
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
            Container       =   "frmOrdenDescuentoContratos.frx":0188
            Text            =   "0.000000"
            Decimales       =   6
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            EnterTab        =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   6
         End
         Begin TAMControls.TAMTextBox txtPorcenIgvInt 
            Height          =   315
            Index           =   0
            Left            =   2760
            TabIndex        =   114
            Top             =   1560
            Width           =   1335
            _ExtentX        =   2355
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
            Container       =   "frmOrdenDescuentoContratos.frx":01A4
            Text            =   "0.000000"
            Decimales       =   6
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            EnterTab        =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   6
         End
         Begin TAMControls.TAMTextBox txtImptoInteresCorrido 
            Height          =   315
            Index           =   0
            Left            =   2760
            TabIndex        =   115
            Top             =   3240
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
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
            Container       =   "frmOrdenDescuentoContratos.frx":01C0
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
         Begin TAMControls.TAMTextBox txtImptoInteres 
            Height          =   315
            Index           =   0
            Left            =   2760
            TabIndex        =   116
            Top             =   840
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
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
            Container       =   "frmOrdenDescuentoContratos.frx":01DC
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
         Begin TAMControls.TAMTextBox txtImptoInteresAdic 
            Height          =   315
            Index           =   0
            Left            =   2760
            TabIndex        =   117
            Top             =   1200
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
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
            Container       =   "frmOrdenDescuentoContratos.frx":01F8
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
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   2
            X1              =   240
            X2              =   6300
            Y1              =   3960
            Y2              =   3960
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "TIR Neta"
            BeginProperty Font 
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
            Left            =   3000
            TabIndex        =   142
            Top             =   3975
            Width           =   795
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "TIR Bruta"
            BeginProperty Font 
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
            Left            =   1320
            TabIndex        =   141
            Top             =   3975
            Width           =   840
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Valor al Vencimiento"
            BeginProperty Font 
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
            Left            =   4620
            TabIndex        =   140
            Top             =   3975
            Width           =   1755
         End
         Begin VB.Label lblMontoTotal 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   0
            Left            =   4290
            TabIndex        =   139
            Tag             =   "0.00"
            ToolTipText     =   "Monto Total de la Orden"
            Top             =   3585
            Width           =   2025
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   1
            X1              =   360
            X2              =   6300
            Y1              =   3120
            Y2              =   3120
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   0
            X1              =   2760
            X2              =   6360
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Label lblSubTotal 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   0
            Left            =   4260
            TabIndex        =   138
            Tag             =   "0.00"
            ToolTipText     =   "SubTotal de la Orden"
            Top             =   360
            Width           =   2025
         End
         Begin VB.Label lblTirNeta 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2715
            TabIndex        =   137
            Tag             =   "0.00"
            Top             =   4200
            Width           =   1335
         End
         Begin VB.Label lblTirBruta 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1080
            TabIndex        =   136
            Tag             =   "0.00"
            Top             =   5580
            Width           =   1335
         End
         Begin VB.Label lblMontoVencimiento 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4320
            TabIndex        =   135
            Tag             =   "0.00"
            Top             =   5610
            Width           =   2025
         End
         Begin VB.Label lblDescripMoneda 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
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
            Index           =   0
            Left            =   4455
            TabIndex        =   134
            Top             =   120
            Width           =   1665
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
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
            Index           =   36
            Left            =   240
            TabIndex        =   133
            Top             =   3600
            Width           =   1035
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "SubTotal"
            BeginProperty Font 
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
            Left            =   3240
            TabIndex        =   132
            Top             =   375
            Width           =   780
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Precio (%)"
            BeginProperty Font 
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
            Left            =   255
            TabIndex        =   131
            Top             =   375
            Width           =   870
         End
         Begin VB.Label lblIntAdelantado 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   0
            Left            =   4260
            TabIndex        =   130
            Tag             =   "0.00"
            ToolTipText     =   "SubTotal de la Orden"
            Top             =   840
            Width           =   2025
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
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
            Index           =   92
            Left            =   240
            TabIndex        =   129
            Top             =   840
            Width           =   585
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisión Conasev"
            BeginProperty Font 
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
            Left            =   -1080
            TabIndex        =   128
            Top             =   4080
            Visible         =   0   'False
            Width           =   1545
         End
         Begin VB.Label lblPorcenConasev 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   0
            Left            =   -1095
            TabIndex        =   127
            Tag             =   "0"
            ToolTipText     =   "Porcentaje de Comisión Conasev"
            Top             =   2640
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisión Fondo Garantía"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   32
            Left            =   -1560
            TabIndex        =   126
            Top             =   4320
            Visible         =   0   'False
            Width           =   1800
         End
         Begin VB.Label lblPorcenFondo 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   0
            Left            =   -1095
            TabIndex        =   125
            Tag             =   "0"
            ToolTipText     =   "Porcentaje de Comisión Fondo Liquidación"
            Top             =   2880
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisión Cavali"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   27
            Left            =   -840
            TabIndex        =   124
            Top             =   3840
            Visible         =   0   'False
            Width           =   1110
         End
         Begin VB.Label lblPorcenCavali 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   0
            Left            =   -1080
            TabIndex        =   123
            Tag             =   "0"
            ToolTipText     =   "Porcentaje de Comisión Cavali"
            Top             =   2400
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Interés Prov.Protesto"
            BeginProperty Font 
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
            Index           =   93
            Left            =   240
            TabIndex        =   122
            Top             =   1185
            Width           =   1800
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "IGV Intereses"
            BeginProperty Font 
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
            Left            =   240
            TabIndex        =   121
            Top             =   1560
            Width           =   1170
         End
         Begin VB.Label lblComisionIgvInt 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   0
            Left            =   4290
            TabIndex        =   120
            Tag             =   "0.00"
            ToolTipText     =   "Monto de Comisión IGV"
            Top             =   1560
            Width           =   2025
         End
         Begin VB.Label lblDiasAdic 
            AutoSize        =   -1  'True
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   2160
            TabIndex        =   119
            Top             =   1200
            Width           =   1005
         End
         Begin VB.Label lblFechaVencimientoAdic 
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   2760
            TabIndex        =   118
            Tag             =   "0.00"
            ToolTipText     =   "Fecha de Vencimiento del Título de la Orden"
            Top             =   1200
            Visible         =   0   'False
            Width           =   1335
         End
      End
      Begin VB.Frame fraResumen 
         Caption         =   "Resumen Negociación"
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
         Left            =   -74760
         TabIndex        =   65
         Top             =   6420
         Width           =   13905
         Begin VB.Line Line3 
            BorderColor     =   &H80000015&
            X1              =   9720
            X2              =   9720
            Y1              =   240
            Y2              =   1680
         End
         Begin VB.Label lblDescripMonedaResumen 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
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
            Index           =   1
            Left            =   7410
            TabIndex        =   99
            Top             =   600
            Width           =   1665
         End
         Begin VB.Label lblDescripMonedaResumen 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
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
            Index           =   0
            Left            =   2580
            TabIndex        =   98
            Top             =   600
            Width           =   1665
         End
         Begin VB.Label lblVencimientoResumen 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7320
            TabIndex        =   97
            Tag             =   "0.00"
            Top             =   240
            Width           =   2025
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Vencimiento"
            BeginProperty Font 
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
            Index           =   78
            Left            =   5280
            TabIndex        =   96
            Top             =   255
            Width           =   1050
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad Facial"
            BeginProperty Font 
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
            Index           =   77
            Left            =   360
            TabIndex        =   95
            Top             =   270
            Width           =   1335
         End
         Begin VB.Label lblCantidadResumen 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2400
            TabIndex        =   94
            Tag             =   "0.00"
            Top             =   240
            Width           =   2025
         End
         Begin VB.Line Line4 
            BorderColor     =   &H80000015&
            X1              =   4800
            X2              =   4800
            Y1              =   240
            Y2              =   1620
         End
         Begin VB.Label lblTirNetaResumen 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   11400
            TabIndex        =   93
            Tag             =   "0.00"
            Top             =   1320
            Width           =   2025
         End
         Begin VB.Label lblTirBrutaResumen 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   11400
            TabIndex        =   92
            Tag             =   "0.00"
            Top             =   960
            Width           =   2025
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tir Neta"
            BeginProperty Font 
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
            Index           =   75
            Left            =   10200
            TabIndex        =   91
            Top             =   1320
            Width           =   705
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tir Bruta"
            BeginProperty Font 
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
            Index           =   74
            Left            =   10200
            TabIndex        =   90
            Top             =   960
            Width           =   750
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Plazo"
            BeginProperty Font 
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
            Index           =   71
            Left            =   5280
            TabIndex        =   89
            Top             =   600
            Width           =   480
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Contado"
            BeginProperty Font 
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
            Index           =   70
            Left            =   360
            TabIndex        =   88
            Top             =   600
            Width           =   720
         End
         Begin VB.Label lblTotalResumen 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   1
            Left            =   7320
            TabIndex        =   87
            Tag             =   "0.00"
            Top             =   1305
            Width           =   2025
         End
         Begin VB.Label lblInteresesResumen 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   1
            Left            =   7230
            TabIndex        =   86
            Tag             =   "0.00"
            Top             =   3645
            Width           =   2025
         End
         Begin VB.Label lblComisionesResumen 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   1
            Left            =   7230
            TabIndex        =   85
            Tag             =   "0.00"
            Top             =   3315
            Width           =   2025
         End
         Begin VB.Label lblSubTotalResumen 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   1
            Left            =   7230
            TabIndex        =   84
            Tag             =   "0.00"
            Top             =   2985
            Width           =   2025
         End
         Begin VB.Label lblPrecioResumen 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   1
            Left            =   7320
            TabIndex        =   83
            Tag             =   "0.00"
            Top             =   915
            Width           =   2025
         End
         Begin VB.Label lblTotalResumen 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   0
            Left            =   2400
            TabIndex        =   82
            Tag             =   "0.00"
            Top             =   1305
            Width           =   2025
         End
         Begin VB.Label lblInteresesResumen 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   0
            Left            =   2310
            TabIndex        =   81
            Tag             =   "0.00"
            Top             =   3645
            Width           =   2025
         End
         Begin VB.Label lblComisionesResumen 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   0
            Left            =   2310
            TabIndex        =   80
            Tag             =   "0.00"
            Top             =   3315
            Width           =   2025
         End
         Begin VB.Label lblSubTotalResumen 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   0
            Left            =   2310
            TabIndex        =   79
            Tag             =   "0.00"
            Top             =   2985
            Width           =   2025
         End
         Begin VB.Label lblPrecioResumen 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   0
            Left            =   2400
            TabIndex        =   78
            Tag             =   "0.00"
            Top             =   915
            Width           =   2025
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
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
            Index           =   69
            Left            =   5250
            TabIndex        =   77
            Top             =   1320
            Width           =   1035
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Intereses Corridos"
            Height          =   195
            Index           =   68
            Left            =   5190
            TabIndex        =   76
            Top             =   3660
            Width           =   1260
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisiones"
            Height          =   195
            Index           =   67
            Left            =   5190
            TabIndex        =   75
            Top             =   3330
            Width           =   795
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "SubTotal"
            Height          =   195
            Index           =   66
            Left            =   5190
            TabIndex        =   74
            Top             =   3000
            Width           =   645
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
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
            Index           =   65
            Left            =   360
            TabIndex        =   73
            Top             =   1320
            Width           =   1035
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Intereses Corridos"
            Height          =   195
            Index           =   64
            Left            =   390
            TabIndex        =   72
            Top             =   3660
            Width           =   1260
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisiones"
            Height          =   195
            Index           =   63
            Left            =   390
            TabIndex        =   71
            Top             =   3330
            Width           =   795
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "SubTotal"
            Height          =   195
            Index           =   61
            Left            =   390
            TabIndex        =   70
            Top             =   3000
            Width           =   645
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Precio"
            BeginProperty Font 
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
            Left            =   360
            TabIndex        =   69
            Top             =   930
            Width           =   555
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Precio"
            BeginProperty Font 
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
            Left            =   5280
            TabIndex        =   68
            Top             =   930
            Width           =   555
         End
         Begin VB.Label lblAnalitica 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "???-????????"
            Height          =   285
            Left            =   11400
            TabIndex        =   67
            Tag             =   "0.00"
            Top             =   240
            Width           =   2025
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Analítica"
            BeginProperty Font 
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
            Index           =   84
            Left            =   10200
            TabIndex        =   66
            Top             =   255
            Width           =   765
         End
      End
      Begin VB.Frame fraComisionMontoFL2 
         Caption         =   "Comisiones y Montos - Plazo (FL2)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4545
         Left            =   7980
         TabIndex        =   19
         Top             =   3600
         Width           =   5175
         Begin VB.TextBox txtPrecioUnitario 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   1
            Left            =   2625
            MaxLength       =   45
            TabIndex        =   30
            Top             =   360
            Width           =   1340
         End
         Begin VB.TextBox txtComisionConasev 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   1
            Left            =   4290
            MaxLength       =   45
            TabIndex        =   29
            Top             =   2610
            Width           =   2025
         End
         Begin VB.TextBox txtComisionFondo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   1
            Left            =   4290
            MaxLength       =   45
            TabIndex        =   28
            Top             =   2250
            Width           =   2025
         End
         Begin VB.TextBox txtComisionCavali 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   1
            Left            =   4290
            MaxLength       =   45
            TabIndex        =   27
            Top             =   1890
            Width           =   2025
         End
         Begin VB.TextBox txtComisionBolsa 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   1
            Left            =   4290
            MaxLength       =   45
            TabIndex        =   26
            Top             =   1530
            Width           =   2025
         End
         Begin VB.TextBox txtComisionAgente 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   1
            Left            =   4290
            MaxLength       =   45
            TabIndex        =   25
            Top             =   1185
            Width           =   2025
         End
         Begin VB.TextBox txtInteresCorrido 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   1
            Left            =   4290
            MaxLength       =   45
            TabIndex        =   24
            Top             =   3465
            Width           =   2025
         End
         Begin VB.CommandButton Command1 
            Caption         =   "#"
            Height          =   375
            Left            =   480
            TabIndex        =   23
            ToolTipText     =   "Calcular Valor al Vencimiento y TIRs de la orden"
            Top             =   4845
            Width           =   375
         End
         Begin VB.CheckBox chkAplicar 
            Caption         =   "Aplicar"
            BeginProperty Font 
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
            Index           =   1
            Left            =   360
            TabIndex        =   22
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox txtIntAdicional 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   1
            Left            =   360
            MaxLength       =   45
            TabIndex        =   21
            Top             =   3840
            Width           =   2025
         End
         Begin VB.CheckBox chkInteresCorrido 
            Caption         =   "Interés Corrido"
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
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   20
            Top             =   2520
            Width           =   1935
         End
         Begin TAMControls.TAMTextBox txtPorcenIgv 
            Height          =   315
            Index           =   1
            Left            =   2640
            TabIndex        =   31
            Top             =   3000
            Width           =   1335
            _ExtentX        =   2355
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
            Container       =   "frmOrdenDescuentoContratos.frx":0214
            Text            =   "0.000000"
            Decimales       =   6
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            EnterTab        =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   6
         End
         Begin TAMControls.TAMTextBox txtPorcenIgvInt 
            Height          =   315
            Index           =   1
            Left            =   2640
            TabIndex        =   32
            Top             =   2790
            Width           =   1335
            _ExtentX        =   2355
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
            Container       =   "frmOrdenDescuentoContratos.frx":0230
            Text            =   "0.000000"
            Decimales       =   6
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            EnterTab        =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   6
         End
         Begin TAMControls.TAMTextBox txtPorcenAgente 
            Height          =   315
            Index           =   1
            Left            =   2640
            TabIndex        =   33
            Top             =   1200
            Width           =   1335
            _ExtentX        =   2355
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
            Container       =   "frmOrdenDescuentoContratos.frx":024C
            Text            =   "0.000000"
            Decimales       =   6
            Estilo          =   4
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            EnterTab        =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   6
         End
         Begin TAMControls.TAMTextBox txtImptoInteresCorrido 
            Height          =   315
            Index           =   1
            Left            =   1440
            TabIndex        =   34
            Top             =   3000
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
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
            Container       =   "frmOrdenDescuentoContratos.frx":0268
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
         Begin TAMControls.TAMTextBox txtImptoInteres 
            Height          =   315
            Index           =   1
            Left            =   1320
            TabIndex        =   35
            Top             =   360
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
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
            Container       =   "frmOrdenDescuentoContratos.frx":0284
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
         Begin TAMControls.TAMTextBox txtImptoInteresAdic 
            Height          =   315
            Index           =   1
            Left            =   1320
            TabIndex        =   36
            Top             =   720
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
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
            Container       =   "frmOrdenDescuentoContratos.frx":02A0
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
            Caption         =   "Precio (%)"
            BeginProperty Font 
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
            Left            =   390
            TabIndex        =   64
            Top             =   375
            Width           =   870
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisión SAB"
            BeginProperty Font 
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
            Left            =   390
            TabIndex        =   63
            Top             =   1200
            Width           =   1170
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisión BVL"
            BeginProperty Font 
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
            Left            =   390
            TabIndex        =   62
            Top             =   1560
            Width           =   1155
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisión Cavali"
            BeginProperty Font 
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
            Left            =   390
            TabIndex        =   61
            Top             =   1935
            Width           =   1335
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisión Fondo Garantía"
            BeginProperty Font 
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
            Index           =   59
            Left            =   390
            TabIndex        =   60
            Top             =   2295
            Width           =   2145
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisión Conasev"
            BeginProperty Font 
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
            Index           =   60
            Left            =   390
            TabIndex        =   59
            Top             =   2670
            Width           =   1545
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "IGV"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   62
            Left            =   390
            TabIndex        =   58
            Top             =   3030
            Width           =   270
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "SubTotal"
            BeginProperty Font 
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
            Index           =   72
            Left            =   2640
            TabIndex        =   57
            Top             =   735
            Width           =   780
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Interés Corrido"
            BeginProperty Font 
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
            Index           =   73
            Left            =   2640
            TabIndex        =   56
            Top             =   3480
            Width           =   1245
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
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
            Index           =   76
            Left            =   2640
            TabIndex        =   55
            Top             =   3840
            Width           =   1035
         End
         Begin VB.Label lblDescripMoneda 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
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
            Index           =   1
            Left            =   4455
            TabIndex        =   54
            Top             =   240
            Width           =   1665
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4290
            TabIndex        =   53
            Tag             =   "0.00"
            Top             =   4950
            Width           =   2025
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1080
            TabIndex        =   52
            Tag             =   "0.00"
            Top             =   4950
            Width           =   1335
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2625
            TabIndex        =   51
            Tag             =   "0.00"
            Top             =   4950
            Width           =   1335
         End
         Begin VB.Label lblSubTotal 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   1
            Left            =   4290
            TabIndex        =   50
            Tag             =   "0.00"
            ToolTipText     =   "SubTotal de la Orden"
            Top             =   720
            Width           =   2025
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   3
            X1              =   2580
            X2              =   6300
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Label lblComisionIgv 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   1
            Left            =   4290
            TabIndex        =   49
            Tag             =   "0.00"
            ToolTipText     =   "Monto de Comisión IGV"
            Top             =   2970
            Width           =   2025
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   4
            X1              =   360
            X2              =   6300
            Y1              =   3345
            Y2              =   3345
         End
         Begin VB.Label lblMontoTotal 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   1
            Left            =   4290
            TabIndex        =   48
            Tag             =   "0.00"
            ToolTipText     =   "Monto Total de la Orden"
            Top             =   3825
            Width           =   2025
         End
         Begin VB.Label lblPorcenBolsa 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   1
            Left            =   2625
            TabIndex        =   47
            Tag             =   "0"
            ToolTipText     =   "Porcentaje de Comisión BVL"
            Top             =   1530
            Width           =   1335
         End
         Begin VB.Label lblPorcenCavali 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   1
            Left            =   2625
            TabIndex        =   46
            Tag             =   "0"
            ToolTipText     =   "Porcentaje de Comisión Cavali"
            Top             =   1890
            Width           =   1335
         End
         Begin VB.Label lblPorcenFondo 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   1
            Left            =   2625
            TabIndex        =   45
            Tag             =   "0"
            ToolTipText     =   "Porcentaje de Comisión Fondo Liquidación"
            Top             =   2250
            Width           =   1335
         End
         Begin VB.Label lblPorcenConasev 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   1
            Left            =   2625
            TabIndex        =   44
            Tag             =   "0"
            ToolTipText     =   "Porcentaje de Comisión Conasev"
            Top             =   2610
            Width           =   1335
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Valor al Vencimiento"
            BeginProperty Font 
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
            Index           =   79
            Left            =   4440
            TabIndex        =   43
            Top             =   4320
            Width           =   1755
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "TIR Bruta"
            BeginProperty Font 
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
            Index           =   80
            Left            =   1320
            TabIndex        =   42
            Top             =   4320
            Width           =   840
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "TIR Neta"
            BeginProperty Font 
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
            Index           =   81
            Left            =   2880
            TabIndex        =   41
            Top             =   4320
            Width           =   795
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   5
            X1              =   360
            X2              =   6300
            Y1              =   4200
            Y2              =   4200
         End
         Begin VB.Label lblIntAdelantado 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   1
            Left            =   360
            TabIndex        =   40
            Tag             =   "0.00"
            ToolTipText     =   "SubTotal de la Orden"
            Top             =   3600
            Width           =   2025
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "IGV Intereses"
            BeginProperty Font 
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
            Index           =   112
            Left            =   360
            TabIndex        =   39
            Top             =   2850
            Width           =   1170
         End
         Begin VB.Label lblComisionIgvInt 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   1
            Left            =   4170
            TabIndex        =   38
            Tag             =   "0.00"
            ToolTipText     =   "Monto de Comisión IGV"
            Top             =   2760
            Width           =   2025
         End
         Begin VB.Label lblDiasAdic 
            AutoSize        =   -1  'True
            Caption         =   "Interés Adicional"
            BeginProperty Font 
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
            TabIndex        =   37
            Top             =   3360
            Width           =   1425
         End
      End
      Begin VB.Frame fraDatosAnexo 
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
         Height          =   1095
         Left            =   -74730
         TabIndex        =   1
         Top             =   3060
         Width           =   13935
         Begin TAMControls.TAMTextBox txtMontoConsumido 
            Height          =   315
            Left            =   5550
            TabIndex        =   2
            Top             =   390
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
            Locked          =   -1  'True
            Container       =   "frmOrdenDescuentoContratos.frx":02BC
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
         Begin TAMControls.TAMTextBox txtMontoSolicitud 
            Height          =   315
            Left            =   2010
            TabIndex        =   3
            Top             =   390
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
            Locked          =   -1  'True
            Container       =   "frmOrdenDescuentoContratos.frx":02D8
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
         Begin TAMControls.TAMTextBox txtMontoDesembolso 
            Height          =   315
            Left            =   11970
            TabIndex        =   295
            Top             =   390
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
            Locked          =   -1  'True
            Container       =   "frmOrdenDescuentoContratos.frx":02F4
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
         Begin TAMControls.TAMTextBox txtNumeroDesembolso 
            Height          =   315
            Left            =   9240
            TabIndex        =   298
            Top             =   390
            Width           =   495
            _ExtentX        =   873
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
            Container       =   "frmOrdenDescuentoContratos.frx":0310
            Decimales       =   2
            Estilo          =   3
            CambiarConFoco  =   -1  'True
            ColorEnfoque    =   8454143
            EnterTab        =   -1  'True
            Apariencia      =   1
            Borde           =   1
            DecimalesValue  =   2
         End
         Begin VB.Label lblMontoDesembolso 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Desembolso Nº"
            BeginProperty Font 
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
            Left            =   7170
            TabIndex        =   297
            Top             =   450
            Width           =   1290
         End
         Begin VB.Label lblMontoDesembolso 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Monto Desembolso"
            BeginProperty Font 
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
            Left            =   10110
            TabIndex        =   296
            Top             =   450
            Width           =   1695
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto Solicitud"
            BeginProperty Font 
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
            Index           =   163
            Left            =   330
            TabIndex        =   5
            Top             =   450
            Width           =   1335
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto Consumido"
            BeginProperty Font 
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
            Index           =   176
            Left            =   3810
            TabIndex        =   4
            Top             =   450
            Width           =   1515
         End
      End
      Begin TAMControls2.ucBotonEdicion2 cmdOpcion1 
         Height          =   735
         Left            =   -74100
         TabIndex        =   292
         Top             =   8520
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
         Height          =   1935
         Left            =   -74730
         TabIndex        =   213
         Top             =   450
         Width           =   13935
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
            Left            =   12300
            Picture         =   "frmOrdenDescuentoContratos.frx":032C
            Style           =   1  'Graphical
            TabIndex        =   290
            ToolTipText     =   "Enviar a BackOffice"
            Top             =   1140
            Width           =   1200
         End
         Begin VB.ComboBox cboLineaClienteLista 
            Height          =   315
            Left            =   8880
            Style           =   2  'Dropdown List
            TabIndex        =   288
            Top             =   1200
            Visible         =   0   'False
            Width           =   3105
         End
         Begin VB.ComboBox cboFondo 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   216
            Top             =   360
            Width           =   4785
         End
         Begin VB.ComboBox cboTipoInstrumento 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   215
            Top             =   780
            Width           =   4785
         End
         Begin VB.ComboBox cboEstado 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   214
            Top             =   1200
            Width           =   4785
         End
         Begin MSComCtl2.DTPicker dtpFechaOrdenDesde 
            Height          =   285
            Left            =   9600
            TabIndex        =   217
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
            Format          =   48234497
            CurrentDate     =   38785
         End
         Begin MSComCtl2.DTPicker dtpFechaOrdenHasta 
            Height          =   285
            Left            =   11955
            TabIndex        =   218
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
            Format          =   48234497
            CurrentDate     =   38785
         End
         Begin MSComCtl2.DTPicker dtpFechaLiquidacionDesde 
            Height          =   285
            Left            =   9600
            TabIndex        =   219
            Top             =   780
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
            Format          =   48234497
            CurrentDate     =   38785
         End
         Begin MSComCtl2.DTPicker dtpFechaLiquidacionHasta 
            Height          =   285
            Left            =   11955
            TabIndex        =   220
            Top             =   780
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
            Format          =   48234497
            CurrentDate     =   38785
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
            Index           =   116
            Left            =   7200
            TabIndex        =   289
            Top             =   1245
            Visible         =   0   'False
            Width           =   1500
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
            Left            =   11280
            TabIndex        =   229
            Top             =   795
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
            Index           =   45
            Left            =   8880
            TabIndex        =   228
            Top             =   795
            Width           =   555
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
            Left            =   7200
            TabIndex        =   227
            Top             =   795
            Width           =   1560
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
            Left            =   7200
            TabIndex        =   226
            Top             =   375
            Width           =   1110
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
            Left            =   360
            TabIndex        =   225
            Top             =   375
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
            Left            =   8880
            TabIndex        =   224
            Top             =   375
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
            Left            =   11280
            TabIndex        =   223
            Top             =   375
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
            Left            =   360
            TabIndex        =   222
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
            Left            =   360
            TabIndex        =   221
            Top             =   1245
            Width           =   600
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmOrdenDescuentoContratos.frx":0887
         Height          =   5175
         Left            =   -74760
         OleObjectBlob   =   "frmOrdenDescuentoContratos.frx":08A1
         TabIndex        =   185
         Top             =   2610
         Width           =   13905
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   35
         Left            =   7080
         TabIndex        =   285
         Top             =   5100
         Visible         =   0   'False
         Width           =   405
      End
   End
   Begin MSAdodcLib.Adodc adoConsulta 
      Height          =   330
      Left            =   10140
      Top             =   8730
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
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   12720
      TabIndex        =   291
      Top             =   8520
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
      Caption0        =   "&Salir"
      Tag0            =   "9"
      Visible0        =   0   'False
      ToolTipText0    =   "Salir"
      UserControlWidth=   1200
   End
   Begin TAMControls2.ucBotonEdicion2 cmdAccion 
      Height          =   735
      Left            =   5640
      TabIndex        =   294
      Top             =   8520
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
End
Attribute VB_Name = "frmOrdenDescuentoContratos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrFondo()                As String, arrFondoOrden()              As String
Dim arrTipoInstrumento()      As String, arrTipoInstrumentoOrden()    As String
Dim arrEstado()               As String, arrTipoOrden()               As String
Dim arrOperacion()            As String, arrNegociacion()             As String
Dim arrEmisor()               As String, arrMoneda()                  As String
Dim arrObligado()             As String, arrGestor()                  As String
Dim arrBaseAnual()            As String, arrTipoTasa()                As String
Dim arrOrigen()               As String, arrClaseInstrumento()        As String
Dim arrLineaClienteLista()    As String, arrComisionista()            As String
Dim arrTitulo()               As String, arrSubClaseInstrumento()     As String

Dim arrConceptoCosto()        As String
Dim arrLineaCliente()         As String

Dim arrResponsablePago()      As String, arrResponsablePagoCancel()   As String
Dim arrViaCobranza()          As String
Dim strCodFondo               As String, strCodFondoOrden             As String
Dim strCodTipoInstrumento     As String, strCodTipoInstrumentoOrden   As String
Dim strCodEstado              As String, strCodTipoOrden              As String
Dim strCodOperacion           As String, strCodNegociacion            As String
Dim strCodEmisor              As String, strCodMoneda                 As String
Dim strCodObligado            As String, strCodGestor                 As String
Dim strCodBaseAnual           As String, strCodTipoTasa               As String
Dim strCodOrigen              As String, strCodClaseInstrumento       As String

Dim strCodTitulo              As String, strCodSubClaseInstrumento    As String
Dim strCodConcepto            As String, strCodReportado              As String
Dim strCodGarantia            As String, strCodAgente                 As String
Dim strEstado                 As String, strSQL                       As String
Dim strCodFiador              As String
Dim strLineaCliente           As String
Dim strLineaClienteLista      As String
Dim strResponsablePago        As String, strResponsablePagoCancel     As String
Dim arrPagoInteres()          As String

Dim strCodFile                As String, strCodAnalitica              As String
Dim strCodAnaliticaOrig       As String
Dim strCodGrupo               As String, strCodCiiu                   As String
Dim strEstadoOrden            As String, strCodCategoria              As String
Dim strCodRiesgo              As String, strCodSubRiesgo              As String
Dim strCalcVcto               As String
Dim strCodTipoCostoBolsa      As String, strCodTipoCostoConasev       As String
Dim strCodTipoCostoFondo      As String, strCodTipoCavali             As String
Dim strIndPacto               As String
Dim strIndNegociable          As String, strCodigosFile               As String

Dim strCodCobroInteres        As String, strViaCobranza               As String
Dim dblTipoCambio             As Double
Dim dblComisionBolsa          As Double, dblComisionConasev           As Double
Dim dblComisionFondo          As Double, dblComisionCavali            As Double
Dim intBaseCalculo            As Integer

Dim indCargaPantalla          As Boolean

Dim strCodComisionista        As String
Dim numSecCondicion           As Integer

Dim rsg                       As New ADODB.Recordset
Dim rsgVcto                   As New ADODB.Recordset

Dim strCodMonedaParEvaluacion As String
Dim strCodMonedaParPorDefecto As String
Dim strNumAnexo               As String

Dim strCodPersonaLim          As String
Dim strTipoPersonaLim         As String
Dim intDiasAdicionales        As Integer
Dim datFechaVctoAdicional     As Date

Dim blnCargadoDesdeCartera    As Boolean
Dim blnCargarCabeceraAnexo    As Boolean
Dim blnCancelaPrepago         As Boolean

Dim dblComisionOperacion      As Double          'Comisión que le corresponde a cada operación
Dim strCodMonedaComision      As String          'Moneda de expresión de la comisión
Dim strPersonalizaComision    As String
Dim dblPorcDescuento          As Double
Dim blnFlag                   As Boolean
Dim strCodFondoDescuento      As String

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
            .TabEnabled(2) = True
          
            .Tab = 1
        End With
        
    Else
        MsgBox "Acceso a Negociación Denegada", vbCritical, Me.Caption
    End If
    
End Sub

Private Sub AplicarCostos(Index As Integer)
        
    If strCodTipoCostoBolsa = Codigo_Tipo_Costo_Monto Then
        txtComisionBolsa(Index).Text = CStr(dblComisionBolsa)
    Else
        AsignaComision strCodTipoCostoBolsa, dblComisionBolsa, txtComisionBolsa(Index)
    End If
    
    If strCodTipoCostoConasev = Codigo_Tipo_Costo_Monto Then
        txtComisionConasev(Index).Text = CStr(dblComisionConasev)
    Else
        AsignaComision strCodTipoCostoConasev, dblComisionConasev, txtComisionConasev(Index)
    End If
    
    If strCodTipoCostoFondo = Codigo_Tipo_Costo_Monto Then
        txtComisionFondo(Index).Text = CStr(dblComisionFondo)
    Else
        AsignaComision strCodTipoCostoFondo, dblComisionFondo, txtComisionFondo(Index)
    End If
    
    If strCodTipoCavali = Codigo_Tipo_Costo_Monto Then
        txtComisionCavali(Index).Text = CStr(dblComisionCavali)
    Else
        AsignaComision strCodTipoCavali, dblComisionCavali, txtComisionCavali(Index)
    End If
                     
    Call CalculoTotal(Index)
    
End Sub

Private Sub CalcularTirBruta()

    Dim dblTasaCalculada As Double

    If CDbl(txtPrecioUnitario(0).Text) = 0 Then
        MsgBox "Por favor ingrese el Precio.", vbCritical, Me.Caption
        Exit Sub
    End If

    Me.MousePointer = vbHourglass

    If CDbl(txtPrecioUnitario(0).Text) > 0 Then
        ReDim Array_Monto(1): ReDim Array_Dias(1)
        Array_Monto(0) = CDec((CCur(lblSubTotal(0).Caption) + txtInteresCorrido(0).Text) * -1)
        Array_Dias(0) = dtpFechaLiquidacion.Value
            
        If strCodBaseAnual = Codigo_Base_Actual_Actual Or strCodBaseAnual = Codigo_Base_Actual_365 Or strCodBaseAnual = Codigo_Base_30_365 Then
            If strCodTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
                dblTasaCalculada = ((1 + (CDbl(txtTasa.Text) / 100)) ^ (CDbl(txtDiasPlazo) / 365)) - 1
            Else
                dblTasaCalculada = 1 + ((CDbl(txtTasa.Text) / 100 / 365) * CDbl(txtDiasPlazo))
            End If

        Else

            If strCodTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
                dblTasaCalculada = ((1 + (CDbl(txtTasa.Text) / 100)) ^ (CDbl(txtDiasPlazo) / 360)) - 1
            Else
                dblTasaCalculada = 1 + ((CDbl(txtTasa.Text) / 100 / 360) * CDbl(txtDiasPlazo))
            End If
        End If

        If strCalcVcto = "D" Then
            Array_Monto(1) = CDec(txtCantidad.Text)
        Else
            Array_Monto(1) = CDbl(txtCantidad.Text) * (1 + dblTasaCalculada)
        End If
            
        Array_Dias(1) = dtpFechaVencimiento.Value
        lblTirBruta.Caption = CStr(TIR(Array_Monto(), Array_Dias(), (10 / 100)) * 100)
        lblTirBrutaResumen.Caption = lblTirBruta.Caption

        If strCodTipoOrden = Codigo_Orden_Pacto Then lblTirBrutaResumen.Caption = "0"
    End If

    Me.MousePointer = vbDefault

End Sub

Private Sub CalcularTirNeta()

    Dim dblTir           As Double
    Dim dblTasaCalculada As Double

    If CDbl(lblSubTotal(0).Caption) <= 0 Then
        MsgBox "Por favor ingrese los datos necesarios para hallar la TIR Neta", vbCritical, Me.Caption
        Exit Sub
    End If

    Me.MousePointer = vbHourglass
    
    ReDim Array_Monto(1): ReDim Array_Dias(1)

    Array_Monto(0) = CDec((CCur(lblSubTotal(0).Caption) + CCur(txtInteresCorrido(0).Text) + CCur(txtComisionAgente(0).Text) + CCur(txtComisionBolsa(0).Text) + CCur(txtComisionConasev(0).Text) + CCur(lblComisionIgv(0).Caption)) * -1)
    Array_Dias(0) = dtpFechaLiquidacion.Value
    
    If strCodBaseAnual = Codigo_Base_Actual_Actual Or strCodBaseAnual = Codigo_Base_Actual_365 Or strCodBaseAnual = Codigo_Base_30_365 Then
        If strCodTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
            dblTasaCalculada = ((1 + (CDbl(txtTasa.Text) / 100)) ^ (CDbl(txtDiasPlazo) / 365)) - 1
        Else
            dblTasaCalculada = 1 + ((CDbl(txtTasa.Text) / 100 / 365) * CDbl(txtDiasPlazo))
        End If

    Else

        If strCodTipoTasa = Codigo_Tipo_Tasa_Efectiva Then
            dblTasaCalculada = ((1 + (CDbl(txtTasa.Text) / 100)) ^ (CDbl(txtDiasPlazo) / 360)) - 1
        Else
            dblTasaCalculada = 1 + ((CDbl(txtTasa.Text) / 100 / 360) * CDbl(txtDiasPlazo))
        End If
    End If

    If strCalcVcto = "D" Then
        Array_Monto(1) = CDec(txtCantidad.Text)
    Else
        Array_Monto(1) = CDbl(txtCantidad.Text) * (1 + dblTasaCalculada)
    End If

    Array_Dias(1) = dtpFechaVencimiento.Value

    dblTir = TIR(Array_Monto(), Array_Dias(), (10 / 100)) * 100

    lblTirNeta.Caption = CStr(dblTir)
    lblTirNetaResumen.Caption = CStr(dblTir)

    If strCodTipoOrden = Codigo_Orden_Pacto Then lblTirNetaResumen.Caption = "0"
    Me.MousePointer = vbDefault

End Sub

Private Sub CalcularValorVencimiento()

    If DateDiff("d", dtpFechaLiquidacion.Value, dtpFechaVencimiento.Value) < 0 Then
        MsgBox "La Fecha de vencimiento debe ser posterior a la Fecha de Liquidación.", vbCritical, Me.Caption
        lblMontoVencimiento.Caption = "0"
    Else
    
        Dim intNumDias30 As Integer
        
        '*** Hallar los días 30/360,30/365 ***
        intNumDias30 = Dias360(dtpFechaLiquidacion.Value, dtpFechaVencimiento.Value, True)
        
        If Not IsNumeric(txtDiasPlazo.Text) Then txtDiasPlazo.Text = "0"
        lblMontoVencimiento.Caption = CStr(ValorVencimiento(CCur(txtCantidad.Text), CDbl(txtTasa.Text), intBaseCalculo, CInt(txtDiasPlazo.Text), intNumDias30, strCodTipoTasa, strCodBaseAnual))

    End If

    lblVencimientoResumen.Caption = lblMontoVencimiento.Caption
    lblDiasPlazo.Caption = txtDiasPlazo.Text

End Sub

Private Sub CalcularPrecio()

    Dim intBaseCalculo As Integer

    If CInt(txtDiasPlazo.Text) <= 0 Then
        MsgBox "Por favor ingrese los datos necesarios para hallar el Precio", vbCritical, Me.Caption
        txtDiasPlazo.SetFocus
        Exit Sub
    End If
    
    If strCalcVcto <> "D" Then
        If CInt(txtTasa.Text) <= 0 Then
            MsgBox "Por favor ingrese los datos necesarios para hallar el Precio", vbCritical, Me.Caption
            txtTasa.SetFocus
            Exit Sub
        End If
    End If
    
    If CDbl(lblSubTotal(0).Caption) <= 0 Then
        MsgBox "Por favor ingrese los datos necesarios para hallar el Precio", vbCritical, Me.Caption
        txtValorNominal.SetFocus
        Exit Sub
    End If
    
    If DateDiff("d", dtpFechaLiquidacion.Value, dtpFechaVencimiento) < 0 Then
        MsgBox "La Fecha de vencimiento debe ser posterior a la Fecha de Emisión.", vbCritical, Me.Caption
        txtPrecioUnitario(0).Text = "0"
    Else
        intBaseCalculo = 360

        If strCodBaseAnual = Codigo_Base_Actual_Actual Then intBaseCalculo = 365
        If strCodBaseAnual = Codigo_Base_Actual_365 Then intBaseCalculo = 365
        If strCodBaseAnual = Codigo_Base_30_365 Then intBaseCalculo = 365
        
        If strCalcVcto = "D" Then
            If Not IsNumeric(txtDiasPlazo.Text) Then txtDiasPlazo.Text = "0"
            txtTasa.Text = CStr((ValorTasa(CCur(lblMontoVencimiento.Caption), CCur(lblSubTotal(0).Caption), intBaseCalculo, CInt(txtDiasPlazo.Text))) * 100)
        Else

            If Not IsNumeric(txtDiasPlazo.Text) Then txtDiasPlazo.Text = "0"
            txtPrecioUnitario(0).Text = "100"
        End If
    End If

    lblDiasPlazo.Caption = txtDiasPlazo.Text

End Sub

Private Sub IniciarComisiones()

    Dim intContador As Integer
    
    For intContador = 0 To 1
        txtComisionAgente(intContador).Text = "0"
        txtComisionBolsa(intContador).Text = "0"
        txtComisionCavali(intContador).Text = "0"
        txtComisionFondo(intContador).Text = "0"
        txtComisionConasev(intContador).Text = "0"
        lblComisionIgv(intContador).Caption = "0"
        
        txtPorcenAgente(intContador).Text = "0"
        lblPorcenBolsa(intContador).Caption = "0"
        lblPorcenCavali(intContador).Caption = "0"
        lblPorcenFondo(intContador).Caption = "0"
        lblPorcenConasev(intContador).Caption = "0"
        
        lblPrecioResumen(intContador).Caption = "0"
        lblSubTotalResumen(intContador).Caption = "0"
        lblComisionesResumen(intContador).Caption = "0"
        lblInteresesResumen(intContador).Caption = "0"
        lblTotalResumen(intContador).Caption = "0"
        
    Next
        
End Sub

Private Sub LlenarFormulario(strModo As String)

    Dim adoRecord       As ADODB.Recordset
    
    Dim intRegistro     As Integer
    Dim strCambiarTCOpe As String
  
    Select Case strModo

        Case Reg_Adicion
        
            If blnCargarCabeceraAnexo = False Then  'si no he precargado datos
            
                chkInteresCorrido(0).Value = vbUnchecked

                lblFechaVencimientoAdic.Visible = False

                chkTitulo.Value = vbUnchecked
                intRegistro = ObtenerItemLista(arrFondoOrden(), strCodFondo)

                If intRegistro >= 0 Then cboFondoOrden.ListIndex = intRegistro
            
                cboTipoInstrumentoOrden.ListIndex = -1

                If cboTipoInstrumentoOrden.ListCount > 0 Then cboTipoInstrumentoOrden.ListIndex = 0
                                        
                cboTipoOrden.ListIndex = -1

                If cboTipoOrden.ListCount > 0 Then cboTipoOrden.ListIndex = 0
        
                cboOperacion.ListIndex = -1

                If cboOperacion.ListCount > 0 Then cboOperacion.ListIndex = 0
            
                cboNegociacion.ListIndex = -1

                If cboNegociacion.ListCount > 0 Then cboNegociacion.ListIndex = 0
                
                cboEmisor.ListIndex = -1

                If cboEmisor.ListCount > 0 Then cboEmisor.ListIndex = 0
                
                cboGestor.ListIndex = -1

                If cboGestor.ListCount > 0 Then cboGestor.ListIndex = 0
            
                txtTasa.Text = "0"
                
                cboBaseAnual.ListIndex = -1

                If cboBaseAnual.ListCount > 0 Then cboBaseAnual.ListIndex = 0
                
                cboTipoTasa.ListIndex = -1

                If cboTipoTasa.ListCount > 0 Then cboTipoTasa.ListIndex = 0
            
                intRegistro = ObtenerItemLista(arrOrigen(), Codigo_Negociacion_Local)

                If intRegistro >= 0 Then cboOrigen.ListIndex = intRegistro
                
                txtPorcenDctoValorNominal.Text = dblPorcDescuento
                txtMontoSolicitud.Text = 0#
                txtMontoConsumido.Text = 0#
                txtPorcenAgente(0).Text = 0#

            End If
            
            If cboResponsablePago.ListCount > 0 Then cboResponsablePago.ListIndex = 1
           
            cboObligado.ListIndex = -1

            If cboObligado.ListCount > 0 Then cboObligado.ListIndex = 0

            intRegistro = ObtenerItemLista(arrMoneda(), strCodMoneda)

            If intRegistro >= 0 Then cboMoneda.ListIndex = intRegistro
            
            txtNumOperacionOrig.Text = ""
            dtpFechaOrden.Value = gdatFechaActual
            dtpFechaOrdenCancel.Value = gdatFechaActual
            dtpFechaLiquidacion.Value = dtpFechaOrden.Value
            dtpFechaLiquidacionCancel.Value = dtpFechaOrden.Value
            lblFechaLiquidacion.Caption = CStr(dtpFechaOrden.Value)
            
            txtDiasPlazo.Text = "0"
            lblDiasPlazo.Caption = "0"
            
            txtInteresCorrido(0).Text = "0"
            txtImptoInteresCorrido(0).Text = "0"
            
            txtDescripOrden.Text = Valor_Caracter
            txtDescripOrdenCancel.Text = Valor_Caracter
            txtNemonico.Text = Valor_Caracter
            txtNemonicoCancel.Text = Valor_Caracter
            txtObservacion.Text = Valor_Caracter
            txtPrecioUnitario(0).Text = "100"
            txtPrecioUnitario(1).Text = "0"
            txtValorNominal.Text = "1"

            If blnCargarCabeceraAnexo = False Then txtCantidad.Text = "0"

            lblAnalitica.Caption = "??? - ????????"
            lblStockNominal.Caption = "0"
            lblClasificacion.Caption = Valor_Caracter

            dtpFechaVencimiento.Value = gdatFechaActual
            dtpFechaPago.Value = dtpFechaVencimiento.Value
            lblFechaEmision.Caption = CStr(dtpFechaVencimiento.Value)
            lblFechaVencimiento.Caption = CStr(dtpFechaVencimiento.Value)
            
            lblIntAdelantado(0).Caption = "0"
            lblIntAdelantado(1).Caption = "0"
            
            txtIntAdicional(0).Text = "0"
            txtIntAdicional(1).Text = "0"
            
            If strPersonalizaComision <> "NO" Then
                chkAplicar(0).Value = vbUnchecked
                chkAplicar(1).Value = vbUnchecked
            End If
            
            'fraComisiones.Enabled = False
           
            lblSubTotal(0).Caption = "0"
            lblSubTotal(1).Caption = "0"
            
            Call IniciarComisiones
            
            txtInteresCorrido(0).Text = "0"
            txtImptoInteresCorrido(0).Text = "0"
            txtInteresCorrido(1).Text = "0"
            txtImptoInteresCorrido(1).Text = "0"
            
            lblMontoTotal(0).Caption = "0"
            lblMontoTotal(1).Caption = "0"
            lblTirBruta.Caption = "0"
            lblTirNeta.Caption = "0"
            lblMontoVencimiento.Caption = "0"
            lblVencimientoResumen.Caption = "0"
                        
            lblFechaCupon.Caption = Valor_Caracter
            lblClasificacion.Caption = Valor_Caracter
            lblBaseTasaCupon.Caption = Valor_Caracter
            lblStockNominal.Caption = "0"
            lblMoneda.Caption = Valor_Caracter
            lblCantidadResumen.Caption = "0"
                                                
            lblTirBrutaResumen.Caption = "0"
            lblTirNetaResumen.Caption = "0"
            
            txtMontoVencimiento1.Text = "0"
            txtTirBruta1.Text = "0"

            txtTirBruta1.Tag = 0         'indica cambio directo en la pantalla
            txtPrecioUnitario(0).Tag = 0
            txtMontoVencimiento1.Tag = 0

    End Select
                  
    Set adoRecord = New ADODB.Recordset
    
    adoComm.CommandText = "SELECT upper(ValorParametro) AS CambiarTCOpe FROM ParametroGeneral WHERE CodParametro = '21'"
    Set adoRecord = adoComm.Execute
 
    If Not (adoRecord.EOF) Then
        strCambiarTCOpe = Trim$(adoRecord("CambiarTCOpe"))
    End If
    
    adoRecord.Close: Set adoRecord = Nothing
    
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
        
        'verificar si la orden no está ya anulada
        
        If strCodEstado <> Estado_Orden_Anulada And strCodEstado <> Estado_Orden_Procesada Then
        
            strMensaje = "Se procederá a eliminar la ORDEN " & tdgConsulta.Columns(0) & " por la " & tdgConsulta.Columns(3) & vbNewLine & vbNewLine & vbNewLine & "¿ Seguro de continuar ?"
            
            If MsgBox(strMensaje, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
        
                adoComm.CommandText = "UPDATE InversionOrden SET EstadoOrden='" & Estado_Orden_Anulada & "' WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & "CodTitulo='" & Trim$(tdgConsulta.Columns(2)) & "' AND NumOrden='" & Trim$(tdgConsulta.Columns(0)) & "'"
                    
                adoConn.Execute adoComm.CommandText
                
                adoComm.CommandText = "UPDATE InstrumentoInversion SET IndVigente='' WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & "CodTitulo='" & Trim$(tdgConsulta.Columns(2)) & "'"
                    
                adoConn.Execute adoComm.CommandText
                
                MsgBox Mensaje_Eliminacion_Exitosa, vbExclamation, Me.Caption
                
                tabRFCortoPlazo.TabEnabled(0) = True
                tabRFCortoPlazo.Tab = 0
                Call Buscar
                
                Exit Sub
            End If
        Else

            If strCodEstado = Estado_Orden_Anulada Then
                MsgBox "La orden " & Trim$(tdgConsulta.Columns(0)) & " ya ha sido anulada.", vbExclamation, "Anular Orden"
            Else
                MsgBox "La orden " & Trim$(tdgConsulta.Columns(0)) & " ya ha sido procesada." & vbNewLine & "No se puede anular.", vbCritical, "Anular Orden"
            End If
        End If
        

    End If
    
End Sub

Public Sub Grabar()

    Call Accion(vSave)

End Sub

Public Sub GrabarNew()

    Dim adoRegistro            As ADODB.Recordset
    Dim strFechaOrden          As String, strFechaLiquidacion      As String
    Dim strFechaEmision        As String, strFechaVencimiento      As String
    Dim strFechaPago           As String
    Dim strFechaVctoDcto       As String
    Dim strFechaInteresAdic    As String
    Dim strMensaje             As String, strIndTitulo             As String
    Dim intAccion              As Integer
    Dim lngNumError            As Long
    Dim dblTasaInteres         As Double

    Dim strMsgError            As String
    Dim intDiasAdicionalesVcto As Integer
    Dim strIndDevolucion       As String
    
    On Error GoTo CtrlError
    
    If strEstado = Reg_Consulta Then Exit Sub
    
    If strEstado = Reg_Adicion Then
        If TodoOK() Then
     
            strEstadoOrden = Estado_Orden_Ingresada

            strMensaje = "_____________________________________________________" & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & _
               Space$(8) & "<<<<<     " & Trim$(UCase$(cboFondoOrden.Text)) & "     >>>>>" & Chr$(vbKeyReturn) & _
               "_____________________________________________________" & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & _
               "Para proceder al Registro de la Orden Confirme lo siquiente : " & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & _
               "Fecha de Vencimiento      " & Space$(3) & ">" & Space$(2) & CStr(dtpFechaVencimiento.Value) & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & _
               "Fecha de Operación        " & Space$(3) & ">" & Space$(2) & CStr(dtpFechaOrden.Value) & Chr$(vbKeyReturn) & _
               "Fecha de Liquidación      " & Space$(3) & ">" & Space$(2) & CStr(dtpFechaLiquidacion.Value) & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & _
               "Fecha de Pago             " & Space$(3) & ">" & Space$(2) & CStr(dtpFechaPago.Value) & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & _
               "Tasa                      " & Space$(3) & ">" & Space$(2) & txtTasa.Text & Chr$(vbKeyReturn) & _
               "Interés TOTAL calculado   " & Space$(3) & ">" & Space$(2) & CStr(CDbl(lblIntAdelantado(0).Caption) + CDbl(txtIntAdicional(0).Text)) & Chr$(vbKeyReturn) & _
               "Días por protesto         " & Space$(3) & ">" & Space$(2) & IIf(chkDiasAdicional.Visible = True And chkDiasAdicional.Value = Checked, DateDiff("d", dtpFechaVencimiento.Value, datFechaVctoAdicional), 0) & Chr$(vbKeyReturn) & _
               "Interés Prov. por Protesto" & Space$(3) & ">" & Space$(2) & txtIntAdicional(0).Text & Chr$(vbKeyReturn) & _
               "Cobro de Intereses        " & Space$(3) & ">" & Space$(2) & cboCobroInteres.Text & Chr$(vbKeyReturn) & _
               "Nominal                   " & Space$(3) & ">" & Space$(2) & txtValorNominal.Text & Chr$(vbKeyReturn) & _
               "Porcentaje de Descuento   " & Space$(3) & ">" & Space$(2) & txtPorcenDctoValorNominal.Text & Chr$(vbKeyReturn) & _
               "Valor Nominal Descontado  " & Space$(3) & ">" & Space$(2) & txtValorNominalDcto.Text & Chr$(vbKeyReturn) & _
               "Cantidad                  " & Space$(3) & ">" & Space$(2) & txtCantidad.Text & Chr$(vbKeyReturn) & _
               "Precio Unitario (%)       " & Space$(3) & ">" & Space$(2) & txtPrecioUnitario(0).Text & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & _
               "Monto Total Desembolsado  " & Space$(3) & ">" & Space$(2) & Trim$(lblDescripMoneda(0).Caption) & Space$(1) & lblMontoTotal(0).Caption & Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & _
               "Tir Neta                  " & Space$(3) & ">" & Space$(2) & lblTirNeta.Caption & Chr$(vbKeyReturn) & _
               Chr$(vbKeyReturn) & Chr$(vbKeyReturn) & _
               "¿ Seguro de continuar ?"

            If MsgBox(strMensaje, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                Me.Refresh: Exit Sub
            End If
        
            Me.MousePointer = vbHourglass
            
            strFechaOrden = Convertyyyymmdd(dtpFechaOrden.Value)
            strFechaEmision = strFechaOrden
            strFechaLiquidacion = Convertyyyymmdd(dtpFechaLiquidacion.Value)
            strFechaVencimiento = Convertyyyymmdd(dtpFechaVencimiento.Value)
            strFechaPago = Convertyyyymmdd(dtpFechaPago.Value)
            strFechaInteresAdic = Convertyyyymmdd(datFechaVctoAdicional)

            If chkDiasAdicional.Visible = True And chkDiasAdicional.Value = Checked Then
                intDiasAdicionalesVcto = DateDiff("d", dtpFechaVencimiento.Value, datFechaVctoAdicional)
            Else
                intDiasAdicionalesVcto = 0
            End If
           
            Set adoRegistro = New ADODB.Recordset

            With adoComm
                strIndTitulo = Valor_Caracter
                                
                If strCodTipoOrden = Codigo_Orden_Pacto Then
                    strIndTitulo = Valor_Caracter
                    strCodTipoTasa = Codigo_Tipo_Tasa_Efectiva
                    strCodBaseAnual = Codigo_Base_Actual_365
                    strCodRiesgo = "00"
                    strCodReportado = Valor_Caracter
                    strCodFile = Left$(Trim$(lblAnalitica.Caption), 3)
                ElseIf (strCodTipoOrden = Codigo_Orden_Compra Or strCodTipoOrden = Codigo_Orden_Renovacion) Then

                    If chkTitulo.Value Then
                        strIndTitulo = "X"
                    Else

                        If strCodFile <> "003" And strCodFile <> "010" And strCodFile <> "014" And strCodFile <> "015" And strCodFile <> "016" And strCodFile <> "021" Then
                            strCodAnalitica = ObtenerNuevaAnalitica(strCodFile)   'DIFERENTE
                        End If   'caso contrario: strCodAnalitica Se obtiene en el sp de Grabación de la orden
                        
                    End If

                Else
                    strIndTitulo = Valor_Indicador
                    strCodGarantia = Valor_Caracter
                    strFechaVencimiento = Convertyyyymmdd(Valor_Fecha)
                    strCodReportado = Valor_Caracter
                End If
                
                If strCalcVcto = "V" Then  'Con tasa de interés
                    dblTasaInteres = CDbl(txtTasa.Text)
                Else 'Con Precio
                    dblTasaInteres = 0#
                End If

                .CommandText = "{ call up_IVAdicInversionOrden('" & strCodFondoOrden & "','" & gstrCodAdministradora & "','','" & strFechaOrden & "','" & _
                   strCodTitulo & "','" & Trim$(txtNemonico.Text) & "','" & gstrPeriodoActual & "','" & gstrMesActual & "','" & _
                   "','" & strEstadoOrden & "','" & strCodAnalitica & "','" & strCodFile & "','" & strCodAnalitica & "','" & strCodClaseInstrumento & "','" & _
                   strCodSubClaseInstrumento & "','" & strCodTipoOrden & "','" & strCodOperacion & "','" & strCodNegociacion & "','" & strCodOrigen & "','" & _
                   Trim$(txtDescripOrden.Text) & "','" & strCodEmisor & "','" & strCodAgente & "','" & strCodGarantia & "','" & _
                   strCodComisionista & "'," & numSecCondicion & ",'" & strFechaPago & "','" & _
                   strFechaVencimiento & "','" & strFechaLiquidacion & "','" & strFechaEmision & "','" & strCodMoneda & "'," & txtValorNominal.Value & ",'','" & _
                   strCodMoneda & "','" & strCodMoneda & "'," & CDec(txtValorNominalDcto.Text) & "," & CDec(txtTipoCambio.Text) & "," & CDec(txtTipoCambio.Text) & _
                   "," & txtValorNominal.Value & "," & txtPorcenDctoValorNominal.Value & "," & CDec(txtValorNominalDcto.Text) & ",1,1," & CDec(lblSubTotal(0).Caption) & _
                   "," & CDec(txtInteresCorrido(0).Text) & "," & CDec(txtComisionAgente(0).Text) & "," & CDec(txtComisionCavali(0).Text) & "," & _
                   CDec(txtComisionConasev(0).Text) & "," & CDec(txtComisionBolsa(0).Text) & "," & CDec(txtComisionFondo(0).Text) & ",0,0,0," & _
                   CDec(lblComisionIgv(0).Caption) & "," & CDec(lblMontoTotal(0).Caption) & ",0,0,0,0,0,0,0,0,0,0,0,0,0,0," & CDec(txtMontoVencimiento1.Value) & _
                   "," & CInt(txtDiasPlazo.Text) & ",'','','','','','" & strCodReportado & "','" & strCodEmisor & "','" & strCodEmisor & "','" & strCodEmisor & _
                   "','" & strCodGestor & "','" & strCodFiador & "',0,'','X','X','" & strCodTipoTasa & "','" & strCodBaseAnual & "'," & _
                   CDec(dblTasaInteres) & ",'05','X','07',''," & CDec(dblTasaInteres) & "," & CDec(dblTasaInteres) & "," & CDec(dblTasaInteres) & ",'" & _
                   strCodRiesgo & "','" & strCodSubRiesgo & "','" & Trim$(txtObservacion.Text) & "','" & gstrLogin & "','" & gstrFechaActual & "','','" & _
                   "19000101','" & strCodTitulo & "','" & strCodCobroInteres & "'," & CDec(lblIntAdelantado(0).Caption) & ",0,0," & CDec(txtIntAdicional(0).Text) & _
                   ",0,0,'01'," & CDec(txtPorcenIgvInt(0).Text) & "," & CDec(txtImptoInteres(0).Text) & "," & CDec(txtImptoInteresAdic(0).Text) & ",0," & _
                   CDec(txtPorcenIgv(0).Text) & "," & CDec(lblComisionIgv(0).Caption) & ",0,0,0,0,0,0,'" & Format$(CStr(txtNumAnexo.Text), "0000000000") & "','','','" & _
                   Format$(CStr(txtNumAnexo.Text), "0000000000") & "','" & strLineaCliente & "','" & Codigo_LimiteRE_Cliente & "','" & strCodEmisor & _
                   "','" & strTipoPersonaLim & "','" & strResponsablePago & "','" & strViaCobranza & "'," & CDec(txtValorNominalDcto.Text) & ",1," & _
                   CDec(txtPorcenAgente(0).Text) & "," & CDec(txtDeudaTotal.Text) & ") }"
                
                adoConn.Execute .CommandText
                
                .CommandText = "UPDATE InstrumentoInversion SET Nemotecnico='" & txtNemonico.Text & "' WHERE CodTitulo = '" & strCodTitulo & "'"
                adoConn.Execute .CommandText
                                                                                                      
            End With
            
            Me.MousePointer = vbDefault
        
            MsgBox Mensaje_Adicion_Exitosa, vbExclamation
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
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
    Me.MousePointer = vbDefault

    If Left$(err.Description, 14) <> "Excede Limites" Then
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

    Else
        strMsgError = strMsgError & err.Description
        MsgBox strMsgError, vbCritical, "Limites"
        
    End If
        
End Sub

Private Function TodoOK() As Boolean
        
    Dim adoRegistro As ADODB.Recordset
    
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
                          
    If chkTitulo.Value Then
        If cboTitulo.ListIndex <= 0 Then
            MsgBox "Debe seleccionar el Título.", vbCritical, Me.Caption

            If cboTitulo.Enabled Then cboTitulo.SetFocus
            Exit Function
        End If

    Else

        If cboEmisor.ListIndex <= 0 Then
            MsgBox "Debe seleccionar el Emisor.", vbCritical, Me.Caption

            If cboEmisor.Enabled Then cboEmisor.SetFocus
            Exit Function
        End If
        
        Set adoRegistro = New ADODB.Recordset
        
    End If
        
    If cboLineaCliente.ListIndex < 0 Then
        MsgBox "Debe seleccionar la Línea a afectar.", vbCritical, Me.Caption

        If cboLineaCliente.Enabled Then cboLineaCliente.SetFocus
        Exit Function
    End If
        
    If Trim$(txtDescripOrden.Text) = Valor_Caracter Then
        MsgBox "Debe indicar la Descripción de la ORDEN.", vbCritical, Me.Caption

        If txtDescripOrden.Enabled Then txtDescripOrden.SetFocus
        Exit Function
    End If
        
    If CVDate(dtpFechaVencimiento.Value) > CVDate(dtpFechaVencimiento.Value) Then
        MsgBox "La Fecha de Vencimiento debe ser mayor a la Fecha de Liquidación.", vbCritical, Me.Caption

        If dtpFechaVencimiento.Enabled Then dtpFechaVencimiento.SetFocus
        Exit Function
    End If
    
    If CInt(txtDiasPlazo.Text) = 0 Then
        MsgBox "Debe indicar el número de días de plazo.", vbCritical, Me.Caption

        If txtDiasPlazo.Enabled Then txtDiasPlazo.SetFocus
        Exit Function
    End If
    
    If CDbl(txtTasa.Text) = 0 And strCalcVcto = "V" Then
        MsgBox "Debe indicar la Tasa Facial.", vbCritical, Me.Caption

        If txtTasa.Enabled Then txtTasa.SetFocus
        Exit Function
    End If
    
    If CCur(txtValorNominal.Text) = 0 Then
        MsgBox "Debe indicar el Valor Nominal.", vbCritical, Me.Caption

        If txtValorNominal.Enabled Then txtValorNominal.SetFocus
        Exit Function
    End If
    
    If CVDate(dtpFechaOrden.Value) > CVDate(dtpFechaLiquidacion.Value) Then
        MsgBox "La Fecha de Liquidación debe ser mayor o igual a la Fecha de la ORDEN.", vbCritical, Me.Caption

        If dtpFechaLiquidacion.Enabled Then dtpFechaLiquidacion.SetFocus
        Exit Function
    End If
            
    If CDbl(txtPrecioUnitario(0).Text) = 0 Then
        MsgBox "Debe indicar el Precio.", vbCritical, Me.Caption

        If txtPrecioUnitario(0).Enabled Then txtPrecioUnitario(0).SetFocus
        Exit Function
    End If
    
    If strCodTipoOrden = Codigo_Orden_Venta Then
        If CCur(txtValorNominal.Text) > CCur(lblStockNominal.Caption) Then
            MsgBox "Stock insuficiente para Registrar la Orden de Venta.", vbCritical, Me.Caption

            If txtValorNominal.Enabled Then txtValorNominal.SetFocus
            Exit Function
        End If
    End If
        
    TodoOK = True
  
End Function

Public Sub SubImprimir(Index As Integer)

    Dim frmReporte As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    
    Dim strSeleccionRegistro As String

    If tabRFCortoPlazo.Tab = 1 Then Exit Sub
    
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
                aReportParamF(1) = Mid$(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10)
                aReportParamF(2) = Mid$(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10)
                aReportParamF(3) = Format$(Time$(), "hh:mm:ss")
                aReportParamF(4) = Trim$(cboFondo.Text)
                aReportParamF(5) = gstrNombreEmpresa & Space$(1)
                            
                aReportParamS(0) = strCodFondo
                aReportParamS(1) = gstrCodAdministradora
                aReportParamS(2) = Convertyyyymmdd(Mid$(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10))
                aReportParamS(3) = Convertyyyymmdd(DateAdd("d", 1, CVDate(Mid$(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10))))
                aReportParamS(4) = strCodMoneda
                aReportParamS(5) = strCodTipoInstrumento
            End If

        Case 2
            gstrNameRepo = "PapeletaInversion"
            
            strSeleccionRegistro = "{InversionOrden.FechaOrden} IN 'Fch1' TO 'Fch2'"
            gstrSelFrml = strSeleccionRegistro
            frmRangoFecha.Show vbModal
                        
            If gstrSelFrml <> "0" Then
                Set frmReporte = New frmVisorReporte
    
                ReDim aReportParamS(5)
                ReDim aReportParamFn(1)
                ReDim aReportParamF(1)
                            
                aReportParamFn(0) = "Fondo"
                aReportParamFn(1) = "NombreEmpresa"
                            
                aReportParamF(0) = Trim$(cboFondo.Text)
                aReportParamF(1) = gstrNombreEmpresa & Space$(1)
                            
                aReportParamS(0) = strCodFondo
                aReportParamS(1) = gstrCodAdministradora
                aReportParamS(2) = Convertyyyymmdd(Mid$(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10))
                aReportParamS(3) = Convertyyyymmdd(DateAdd("d", 1, CVDate(Mid$(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10))))
                aReportParamS(4) = strCodMoneda
                aReportParamS(5) = strCodTipoInstrumento
            End If
            
        Case 3, 4

            If Index = 3 Then
                gstrNameRepo = "Anexo"
            Else
                gstrNameRepo = "AnexoCliente"
            End If
                        
            Set frmReporte = New frmVisorReporte
    
            ReDim aReportParamS(9)
            ReDim aReportParamFn(3)
            ReDim aReportParamF(3)
                            
            aReportParamFn(0) = "Usuario"
            aReportParamFn(1) = "Hora"
            aReportParamFn(2) = "Fondo"
            aReportParamFn(3) = "NombreEmpresa"
                            
            aReportParamF(0) = gstrLogin
            aReportParamF(1) = Format$(Time$(), "hh:mm:ss")
            aReportParamF(2) = Trim$(cboFondo.Text)
            aReportParamF(3) = gstrNombreEmpresa & Space$(1)
            
            aReportParamS(0) = strCodFondo
            aReportParamS(1) = gstrCodAdministradora
            aReportParamS(2) = tdgConsulta.Columns(11)
            aReportParamS(3) = tdgConsulta.Columns(12)
            aReportParamS(4) = tdgConsulta.Columns(13)
            aReportParamS(5) = tdgConsulta.Columns(21)
            aReportParamS(6) = tdgConsulta.Columns(23)
            aReportParamS(7) = tdgConsulta.Columns(24)
            aReportParamS(8) = tdgConsulta.Columns(25)
            aReportParamS(9) = tdgConsulta.Columns(9)   'Nro de Anexo
            
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

Public Sub Salir()

    Unload Me
    
End Sub

Private Sub cboBaseAnual_Click()

    strCodBaseAnual = Valor_Caracter

    If cboBaseAnual.ListIndex < 0 Then Exit Sub
    
    strCodBaseAnual = Trim$(arrBaseAnual(cboBaseAnual.ListIndex))
    
    intBaseCalculo = 365

    Select Case strCodBaseAnual

        Case Codigo_Base_30_360: intBaseCalculo = 360

        Case Codigo_Base_Actual_365: intBaseCalculo = 365

        Case Codigo_Base_Actual_360: intBaseCalculo = 360

        Case Codigo_Base_30_365: intBaseCalculo = 365
    End Select
    
    txtValorNominal_Change
    
End Sub

Private Sub cboClaseInstrumento_Click()

    strCodClaseInstrumento = Valor_Caracter

    If cboClaseInstrumento.ListIndex < 0 Then Exit Sub
    
    strCodClaseInstrumento = Trim$(arrClaseInstrumento(cboClaseInstrumento.ListIndex))
    
    strSQL = "SELECT CodLimite CODIGO, DescripLimite DESCRIP FROM LimiteReglamentoEstructuraDetalle where CodEstructura = '" & Codigo_LimiteRE_Cliente & "'" & _
             " and CodLimite = '" & Linea_Contrato_Flujo_Dinerario & "' and Estado  = '01' "
    CargarControlLista strSQL, cboLineaCliente, arrLineaCliente(), ""
    Call cboLineaCliente_Click   'Para obligar a que se seleccione el único elemento de la lista

    If cboLineaCliente.ListCount > 0 Then cboLineaCliente.ListIndex = 0
    
    strSQL = "SELECT CodSubDetalleFile CODIGO,DescripSubDetalleFile DESCRIP FROM InversionSubDetalleFile WHERE " & "CodDetalleFile='" & strCodClaseInstrumento & "' AND CodFile='" & strCodTipoInstrumentoOrden & "' AND IndVigente='X' ORDER BY DescripSubDetalleFile"
        
    CargarControlLista strSQL, cboSubClaseInstrumento, arrSubClaseInstrumento(), Sel_Defecto
    
    If cboSubClaseInstrumento.ListCount > 1 Then
        cboSubClaseInstrumento.ListIndex = ObtenerItemLista(arrSubClaseInstrumento(), strCodClaseInstrumento)
    Else

        If cboSubClaseInstrumento.ListCount > 0 Then cboSubClaseInstrumento.ListIndex = 0
    End If
    
    cboSubClaseInstrumento.Enabled = True

    If strCodClaseInstrumento = "001" Then strCalcVcto = "V"   'tasa de interés
    If strCodClaseInstrumento = "002" Then strCalcVcto = "D"    'Al descuento

    txtNemonico.Text = GenerarNemonico(strCodTipoInstrumentoOrden, strCodClaseInstrumento, strCodEmisor, "")
    txtDescripOrdenCancel.Text = Trim$(cboTipoInstrumentoOrden.Text) & " - " & txtNemonicoCancel.Text

End Sub

Private Function GenerarNemonico(strTipoInstrumento As String, _
                                 strClaseOperacion As String, _
                                 strCodEmisor As String, _
                                 strNumDocumento) As String

    Dim adoTemporal       As ADODB.Recordset
    Dim strNemotecnico    As String
    Dim strValorParametro As String
    Dim strCodParametro   As String

    GenerarNemonico = Valor_Caracter

    If Trim$(strTipoInstrumento) = "" Or Trim$(strClaseOperacion) = "" Or Trim$(strCodEmisor) = "" Then
        Exit Function
    End If
    
    Set adoTemporal = New ADODB.Recordset

    With adoComm
        .CommandText = "SELECT DescripNemonico FROM InstitucionPersona WHERE CodPersona='" & strCodEmisor & "' AND TipoPersona = '" & Codigo_Tipo_Persona_Emisor & "'"
        
        Set adoTemporal = .Execute

        If Not adoTemporal.EOF Then
            strNemotecnico = Trim$(adoTemporal("DescripNemonico"))

            If strNemotecnico = "" Then
                adoTemporal.Close: Set adoTemporal = Nothing
                Exit Function
            End If

        Else
            adoTemporal.Close: Set adoTemporal = Nothing
            Exit Function
        End If
      
        Select Case strTipoInstrumento
        
            Case CodFile_Descuento_Comprobantes_Pago   '"014"

                Select Case strClaseOperacion

                    Case "001"   'Factura
                        strCodParametro = "02"

                    Case "002"   'Recibos por honorarios
                        strCodParametro = "05"

                    Case "003"   'Coleta de venta
                        strCodParametro = "06"
                End Select
                
            Case CodFile_Descuento_Documentos_Cambiario  '"015"

                Select Case strClaseOperacion

                    Case "001"  'Letra
                        strCodParametro = "01"

                    Case "002"  'Pagaré
                        strCodParametro = "08"

                    Case "003"  'Cheque
                        strCodParametro = "07"
                End Select
            
            Case CodFile_Descuento_Flujos_Dinerarios   '"016"

                Select Case strClaseOperacion

                    Case "001"  'Contratos
                        strCodParametro = "09"
                End Select
            
            Case "010" 'Letras
                strCodParametro = "04"
            
        End Select
        
        .CommandText = "SELECT ValorParametro FROM AuxiliarParametro WHERE CodParametro='" & strCodParametro & "' AND CodTipoParametro = 'NEMODF' "
        Set adoTemporal = .Execute

        If Not adoTemporal.EOF Then
            strValorParametro = Trim$(adoTemporal("ValorParametro"))
        End If
        
        adoTemporal.Close: Set adoTemporal = Nothing
        
    End With

    GenerarNemonico = Trim$(strValorParametro) & Trim$(strNemotecnico) & Trim$(strNumDocumento)

End Function

Private Sub cboConceptoCosto_Click()

    Dim adoRegistro As ADODB.Recordset
    
    strCodConcepto = Valor_Caracter

    If cboConceptoCosto.ListIndex < 0 Then Exit Sub
    
    strCodConcepto = Trim$(arrConceptoCosto(cboConceptoCosto.ListIndex))
    
    strCodTipoCostoBolsa = Valor_Caracter: strCodTipoCostoConasev = Valor_Caracter
    strCodTipoCavali = Valor_Caracter: strCodTipoCostoFondo = Valor_Caracter
    dblComisionBolsa = 0: dblComisionConasev = 0
    dblComisionCavali = 0: dblComisionFondo = 0
        
    With adoComm
        Set adoRegistro = New ADODB.Recordset
                
        .CommandText = "SELECT CodCosto,TipoCosto,ValorCosto FROM CostoNegociacion WHERE TipoOperacion='" & strCodConcepto & "' AND TipoValor='" & Codigo_Valor_RentaFija & "' ORDER BY CodCosto"
        Set adoRegistro = .Execute

        Do Until adoRegistro.EOF

            Select Case Trim$(adoRegistro("CodCosto"))

                Case Codigo_Costo_Bolsa
                    strCodTipoCostoBolsa = Trim$(adoRegistro("TipoCosto"))
                    dblComisionBolsa = CDbl(adoRegistro("ValorCosto"))

                Case Codigo_Costo_Conasev
                    strCodTipoCostoConasev = Trim$(adoRegistro("TipoCosto"))
                    dblComisionConasev = CDbl(adoRegistro("ValorCosto"))

                Case Codigo_Costo_Cavali
                    strCodTipoCavali = Trim$(adoRegistro("TipoCosto"))
                    dblComisionCavali = CDbl(adoRegistro("ValorCosto"))

                Case Codigo_Costo_FLiquidacion
                    strCodTipoCostoFondo = Trim$(adoRegistro("TipoCosto"))
                    dblComisionFondo = CDbl(adoRegistro("ValorCosto"))
            End Select

            adoRegistro.MoveNext
        Loop

        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
End Sub

Private Sub cboEmisor_Click()

    Dim adoRegistro As ADODB.Recordset
    
    strCodTitulo = Valor_Caracter: strCodGrupo = Valor_Caracter: strCodCiiu = Valor_Caracter
    strCodEmisor = Valor_Caracter: strCodAnalitica = Valor_Caracter
    lblAnalitica = strCodTipoInstrumentoOrden & "-??????": 'txtValorNominal.Text = "1"
    lblStockNominal = "0": strCodGrupo = Valor_Caracter
    
    If cboEmisor.ListIndex < 0 Then Exit Sub

    strCodEmisor = Left$(Trim$(arrEmisor(cboEmisor.ListIndex)), 8)
    strCodGrupo = Mid$(Trim$(arrEmisor(cboEmisor.ListIndex)), 9, 3)
    strCodCiiu = Right$(Trim$(arrEmisor(cboEmisor.ListIndex)), 4)

    txtNemonico.Text = GenerarNemonico(strCodTipoInstrumentoOrden, strCodClaseInstrumento, strCodEmisor, "")
    txtNemonicoCancel.Text = GenerarNemonico(strCodTipoInstrumentoOrden, strCodClaseInstrumento, strCodEmisor, txtNumDocDsctoCancel.Text)
    txtDescripOrden.Text = Trim$(cboTipoInstrumentoOrden.Text) & " - " & txtNemonico.Text
    txtDescripOrdenCancel.Text = Trim$(cboTipoInstrumentoOrden.Text) & " - " & txtNemonicoCancel.Text

    If strCodTipoInstrumentoOrden = Valor_Caracter Then Exit Sub
    If Not PosicionLimites() Then Exit Sub
    
    With adoComm
        Set adoRegistro = New ADODB.Recordset
                        
        .CommandText = "SELECT CodCategoriaRiesgo,CodRiesgoFinal,CodSubRiesgoFinal FROM EmisionInstitucionPersona " & "WHERE CodEmisor='" & strCodEmisor & "' AND CodFile='" & strCodTipoInstrumentoOrden & "' AND " & "CodDetalleFile='" & strCodClaseInstrumento & "'"
            
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            strCodRiesgo = Trim$(adoRegistro("CodRiesgoFinal"))
            strCodSubRiesgo = Trim$(adoRegistro("CodSubRiesgoFinal"))
        Else

            If strCodEmisor <> Valor_Caracter Then
              
                Exit Sub
            End If
        End If

        adoRegistro.Close
        
        .CommandText = "SELECT ValorParametro FROM AuxiliarParametro WHERE CodTipoParametro='TIPRIE' AND CodParametro='" & strCodRiesgo & "'"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            strCodCategoria = Trim$(adoRegistro("ValorParametro"))
        End If

        adoRegistro.Close: Set adoRegistro = Nothing
        
        lblClasificacion.Caption = strCodCategoria & Space$(1) & strCodSubRiesgo
    End With
    
    strCodObligado = strCodEmisor
    
End Sub

Private Function PosicionLimites() As Boolean

    PosicionLimites = False
        
    If cboTipoInstrumentoOrden.ListIndex <= 0 Then
        MsgBox "Debe seleccionar el Tipo de Instrumento.", vbCritical, Me.Caption
        cboEmisor.ListIndex = -1: cboTitulo.ListIndex = -1

        If cboTipoInstrumentoOrden.Enabled Then cboTipoInstrumentoOrden.SetFocus
        Exit Function
    End If

    PosicionLimites = True
    
End Function

Private Sub cboEstado_Click()

    strCodEstado = Valor_Caracter

    If cboEstado.ListIndex < 0 Then Exit Sub
    
    strCodEstado = Trim$(arrEstado(cboEstado.ListIndex))
    
    Call Buscar
End Sub

Public Sub setFondo(cf As String)
    blnFlag = True
    strCodFondoDescuento = cf
End Sub

Private Sub cboFondo_Click()

    Dim adoRegistro As ADODB.Recordset
    
    strCodFondo = Valor_Caracter

    If cboFondo.ListIndex < 0 Then Exit Sub
    
    strCodFondo = Trim$(arrFondo(cboFondo.ListIndex))
    
    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
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
    
    strSQL = "SELECT FIF.CodFile CODIGO,DescripFile DESCRIP " & "FROM FondoInversionFile FIF JOIN InversionFile IVF ON(IVF.CodFile=FIF.CodFile) " & "WHERE TipoValor='" & Codigo_Valor_RentaFija & "' AND TipoPlazo='" & Codigo_Valor_CortoPlazo & "' AND IndInstrumento='X' AND IndVigente='X' AND " & "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND FIF.CodFile = '" & CodFile_Descuento_Flujos_Dinerarios & "' ORDER BY DescripFile"
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
        .CommandText = "{ call up_ACSelDatosParametro(23,'" & strCodFondoOrden & "','" & gstrCodAdministradora & "','000') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            gdatFechaActual = CVDate(adoRegistro("FechaCuota"))
            gstrFechaActual = Convertyyyymmdd(gdatFechaActual)
            dtpFechaOrden.Value = gdatFechaActual
            dtpFechaOrdenCancel.Value = gdatFechaActual
            dtpFechaLiquidacion.Value = dtpFechaOrden.Value
            dtpFechaLiquidacionCancel.Value = dtpFechaOrden.Value
            dtpFechaVencimiento.Value = dtpFechaLiquidacion.Value
            strCodMoneda = Trim$(adoRegistro("CodMoneda"))
            txtTipoCambio.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, dtpFechaOrden.Value, Codigo_Moneda_Local, strCodMoneda))

            If CDbl(txtTipoCambio.Text) = 0 Then txtTipoCambio.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, DateAdd("d", -1, dtpFechaOrden.Value), Codigo_Moneda_Local, strCodMoneda))
            dblTipoCambio = CDbl(txtTipoCambio.Text)
                        
            frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
            Me.Refresh
        End If

        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    strSQL = "SELECT FIF.CodFile CODIGO,DescripFile DESCRIP " & "FROM FondoInversionFile FIF JOIN InversionFile IVF ON(IVF.CodFile=FIF.CodFile) " & "WHERE TipoValor='" & Codigo_Valor_RentaFija & "' AND TipoPlazo='" & Codigo_Valor_CortoPlazo & "' AND IndInstrumento='X' AND IndVigente='X' AND " & "CodFondo='" & strCodFondoOrden & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND FIF.CodFile = '" & CodFile_Descuento_Flujos_Dinerarios & "' ORDER BY DescripFile"
    CargarControlLista strSQL, cboTipoInstrumentoOrden, arrTipoInstrumentoOrden(), Sel_Defecto
        
End Sub

Private Sub cboGestor_Click()
    
    strCodGestor = Valor_Caracter

    If cboGestor.ListIndex < 0 Then Exit Sub
    
    strCodGestor = Trim$(arrGestor(cboGestor.ListIndex))

End Sub

Private Sub cboNegociacion_Click()

    strCodNegociacion = Valor_Caracter

    If cboNegociacion.ListIndex < 0 Then Exit Sub
    
    strCodNegociacion = Trim$(arrNegociacion(cboNegociacion.ListIndex))
            
    cboConceptoCosto.ListIndex = -1

    If cboConceptoCosto.ListCount > 0 Then cboConceptoCosto.ListIndex = 0
    
    cboConceptoCosto.Enabled = False

    If strCodNegociacion = Codigo_Mecanismo_Rueda Then cboConceptoCosto.Enabled = True
     
End Sub

Private Sub cboObligado_Click()

    strCodObligado = Valor_Caracter

    If cboObligado.ListIndex < 0 Then Exit Sub
    
    strCodObligado = Trim$(arrObligado(cboObligado.ListIndex))

End Sub

Private Sub cboCobroInteres_Click()

    strCodCobroInteres = Valor_Caracter

    If cboCobroInteres.ListIndex < 0 Then Exit Sub

    strCodCobroInteres = Mid$(Trim$(arrPagoInteres(cboCobroInteres.ListIndex)), 7, 2)

    If strCodCobroInteres = Codigo_Modalidad_Pago_Adelantado Then   'Si es pago de intereses adelantados permitir la edición de int. adicionales
        If (strCodTipoInstrumentoOrden = "015" And strCodClaseInstrumento = "001" And strCodSubClaseInstrumento = "001") Then   'Sòlo en caso de letras
        
            txtIntAdicional(0).Enabled = True
            chkDiasAdicional.Visible = True
            chkDiasAdicional.Value = Checked
        
            datFechaVctoAdicional = DateAdd("d", intDiasAdicionales, CVDate(dtpFechaVencimiento.Value))

            If Not EsDiaUtil(datFechaVctoAdicional) Then
                datFechaVctoAdicional = ProximoDiaUtil(datFechaVctoAdicional)
            End If

            lblFechaVencimientoAdic.Caption = datFechaVctoAdicional
            lblDiasAdic(0).Caption = "( " & CStr(DateDiff("d", dtpFechaVencimiento.Value, datFechaVctoAdicional)) & " días )"
            lblFechaVencimientoAdic.Visible = True
           
            txtDiasPlazo.Text = CStr(DateDiff("d", dtpFechaOrden.Value, dtpFechaVencimiento.Value))
            lblDiasPlazo.Caption = txtDiasPlazo.Text
            Call txtPrecioUnitario1_Change '(0)
            Call CalculoTotal(0)
            
        End If

    Else

        If (strCodTipoInstrumentoOrden = "015" And strCodClaseInstrumento = "001" And strCodSubClaseInstrumento = "001") Then   'Sòlo en caso de letras
            
            txtIntAdicional(0).Text = 0#
            txtIntAdicional(0).Enabled = False
            chkDiasAdicional.Value = Unchecked
            
            datFechaVctoAdicional = dtpFechaVencimiento.Value
            lblFechaVencimientoAdic.Caption = datFechaVctoAdicional
            lblFechaVencimientoAdic.Visible = False
            lblDiasAdic(0).Caption = 0
           
            txtDiasPlazo.Text = CStr(DateDiff("d", dtpFechaOrden.Value, dtpFechaVencimiento.Value))
            lblDiasPlazo.Caption = txtDiasPlazo.Text
            Call txtPrecioUnitario1_Change '(0)
            Call CalculoTotal(0)
           
        End If
    End If
    
End Sub

Private Sub cboResponsablePago_Click()

    strResponsablePago = Valor_Caracter

    If cboResponsablePago.ListIndex < 0 Then Exit Sub

    strResponsablePago = Trim$(arrResponsablePago(cboResponsablePago.ListIndex))

End Sub

Private Sub cboResponsablePagoCancel_Click()

    strResponsablePagoCancel = Valor_Caracter

    If cboResponsablePagoCancel.ListIndex < 0 Then Exit Sub

    strResponsablePagoCancel = Trim$(arrResponsablePagoCancel(cboResponsablePagoCancel.ListIndex))

End Sub

Private Sub cboSubClaseInstrumento_Click()

    Dim adoRegistro As ADODB.Recordset

    strCodSubClaseInstrumento = Valor_Caracter

    If cboSubClaseInstrumento.ListIndex < 0 Then Exit Sub
    
    strCodSubClaseInstrumento = Trim$(arrSubClaseInstrumento(cboSubClaseInstrumento.ListIndex))
    
    If strCodSubClaseInstrumento = "001" Then 'Con Interés
        strCalcVcto = "V"
        txtCantidad.Text = "1"
        cboTipoTasa.Enabled = False
        txtIntAdicional(0).Enabled = True
        
        txtTasa.Enabled = True
        txtPorcenDctoValorNominal.Text = dblPorcDescuento   'Obtenido de parámetros globales  '"100"
        txtPrecioUnitario1.Text = "100"
        txtPorcenDctoValorNominal.Enabled = False
        txtPorcenDctoValorNominal.Enabled = True
        txtPrecioUnitario1.Enabled = False
        txtTirBruta1.Enabled = False
        txtPorcenIgvInt(0).Locked = True
        txtPorcenIgv(0).Locked = True
    End If
    
    If strCodSubClaseInstrumento = "002" Then 'Al descuento
        strCalcVcto = "D"
        txtCantidad.Text = "1"
        txtCantidad.Enabled = False
        cboTipoTasa.Enabled = False
        txtTasa.Enabled = 0#
        txtTasa.Enabled = False
        txtIntAdicional(0).Text = 0#
        txtIntAdicional(0).Enabled = False
        txtPorcenDctoValorNominal.Text = "100"
        txtPrecioUnitario1.Text = "100"
        txtPorcenDctoValorNominal.Enabled = True
        txtPrecioUnitario1.Enabled = True
        txtTirBruta1.Enabled = True
        cboCobroInteres.ListIndex = 1
        txtPorcenIgvInt(0).Locked = False
        txtPorcenIgv(0).Locked = False
    End If
    
    intDiasAdicionales = 0

    If strCodTipoInstrumentoOrden = "015" And strCodClaseInstrumento = "001" And strCodSubClaseInstrumento = "001" Then

        chkDiasAdicional.Visible = True
        chkDiasAdicional.Value = Checked
        chkDiasAdicional_Click
        cboCobroInteres.Enabled = True
     
        Set adoRegistro = New ADODB.Recordset
        adoComm.CommandText = "SELECT CONVERT(int,ValorParametro) AS DiasAdicionales FROM ParametroGeneral WHERE CodParametro = '20'"
        Set adoRegistro = adoComm.Execute

        If Not (adoRegistro.EOF) Then
            intDiasAdicionales = adoRegistro("DiasAdicionales")
        End If

        If intDiasAdicionales = Null Then
            intDiasAdicionales = 0
        End If

        adoRegistro.Close: Set adoRegistro = Nothing
    
    Else

        If strCodTipoInstrumentoOrden <> "010" And strCodTipoInstrumentoOrden <> "016" And strCodTipoInstrumentoOrden <> "021" Then  'Que no sean letras por maquinarias pues ellas son con interés al vencimiento
            cboCobroInteres.ListIndex = ObtenerItemLista(arrPagoInteres(), "MODPAG" + Codigo_Modalidad_Pago_Adelantado)
            chkDiasAdicional.Value = Unchecked
            chkDiasAdicional.Visible = False
            cboCobroInteres.Enabled = False
        End If
    End If
    
    Call cboTipoOrden_Click
    
End Sub

Private Sub cboOrigen_Click()

    strCodOrigen = Valor_Caracter

    If cboOrigen.ListIndex < 0 Then Exit Sub
    
    strCodOrigen = Trim$(arrOrigen(cboOrigen.ListIndex))
    
End Sub

Private Sub cboTipoInstrumento_Click()

    strCodTipoInstrumento = Valor_Caracter

    If cboTipoInstrumento.ListIndex < 0 Then Exit Sub
    
    strCodTipoInstrumento = Trim$(arrTipoInstrumento(cboTipoInstrumento.ListIndex))

    strSQL = "SELECT CodLimite CODIGO, DescripLimite DESCRIP FROM LimiteReglamentoEstructuraDetalle where CodEstructura = '" & Codigo_LimiteRE_Cliente & "'" & _
             " and CodLimite = '" & Linea_Contrato_Flujo_Dinerario & "' and Estado  = '01' "
    CargarControlLista strSQL, cboLineaCliente, arrLineaCliente(), ""
    Call cboLineaCliente_Click   'Para obligar a que se seleccione el único elemento de la lista

    If cboLineaClienteLista.ListCount > 0 Then cboLineaClienteLista.ListIndex = 0
    
    Call Buscar
    
End Sub
Private Sub cboLineaCliente_Click()

   Dim adoRegistro As New ADODB.Recordset

    strLineaCliente = Valor_Caracter
    strTipoPersonaLim = Valor_Caracter
    strCodPersonaLim = Valor_Caracter
    
    If cboLineaCliente.ListIndex < 0 Then Exit Sub
    
    strLineaCliente = Trim(arrLineaCliente(cboLineaCliente.ListIndex))
    
    strTipoPersonaLim = Codigo_Tipo_Persona_Emisor
    If strLineaCliente = Linea_Financiamiento_Proveedores Then
        strCodPersonaLim = strCodObligado
        strTipoPersonaLim = Codigo_Tipo_Persona_Obligado
    Else
        strCodPersonaLim = strCodEmisor
        strTipoPersonaLim = Codigo_Tipo_Persona_Emisor
    End If
    
End Sub
Private Sub cboLineaClienteLista_Click()

    strLineaClienteLista = Valor_Caracter
    If cboLineaClienteLista.ListIndex < 0 Then Exit Sub
    
    strLineaClienteLista = Trim(arrLineaClienteLista(cboLineaClienteLista.ListIndex))

    Call Buscar

End Sub
Private Sub cboTipoInstrumentoOrden_Click()
    
    strCodTipoInstrumentoOrden = Valor_Caracter
    strIndPacto = Valor_Caracter: strIndNegociable = Valor_Caracter

    If cboTipoInstrumentoOrden.ListIndex < 0 Then Exit Sub
    
    strCodTipoInstrumentoOrden = Trim$(arrTipoInstrumentoOrden(cboTipoInstrumentoOrden.ListIndex))

    If strCodTipoInstrumentoOrden = "010" Or strCodTipoInstrumentoOrden = "016" Or strCodTipoInstrumentoOrden = "021" Then   'Letras o flujos dinerarios
        cboCobroInteres.ListIndex = ObtenerItemLista(arrPagoInteres(), "MODPAG" + Codigo_Modalidad_Pago_Vencimiento)
        cboCobroInteres.Enabled = False
    Else
        cboCobroInteres.ListIndex = ObtenerItemLista(arrPagoInteres(), "MODPAG" + Codigo_Modalidad_Pago_Adelantado)
        cboCobroInteres.Enabled = True
    End If
    
    txtNemonico.Text = GenerarNemonico(strCodTipoInstrumentoOrden, strCodClaseInstrumento, strCodEmisor, txtNumDocDsctoCancel.Text)
    txtDescripOrdenCancel.Text = Trim$(cboTipoInstrumentoOrden.Text) & " - " & txtNemonicoCancel.Text
    
    strSQL = "SELECT IFTON.CodTipoOperacion CODIGO,DescripParametro DESCRIP " & "FROM InversionFileTipoOperacionNegociacion IFTON JOIN AuxiliarParametro AUX ON(AUX.CodParametro=IFTON.CodTipoOperacion AND AUX.CodTipoParametro = 'OPECAJ') " & "WHERE IFTON.CodFile='" & strCodTipoInstrumentoOrden & "' ORDER BY CodTipoOperacion"
    CargarControlLista strSQL, cboTipoOrden, arrTipoOrden(), Valor_Caracter

    If cboTipoOrden.ListCount > 0 Then
        cboTipoOrden.ListIndex = 0
    End If
    
    lblAnalitica.Caption = strCodTipoInstrumentoOrden & " - ????????"
    strCodFile = strCodTipoInstrumentoOrden

    strSQL = "SELECT CodDetalleFile CODIGO,DescripDetalleFile DESCRIP FROM InversionDetalleFile WHERE CodFile='" & strCodTipoInstrumentoOrden & "' AND IndVigente='X' ORDER BY DescripDetalleFile"
    CargarControlLista strSQL, cboClaseInstrumento, arrClaseInstrumento(), Sel_Defecto
    
    If cboClaseInstrumento.ListCount > 0 Then
        cboClaseInstrumento.ListIndex = 0
        cboClaseInstrumento.Enabled = True
    End If
    
    cboLineaCliente.Clear
            
End Sub

Private Sub cboMoneda_Click()
    
    lblDescripMoneda(0).Caption = "S/.": lblDescripMoneda(0).Tag = Codigo_Moneda_Local
    lblDescripMoneda(1).Caption = "S/.": lblDescripMoneda(1).Tag = Codigo_Moneda_Local
    lblDescripMonedaResumen(0) = "S/.": lblDescripMonedaResumen(0).Tag = Codigo_Moneda_Local
    lblDescripMonedaResumen(1) = "S/.": lblDescripMonedaResumen(1).Tag = Codigo_Moneda_Local
    
    If cboMoneda.ListIndex < 0 Then Exit Sub
    
    strCodMoneda = Trim$(arrMoneda(cboMoneda.ListIndex))
        
    lblDescripMoneda(0).Caption = ObtenerSignoMoneda(strCodMoneda)
    lblDescripMoneda(0).Tag = strCodMoneda
    lblDescripMoneda(1).Caption = ObtenerSignoMoneda(strCodMoneda)
    lblDescripMoneda(1).Tag = strCodMoneda
    lblDescripMonedaResumen(0).Caption = ObtenerSignoMoneda(strCodMoneda)
    lblDescripMonedaResumen(0).Tag = strCodMoneda
    lblDescripMonedaResumen(1).Caption = ObtenerSignoMoneda(strCodMoneda)
    lblDescripMonedaResumen(1).Tag = strCodMoneda
    lblMoneda.Caption = ObtenerDescripcionMoneda(strCodMoneda)
    
    Call AsignarComisionOperacion
    
End Sub

Private Sub cboOperacion_Click()

    strCodOperacion = Valor_Caracter

    If cboOperacion.ListIndex < 0 Then Exit Sub
    
    strCodOperacion = Trim$(arrOperacion(cboOperacion.ListIndex))
    
End Sub

Public Sub CargarComisiones(ByVal strCodComision As String, Index As Integer)
     
    Call AplicarCostos(Index)
     
End Sub

Private Sub cboTipoOrden_Click()

    Dim adoRegistro As ADODB.Recordset
    
    strCodTipoOrden = Valor_Caracter

    If cboTipoOrden.ListIndex < 0 Then Exit Sub

    strCodTipoOrden = Trim$(arrTipoOrden(cboTipoOrden.ListIndex))
    blnCancelaPrepago = False
    
    Me.MousePointer = vbHourglass

    Select Case strCodTipoOrden
    
        Case Codigo_Orden_Compra
            chkTitulo.Enabled = True
            cboTitulo.Visible = False: cboEmisor.Visible = True
            lblDescrip(6) = "Emisor"
            
            fraDatosTituloCancel.Visible = False
            fraDatosAnexo.Visible = True
            fraDatosTitulo.Visible = True
            fraResumen.Visible = True
 
            If chkTitulo.Value Then
                strSQL = "SELECT CodTitulo CODIGO,(Nemotecnico + ' ' + DescripTitulo)DESCRIP FROM InstrumentoInversion " & "WHERE CodFile='" & strCodFile & "' AND CodDetalleFile='" & strCodClaseInstrumento & "' AND IndVigente='X' ORDER BY DescripTitulo"
                CargarControlLista strSQL, cboTitulo, arrTitulo(), Sel_Defecto
            
                If cboTitulo.ListCount > 0 Then cboTitulo.ListIndex = 0
            End If
            
            fraComisionMontoFL2.Visible = False
        
        Case Codigo_Orden_Renovacion
            chkTitulo.Enabled = True
            cboTitulo.Visible = False: cboEmisor.Visible = True
            lblDescrip(6) = "Emisor"
            
            fraDatosTituloCancel.Visible = False
            fraDatosAnexo.Visible = True
            fraDatosTitulo.Visible = True
            fraResumen.Visible = True
 
            If chkTitulo.Value Then
                strSQL = "SELECT CodTitulo CODIGO,(Nemotecnico + ' ' + DescripTitulo)DESCRIP FROM InstrumentoInversion " & "WHERE CodFile='" & strCodFile & "' AND CodDetalleFile='" & strCodClaseInstrumento & "' AND IndVigente='X' ORDER BY DescripTitulo"
                CargarControlLista strSQL, cboTitulo, arrTitulo(), Sel_Defecto
            
                If cboTitulo.ListCount > 0 Then cboTitulo.ListIndex = 0
            End If
            
            fraComisionMontoFL2.Visible = False

        Case Codigo_Orden_Venta
            chkTitulo.Enabled = False
            cboTitulo.Visible = True: cboEmisor.Visible = False
            lblDescrip(6) = "Título"
            fraDatosTituloCancel.Visible = False
            fraDatosAnexo.Visible = True
            fraDatosTitulo.Visible = True
            fraResumen.Visible = True
 
            strSQL = "SELECT InstrumentoInversion.CodTitulo CODIGO," & "(RTRIM(InstrumentoInversion.Nemotecnico) + ' ' + RTRIM(InstrumentoInversion.DescripTitulo)) DESCRIP FROM InstrumentoInversion,InversionKardex " & "WHERE SaldoFinal > 0 AND IndUltimoMovimiento='X' AND InstrumentoInversion.CodFile=InversionKardex.CodFile AND " & "InstrumentoInversion.CodAnalitica=InversionKardex.CodAnalitica AND InversionKardex.CodFile='" & strCodFile & "' AND " & "InstrumentoInversion.CodFondo='" & strCodFondoOrden & "' AND InversionKardex.CodFondo='" & strCodFondoOrden & "' " & "ORDER BY InstrumentoInversion.Nemotecnico"
            CargarControlLista strSQL, cboTitulo, arrTitulo(), Sel_Defecto
        
            If cboTitulo.ListCount > 0 Then cboTitulo.ListIndex = 0

            fraComisionMontoFL2.Visible = False
            
        Case Codigo_Orden_Pacto
            chkTitulo.Enabled = True
            cboTitulo.Visible = False: cboEmisor.Visible = True
            lblDescrip(6) = "Emisor"
            fraDatosTituloCancel.Visible = False
            fraDatosAnexo.Visible = True
            fraDatosTitulo.Visible = True
            fraResumen.Visible = True
           
            If chkTitulo.Value Then
                strSQL = "SELECT CodTitulo CODIGO,(Nemotecnico + ' ' + DescripTitulo)DESCRIP FROM InstrumentoInversion " & "WHERE CodFile='" & strCodFile & "' AND IndVigente='X' ORDER BY DescripTitulo"
                CargarControlLista strSQL, cboTitulo, arrTitulo(), Sel_Defecto
            
                If cboTitulo.ListCount > 0 Then cboTitulo.ListIndex = 0
            End If
            
            fraComisionMontoFL2.Visible = True
            
        Case Codigo_Orden_PagoCancelacion, Codigo_Orden_Prepago
            chkTitulo.Enabled = True
            cboTitulo.Visible = False: cboEmisor.Visible = True
            lblDescrip(6) = "Emisor"
            fraDatosTituloCancel.Visible = True
            fraDatosAnexo.Visible = False
            fraDatosTitulo.Visible = False
            fraResumen.Visible = False
            
            fraComisionMontoFL2.Visible = False
            blnCancelaPrepago = True
 
    End Select
    
    strCodFile = strCodTipoInstrumentoOrden
    
    Me.MousePointer = vbDefault
    
End Sub

Private Sub cboTipoTasa_Click()

    strCodTipoTasa = Valor_Caracter

    If cboTipoTasa.ListIndex < 0 Then Exit Sub
    
    strCodTipoTasa = Trim$(arrTipoTasa(cboTipoTasa.ListIndex))
    
End Sub

Private Sub cboTitulo_Click()

    Dim adoRegistro As ADODB.Recordset
    Dim intRegistro As Integer
    
    strCodGarantia = Valor_Caracter: txtDescripOrden = Valor_Caracter: txtDescripOrdenCancel = Valor_Caracter: strCodAnalitica = Valor_Caracter
    lblAnalitica = strCodTipoInstrumentoOrden & "-????????":
    lblStockNominal = "0"
    strCodEmisor = Valor_Caracter: strCodGrupo = Valor_Caracter

    If cboTitulo.ListIndex < 0 Then Exit Sub

    strCodGarantia = Trim$(arrTitulo(cboTitulo.ListIndex))

    With adoComm
        Set adoRegistro = New ADODB.Recordset

        .CommandText = "SELECT CodAnalitica,ValorNominal,CodMoneda,CodEmisor,CodGrupo,FechaEmision,FechaVencimiento," & "TasaInteres,CodRiesgo,CodSubRiesgo,CodTipoTasa,BaseAnual,Nemotecnico " & "FROM InstrumentoInversion WHERE CodTitulo='" & strCodGarantia & "'"
        Set adoRegistro = .Execute

        If Not adoRegistro.EOF Then
            strCodAnalitica = Trim$(adoRegistro("CodAnalitica"))
            lblAnalitica.Caption = strCodFile & "-" & strCodAnalitica
                        
            intRegistro = ObtenerItemLista(arrMoneda(), adoRegistro("CodMoneda"))

            If intRegistro >= 0 Then cboMoneda.ListIndex = intRegistro
            
            dtpFechaVencimiento.Value = adoRegistro("FechaVencimiento")
            dtpFechaVencimiento_Change
            txtNemonico.Text = Trim$(adoRegistro("Nemotecnico"))
            txtNemonicoCancel.Text = Trim$(adoRegistro("Nemotecnico"))
            
            intRegistro = ObtenerItemLista(arrTipoTasa(), adoRegistro("CodTipoTasa"))

            If intRegistro >= 0 Then cboTipoTasa.ListIndex = intRegistro
            
            intRegistro = ObtenerItemLista(arrBaseAnual(), adoRegistro("BaseAnual"))

            If intRegistro >= 0 Then cboBaseAnual.ListIndex = intRegistro
            
            txtTasa.Text = adoRegistro("TasaInteres")
            txtValorNominal.Text = CStr(adoRegistro("ValorNominal"))
                
            strCodEmisor = Trim$(adoRegistro("CodEmisor")): strCodGrupo = Trim$(adoRegistro("CodGrupo"))
            strCodRiesgo = Trim$(adoRegistro("CodRiesgo"))
            strCodSubRiesgo = Trim$(adoRegistro("CodSubRiesgo"))
            lblMoneda.Caption = ObtenerDescripcionMoneda(adoRegistro("CodMoneda"))
            lblBaseTasaCupon.Caption = "360" & Space$(1) & "-" & Space$(1) & Trim$(txtTasa.Text) & "%"

            If adoRegistro("BaseAnual") = Codigo_Base_Actual_Actual Then lblBaseTasaCupon.Caption = "365" & Space$(1) & "-" & Space$(1) & Trim$(txtTasa.Text) & "%"
            If adoRegistro("BaseAnual") = Codigo_Base_Actual_365 Then lblBaseTasaCupon.Caption = "365" & Space$(1) & "-" & Space$(1) & Trim$(txtTasa.Text) & "%"
            If adoRegistro("BaseAnual") = Codigo_Base_30_365 Then lblBaseTasaCupon.Caption = "365" & Space$(1) & "-" & Space$(1) & Trim$(txtTasa.Text) & "%"
            
            cboMoneda.Enabled = False
            cboTipoTasa.Enabled = False
            cboBaseAnual.Enabled = False
            dtpFechaVencimiento.Enabled = False
            txtDiasPlazo.Enabled = False
            txtValorNominal.Enabled = False
            txtTasa.Enabled = False
            txtNemonico.Enabled = False
            txtNemonicoCancel.Enabled = False
        End If

        adoRegistro.Close

        .CommandText = "SELECT FechaPago " & "FROM InstrumentoInversionCalendario WHERE CodTitulo='" & strCodGarantia & "'"
        Set adoRegistro = .Execute

        If Not adoRegistro.EOF Then
            dtpFechaPago.Value = adoRegistro("FechaPago")
            dtpFechaPago_Change
            dtpFechaPago.Enabled = False
        End If

        adoRegistro.Close
        
        .CommandText = "SELECT ValorParametro FROM AuxiliarParametro WHERE CodTipoParametro='TIPRIE' AND CodParametro='" & strCodRiesgo & "'"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            strCodCategoria = Trim$(adoRegistro("ValorParametro"))
        End If

        adoRegistro.Close
        
        lblClasificacion.Caption = strCodCategoria & Space$(1) & strCodSubRiesgo
        
        If Not PosicionLimites() Then Exit Sub

        .CommandText = "SELECT SaldoFinal,ValorPromedio FROM InversionKardex WHERE CodAnalitica='" & strCodAnalitica & "' AND " & "CodFile='" & strCodFile & "' AND CodFondo='" & strCodFondoOrden & "' AND " & "CodAdministradora='" & gstrCodAdministradora & "' AND IndUltimoMovimiento='X' AND SaldoFinal > 0"
            
        Set adoRegistro = .Execute

        If Not adoRegistro.EOF Then
            lblStockNominal.Caption = CStr(adoRegistro("SaldoFinal"))
        End If

        adoRegistro.Close: Set adoRegistro = Nothing

    End With

    txtDescripOrden = Trim$(cboTipoInstrumentoOrden.Text) & " - " & Left$(cboTitulo.Text, 15)
        
End Sub

Private Sub cboViaCobranza_Click()

    strViaCobranza = Valor_Caracter

    If cboViaCobranza.ListIndex < 0 Then Exit Sub
    
    strViaCobranza = Trim$(arrViaCobranza(cboViaCobranza.ListIndex))

End Sub

Private Sub chkAplicar_Click(Index As Integer)

    If chkAplicar(Index).Value Then
        fraComisiones.Enabled = True
        Call CalcularComision
        Call AplicarCostos(Index)
    Else
        Call IniciarComisiones
        fraComisiones.Enabled = False
    End If

    Call CalculoTotal(Index)
  
End Sub

Private Sub chkDiasAdicional_Click()

    If chkDiasAdicional.Visible = True Then

        If chkDiasAdicional.Value = Checked Then
            
            datFechaVctoAdicional = DateAdd("d", intDiasAdicionales, CVDate(dtpFechaVencimiento.Value))

            If Not EsDiaUtil(datFechaVctoAdicional) Then
                datFechaVctoAdicional = ProximoDiaUtil(datFechaVctoAdicional)
            End If

            lblFechaVencimientoAdic.Caption = datFechaVctoAdicional
            lblDiasAdic(0).Caption = "( " & CStr(DateDiff("d", dtpFechaVencimiento.Value, datFechaVctoAdicional)) & " días )"
            lblFechaVencimientoAdic.Visible = True
            
        Else

            datFechaVctoAdicional = dtpFechaVencimiento.Value
            lblFechaVencimientoAdic.Caption = datFechaVctoAdicional
            lblFechaVencimientoAdic.Visible = False
            lblDiasAdic(0).Caption = 0
            
        End If
    
        txtDiasPlazo.Text = CStr(DateDiff("d", dtpFechaOrden.Value, dtpFechaVencimiento.Value))
        lblDiasPlazo.Caption = txtDiasPlazo.Text
        Call txtPrecioUnitario1_Change '(0)
        Call CalculoTotal(0)
    
    End If

End Sub

Private Sub chkInteresCorrido_Click(Index As Integer)

    If chkInteresCorrido(Index).Value Then
        txtInteresCorrido(Index).Text = CalculoInteresDescuento(CDbl(txtTasa.Text), strCodTipoTasa, "01", strCodBaseAnual, txtValorNominalDcto, CStr(dtpFechaLiquidacion.Value), CStr(dtpFechaOrden.Value))
    Else
        txtInteresCorrido(Index).Text = 0#
    End If

    txtImptoInteresCorrido(Index).Text = (txtInteresCorrido(Index).Text * txtPorcenIgvInt(Index).Text / 100)
    Call CalculoTotal(Index)

End Sub

Private Sub chkTitulo_Click()

    If chkTitulo.Value = 1 Then
        cboTitulo.Visible = True: cboEmisor.Visible = False
        lblDescrip(6) = "Título"
        
        Me.MousePointer = vbHourglass

        Select Case strCodTipoOrden

            Case Codigo_Orden_Compra, Codigo_Orden_Renovacion, Codigo_Orden_Pacto
                strSQL = "SELECT CodTitulo CODIGO,(Nemotecnico + ' ' + DescripTitulo)DESCRIP FROM InstrumentoInversion " & "WHERE CodFile='" & strCodFile & "' AND CodDetalleFile='" & strCodClaseInstrumento & "' AND IndVigente='X' " & "ORDER BY DescripTitulo"
                CargarControlLista strSQL, cboTitulo, arrTitulo(), Sel_Defecto
                            
            Case Codigo_Orden_Venta
                strSQL = "SELECT InstrumentoInversion.CodTitulo CODIGO," & "(RTRIM(InstrumentoInversion.CodTitulo) + ' ' + RTRIM(InstrumentoInversion.Nemotecnico) + ' ' + RTRIM(InstrumentoInversion.DescripTitulo)) DESCRIP " & "FROM InstrumentoInversion,InversionKardex " & "WHERE SaldoFinal > 0 AND IndUltimo='X' AND InstrumentoInversion.CodFile=InversionKardex.CodFile AND " & "InstrumentoInversion.CodAnalitica=InversionKardex.CodAnalitica AND InversionKardex.CodFile='" & strCodFile & "' AND " & "(InstrumentoInversion.CodFondo='" & strCodFondoOrden & "' OR InstrumentoInversion.CodFondo='') AND " & "(InstrumentoInversion.CodAdministradora='" & gstrCodAdministradora & "' OR InstrumentoInversion.CodAdministradora='') AND " & "InversionKardex.CodFondo='" & strCodFondoOrden & "' " & "ORDER BY InstrumentoInversion.Nemotecnico"
                CargarControlLista strSQL, cboTitulo, arrTitulo(), Sel_Defecto
                        
        End Select

        Me.MousePointer = vbDefault

        If cboTitulo.ListCount > 0 Then cboTitulo.ListIndex = 0
                
    Else
        cboTitulo.Visible = False: cboEmisor.Visible = True
        lblDescrip(6).Caption = "Emisor"

        If cboEmisor.ListCount > 0 Then cboEmisor.ListIndex = 0
        
        cboMoneda.Enabled = True
        cboTipoTasa.Enabled = True
        cboBaseAnual.Enabled = True
        dtpFechaVencimiento.Enabled = True
        dtpFechaPago.Enabled = True
        txtDiasPlazo.Enabled = True
        txtValorNominal.Enabled = True
        txtTasa.Enabled = True
        txtNemonico.Enabled = True
        txtNemonicoCancel.Enabled = True
    End If
        
End Sub

Private Sub cmdCalculo_Click()

    Call CalcularTirBruta
    
End Sub

Private Sub cmdEnviar_Click()

    Dim strFechaDesde As String, strFechaHasta        As String
    Dim intRegistro   As Integer, intContador         As Integer
    Dim datFecha      As Date
    
    If adoConsulta.Recordset.RecordCount = 0 Then Exit Sub
    
    strFechaDesde = Convertyyyymmdd(dtpFechaOrdenDesde.Value)
    datFecha = DateAdd("d", 1, dtpFechaOrdenHasta.Value)
    strFechaHasta = Convertyyyymmdd(datFecha)
    
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
    Else
        MsgBox Mensaje_Desenvio_Exitoso, vbExclamation, gstrNombreEmpresa
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

    lblFechaLiquidacion.Caption = CStr(dtpFechaLiquidacion.Value)
    
    If strCodTipoInstrumentoOrden = "015" Then
        txtDiasPlazo.Text = CStr(DateDiff("d", dtpFechaOrden.Value, dtpFechaVencimiento.Value))
        lblDiasPlazo.Caption = txtDiasPlazo.Text
        Call CalculoTotal(0)
    End If
    
End Sub

Private Sub dtpFechaLiquidacionCancel_Change()

    If dtpFechaLiquidacionCancel.Value < dtpFechaOrdenCancel.Value Then
        dtpFechaLiquidacionCancel.Value = dtpFechaOrdenCancel.Value
    End If
        
    If Not EsDiaUtil(dtpFechaLiquidacionCancel.Value) Then
        MsgBox "La Fecha no es un día útil...se cambiará por una fecha correcta !", vbInformation, Me.Caption
        dtpFechaLiquidacionCancel.Value = ProximoDiaUtil(dtpFechaLiquidacionCancel.Value)
    End If

End Sub

Private Sub dtpFechaLiquidacionDesde_Click()

    If IsNull(dtpFechaLiquidacionDesde.Value) Then
        dtpFechaLiquidacionHasta.Value = Null
    Else
        dtpFechaLiquidacionDesde.Value = gdatFechaActual
        dtpFechaLiquidacionHasta.Value = dtpFechaLiquidacionDesde.Value
    End If
    
End Sub

Private Sub dtpFechaLiquidacionHasta_Click()

    If IsNull(dtpFechaLiquidacionHasta.Value) Then
        dtpFechaLiquidacionDesde.Value = Null
    Else
        dtpFechaLiquidacionDesde.Value = gdatFechaActual
        dtpFechaLiquidacionHasta.Value = dtpFechaLiquidacionDesde.Value
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

Private Sub dtpFechaPago_Change()

    If dtpFechaPago.Value < dtpFechaVencimiento.Value Then
        dtpFechaPago.Value = dtpFechaVencimiento.Value
    End If
    
    If Not EsDiaUtil(dtpFechaPago.Value) Then
        MsgBox "La Fecha de Pago no es un día útil...se cambiará por una fecha correcta !", vbInformation, Me.Caption
        dtpFechaPago.Value = ProximoDiaUtil(dtpFechaPago.Value)
        dtpFechaVencimiento.Value = dtpFechaPago.Value
    End If
    
    Call CalculoTotal(0)
    
End Sub

Private Sub dtpFechaVencimiento_Change()

    If dtpFechaVencimiento.Value < dtpFechaOrden.Value Then
        dtpFechaVencimiento.Value = dtpFechaOrden.Value
    End If
    
    If dtpFechaVencimiento.Value < dtpFechaLiquidacion.Value Then
        dtpFechaVencimiento.Value = dtpFechaLiquidacion.Value
    End If
    
    dtpFechaPago.Value = dtpFechaVencimiento.Value
    
    txtDiasPlazo.Text = CStr(DateDiff("d", dtpFechaOrden.Value, dtpFechaVencimiento.Value))
    lblDiasPlazo.Caption = txtDiasPlazo.Text

    Call txtPrecioUnitario1_Change '(0)
    Call CalculoTotal(0)
    
    lblFechaVencimiento.Caption = CStr(dtpFechaVencimiento.Value)
    lblFechaCupon.Caption = CStr(dtpFechaVencimiento.Value)
    
End Sub

Private Sub Form_Activate()

    frmMainMdi.stbMdi.Panels(3).Text = Me.Caption
    Call CargarReportes
    
End Sub

Private Sub Form_Deactivate()

    Call OcultarReportes
    
End Sub

Private Sub Form_Load()

    indCargaPantalla = True
    blnCargadoDesdeCartera = False
    
    Call InicializarValores
    Call CargarListas
    Call CargarReportes
    Call Buscar
        
    Call ValidarPermisoUsoControl(Trim$(gstrLogin), Me, Trim$(App.Title) + Separador_Codigo_Objeto + gstrNombreObjetoMenuPulsado + Separador_Codigo_Objeto + Me.Name, Separador_Codigo_Objeto)

    If blnFlag Then
        cboFondo.ListIndex = 0
    End If
    
    CentrarForm Me
    indCargaPantalla = False
            
    Call ValidaExisteTipoCambio(Codigo_TipoCambio_SBS, gstrFechaActual)
    
End Sub


Public Sub Buscar()

    Dim strFechaOrdenDesde       As String, strFechaOrdenHasta        As String
    Dim strFechaLiquidacionDesde As String, strFechaLiquidacionHasta  As String
    Dim datFechaSiguiente        As Date

    Me.MousePointer = vbHourglass
    
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
        
    strSQL = "SELECT IOR.NumOrden,FechaOrden,FechaLiquidacion,CodTitulo,Nemotecnico,EstadoOrden,IOR.CodFile,CodAnalitica,TipoOrden,IOR.CodMoneda," & _
       "(RTRIM(DescripParametro) + Space(1) + DescripOrden) DescripOrden,CantOrden,ValorNominal,PrecioUnitarioMFL1,MontoTotalMFL1, " & _
       "CodSigno DescripMoneda, IOR.NumAnexo, NumDocumentoFisico,IOR.CodDetalleFile, IOR.CodSubDetalleFile, IOR.CodFondo, IOR.CodGirador, " & _
       "IP1.DescripPersona DesGirador, IOR.CodObligado, IP1.DescripPersona DesObligado, IOR.CodGestor, IP3.DescripPersona DesGestor " & _
       "FROM InversionOrden IOR JOIN AuxiliarParametro AUX ON(AUX.CodParametro=IOR.TipoOrden AND AUX.CodTipoParametro = 'OPECAJ') " & _
       "JOIN Moneda MON ON(MON.CodMoneda=IOR.CodMoneda) " & _
       "Left JOIN InstitucionPersona IP1 ON (IP1.CodPersona = IOR.CodGirador AND IP1.TipoPersona = '" & Codigo_Tipo_Persona_Emisor & "') " & _
       " /* Left JOIN InstitucionPersona IP2 ON (IP2.CodPersona = IOR.CodObligado AND IP2.TipoPersona = '" & Codigo_Tipo_Persona_Contratante & "') */ " & _
       "Left JOIN InstitucionPersona IP3 ON (IP3.CodPersona = IOR.CodGestor AND IP3.TipoPersona = '" & Codigo_Tipo_Persona_Emisor & "') " & _
       "WHERE IOR.CodFile = '" & CodFile_Descuento_Flujos_Dinerarios & "' AND IOR.CodAdministradora='" & gstrCodAdministradora & "' AND IOR.CodFondo='" & strCodFondo & "' "
        
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

Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Ordenes de Inversión"
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo2").Text = "Papeleta de Inversión"
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo3").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo3").Text = "Anexo"
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo4").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo4").Text = "Anexo Cliente"
    
End Sub

Private Sub CargarListas()

    Dim intRegistro As Integer
    
     '*** Fondos ***
    strSQL = "SELECT CodFondo CODIGO,DescripFondo DESCRIP FROM Fondo WHERE CodAdministradora='" & gstrCodAdministradora & "' AND Estado='01' and CodFondo = '" & gstrCodFondoContable & "' ORDER BY DescripFondo"
    CargarControlLista strSQL, cboFondo, arrFondo(), Valor_Caracter
    CargarControlLista strSQL, cboFondoOrden, arrFondoOrden(), Valor_Caracter

    If cboFondo.ListCount > 0 Then
        
        If strCodFondoDescuento <> Valor_Caracter Then
            intRegistro = ObtenerItemLista(arrFondoOrden(), strCodFondoDescuento)

            If intRegistro >= 0 Then cboFondoOrden.ListIndex = intRegistro
        
            intRegistro = ObtenerItemLista(arrFondo(), strCodFondoDescuento)

            If intRegistro >= 0 Then cboFondo.ListIndex = intRegistro
        Else
            cboFondo.ListIndex = 0
            cboFondoOrden.ListIndex = 0
        End If
    End If
    
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='ESTORD' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboEstado, arrEstado(), Sel_Todos
    
    intRegistro = ObtenerItemLista(arrEstado(), Estado_Orden_Ingresada)

    If intRegistro >= 0 Then cboEstado.ListIndex = intRegistro
        
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='OPECAJ' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboTipoOrden, arrTipoOrden(), Valor_Caracter
    
    strSQL = "SELECT (CodPersona) CODIGO,DescripPersona DESCRIP FROM InstitucionPersona WHERE TipoPersona='" & Codigo_Tipo_Persona_Emisor & "' AND IndVigente='X' ORDER BY DescripPersona"
    CargarControlLista strSQL, cboEmisor, arrEmisor(), Sel_Defecto

    strSQL = "SELECT (CodPersona) CODIGO,DescripPersona DESCRIP FROM InstitucionPersona WHERE TipoPersona='" & Codigo_Tipo_Persona_Obligado & "' AND IndVigente='X' ORDER BY DescripPersona"
    CargarControlLista strSQL, cboObligado, arrObligado(), Sel_Defecto

    strSQL = "SELECT (CodPersona) CODIGO,DescripPersona DESCRIP FROM InstitucionPersona WHERE TipoPersona='" & Codigo_Tipo_Persona_Emisor & "' AND IndVigente='X' and IndBanco = 'X' ORDER BY DescripPersona"
    CargarControlLista strSQL, cboGestor, arrGestor(), Sel_Defecto
    
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='MDONEG' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboOrigen, arrOrigen(), Valor_Caracter
            
    strSQL = "{ call up_ACSelDatos(2) }"
    CargarControlLista strSQL, cboMoneda, arrMoneda(), Sel_Defecto
    
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='BASANU' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboBaseAnual, arrBaseAnual(), Valor_Caracter
    
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='NATTAS' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboTipoTasa, arrTipoTasa(), ""
    
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPLIQ' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboOperacion, arrOperacion(), Valor_Caracter
    
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='MECNEG' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboNegociacion, arrNegociacion(), Valor_Caracter

    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPCCO' AND ValorParametro='RF' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboConceptoCosto, arrConceptoCosto(), Sel_Defecto
    
    strSQL = "SELECT (CodTipoParametro + CodParametro) CODIGO, DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='MODPAG' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboCobroInteres, arrPagoInteres(), ""

    If cboCobroInteres.ListCount > 0 Then cboCobroInteres.ListIndex = 0
    
    indCargaPantalla = True
    
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='RESPAG' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboResponsablePago, arrResponsablePago(), Sel_Defecto

    If cboResponsablePago.ListCount > 0 Then cboResponsablePago.ListIndex = 0

    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='RESPAG' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboResponsablePagoCancel, arrResponsablePagoCancel(), Valor_Caracter

    If cboResponsablePagoCancel.ListCount > 0 Then cboResponsablePagoCancel.ListIndex = 0

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

        Case vSave
            Call GrabarNew

        Case vCancel
            blnCancelaPrepago = False
            Call Cancelar

        Case vExit
            Call Salir
        
    End Select
    
End Sub

Private Sub InicializarValores()
    
    Dim adoRegistro As ADODB.Recordset
    
    strEstado = Reg_Defecto
    tabRFCortoPlazo.Tab = 0

    dtpFechaOrdenDesde.Value = gdatFechaActual
    dtpFechaOrdenHasta.Value = dtpFechaOrdenDesde.Value
    dtpFechaLiquidacionDesde.Value = Null
    dtpFechaLiquidacionHasta.Value = dtpFechaLiquidacionDesde.Value
    
    txtPorcenIgvInt(0).Text = CStr(gdblTasaIgv * 100)
    txtPorcenIgvInt(1).Text = CStr(gdblTasaIgv * 100)
    
    txtPorcenIgv(0).Text = CStr(gdblTasaIgv * 100)
    txtPorcenIgv(1).Text = CStr(gdblTasaIgv * 100)
    
    txtNumAnexo.Text = strNumAnexo
    
    strTipoPersonaLim = Valor_Caracter
    strCodPersonaLim = Valor_Caracter
    strCodAnaliticaOrig = Valor_Caracter
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
    
    tdgConsulta.Columns(0).Width = tdgConsulta.Width * 0.01 * 8
    tdgConsulta.Columns(1).Width = tdgConsulta.Width * 0.01 * 10
    tdgConsulta.Columns(3).Width = tdgConsulta.Width * 0.01 * 20
    tdgConsulta.Columns(4).Width = tdgConsulta.Width * 0.01 * 4
    tdgConsulta.Columns(5).Width = tdgConsulta.Width * 0.01 * 5
    tdgConsulta.Columns(7).Width = tdgConsulta.Width * 0.01 * 6
    tdgConsulta.Columns(8).Width = tdgConsulta.Width * 0.01 * 9
    tdgConsulta.Columns(9).Width = tdgConsulta.Width * 0.01 * 8
    tdgConsulta.Columns(10).Width = tdgConsulta.Width * 0.01 * 7
    tdgConsulta.Columns(16).Width = tdgConsulta.Width * 0.01 * 8
    tdgConsulta.Columns(18).Width = tdgConsulta.Width * 0.01 * 8
    tdgConsulta.Columns(20).Width = tdgConsulta.Width * 0.01 * 8
    tdgConsulta.Columns(22).Width = tdgConsulta.Width * 0.01 * 15
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    
    Set adoRegistro = New ADODB.Recordset
    adoComm.CommandText = "SELECT ValorParametro AS PersonalizaComi FROM ParametroGeneral WHERE CodParametro = '32'"
    Set adoRegistro = adoComm.Execute

    If Not (adoRegistro.EOF) Then
        strPersonalizaComision = Trim$(adoRegistro("PersonalizaComi"))
    End If
    
    dblPorcDescuento = 100
        
    adoRegistro.Close: Set adoRegistro = Nothing
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmOrdenRentaFijaCortoPlazo = Nothing
    Call OcultarReportes
    frmMainMdi.stbMdi.Panels(3).Text = "Acción"
    
End Sub

Private Sub lblCantidadResumen_Change()

    Call FormatoMillarEtiqueta(lblCantidadResumen, Decimales_Monto)
    
End Sub

Private Sub lblComisionesResumen_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblComisionesResumen(Index), Decimales_Monto)
    
End Sub

Private Sub lblComisionIgv_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblComisionIgv(Index), Decimales_Monto)
    
End Sub

Private Sub lblInteresesResumen_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblInteresesResumen(Index), Decimales_Monto)
    
End Sub

Private Sub lblMontoTotal_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblMontoTotal(Index), Decimales_Monto)
    
    lblTotalResumen(Index).Caption = lblMontoTotal(Index).Caption
    
End Sub

Private Sub lblMontoVencimiento_Change()

    Call FormatoMillarEtiqueta(lblMontoVencimiento, Decimales_Monto)
    
End Sub

Private Sub lblPorcenBolsa_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblPorcenBolsa(Index), Decimales_Tasa)
    
End Sub

Private Sub lblPorcenCavali_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblPorcenCavali(Index), Decimales_Tasa)
    
End Sub

Private Sub lblPorcenConasev_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblPorcenConasev(Index), Decimales_Tasa)
    
End Sub

Private Sub lblPorcenFondo_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblPorcenFondo(Index), Decimales_Tasa)
    
End Sub

Private Sub lblPrecioResumen_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblPrecioResumen(Index), Decimales_Precio)
    
End Sub

Private Sub lblStockNominal_Change()

    Call FormatoMillarEtiqueta(lblStockNominal, Decimales_Monto)
    
End Sub

Private Sub lblSubTotal_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblSubTotal(Index), Decimales_Monto)
    
    If Not IsNumeric(txtPorcenAgente(Index).Text) Or Not IsNumeric(lblPorcenBolsa(Index).Caption) Or Not IsNumeric(lblPorcenCavali(Index).Caption) Or Not IsNumeric(lblPorcenFondo(Index).Caption) Or Not IsNumeric(lblPorcenConasev(Index).Caption) Then Exit Sub
    
    txtComisionBolsa(Index).Text = CDbl((lblSubTotal(Index).Caption)) * CDbl(lblPorcenBolsa(Index).Caption) / 100
    txtComisionCavali(Index).Text = CDbl((lblSubTotal(Index).Caption)) * CDbl(lblPorcenCavali(Index).Caption) / 100
    txtComisionFondo(Index).Text = CDbl((lblSubTotal(Index).Caption)) * CDbl(lblPorcenFondo(Index).Caption) / 100
    txtComisionConasev(Index).Text = CDbl((lblSubTotal(Index).Caption)) * CDbl(lblPorcenConasev(Index).Caption) / 100
    
    If Not IsNumeric(txtTasa.Text) Or Not IsNumeric(txtCantidad.Text) Then Exit Sub
    
    If strCalcVcto = "V" Then
    Else
    End If

    Call CalculoTotal(Index)
    
    lblSubTotalResumen(Index).Caption = CStr(CCur(lblSubTotal(Index).Caption))
    
End Sub

Private Sub lblSubTotalResumen_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblSubTotalResumen(Index), Decimales_Monto)
    
End Sub

Private Sub lblTirBruta_Change()

    Call FormatoMillarEtiqueta(lblTirBruta, Decimales_Tasa)
    
End Sub

Private Sub lblTirBrutaResumen_Change()

    Call FormatoMillarEtiqueta(lblTirBrutaResumen, Decimales_Tasa)
    
End Sub

Private Sub lblTirNeta_Change()

    Call FormatoMillarEtiqueta(lblTirNeta, Decimales_Tasa)
    
End Sub

Private Sub lblTirNetaResumen_Change()

    Call FormatoMillarEtiqueta(lblTirNetaResumen, Decimales_Tasa)
    
End Sub

Private Sub lblTotalResumen_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblTotalResumen(Index), Decimales_Monto)
    
End Sub

Private Sub lblVencimientoResumen_Change()

    Call FormatoMillarEtiqueta(lblVencimientoResumen, Decimales_Monto)
    
End Sub

Private Sub tabRFCortoPlazo_Click(PreviousTab As Integer)
            
    Dim dblMontoInicio As Double
    Dim dblMontoFin    As Double
    
    If lblMontoTotal(0).Caption <> "" Then
        dblMontoInicio = CDbl(lblMontoTotal(0).Caption)
    End If

    dblMontoFin = CDbl(txtMontoVencimiento1.Value)
    
    Select Case tabRFCortoPlazo.Tab

        Case 1, 2, 3, 4

            If PreviousTab = 0 And blnCargadoDesdeCartera = False And strEstado = Reg_Consulta Then tabRFCortoPlazo.Tab = 0
            If strEstado = Reg_Defecto Then tabRFCortoPlazo.Tab = 0
            
            If tabRFCortoPlazo.Tab = 2 Then
                If strCodTipoOrden <> Codigo_Orden_PagoCancelacion And strCodTipoOrden <> Codigo_Orden_Prepago Then
                    If ValidaRequisitosTab(2, PreviousTab) = True Then
                        fraDatosNegociacion.Caption = "Negociación" & Space$(1) & "-" & Space$(1) & Trim$(cboTipoOrden.Text) & Space$(1) & Trim$(Left$(cboTitulo.Text, 15))
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

Private Sub CalcularComision()

    Dim adoRegistro           As ADODB.Recordset
    Dim TCComision            As Double
    Dim dblTotalMEAnexo       As Double
    Dim dblTotalMEAnexoDesc   As Double
    Dim dblsubTotal1          As Double
    Dim dblsubTotal2          As Double
    Dim dblMontoMinComisiones As Double
    Dim dblPorcentajeComision As Double

    Set adoRegistro = New ADODB.Recordset
    adoComm.CommandText = "SELECT ValorParametro AS CodMonedaComisiones FROM ParametroGeneral WHERE CodParametro = '29'"
    Set adoRegistro = adoComm.Execute

    If Not (adoRegistro.EOF) Then
        strCodMonedaComision = Trim$(adoRegistro("CodMonedaComisiones"))
    End If

    If Trim$(strCodMonedaComision) = "" Then
        MsgBox "Moneda de cobro de las comisiones no está definida en Parámetros Globales.", vbCritical
        Exit Sub
    End If
     
    'txtPorcenAgente(0).Text = 0#
     
    Set adoRegistro = New ADODB.Recordset
    adoComm.CommandText = "SELECT ValorParametro AS PorcentajeComision FROM ParametroGeneral WHERE CodParametro = '37'"
    Set adoRegistro = adoComm.Execute

    If Not (adoRegistro.EOF) Then
        If Trim$(adoRegistro("PorcentajeComision")) <> "" Then
            dblPorcentajeComision = CDbl(adoRegistro("PorcentajeComision"))
            txtPorcenAgente(0).Text = dblPorcentajeComision
        Else
            MsgBox "Porcentaje de comisión no está definido en Parámetros Globales.", vbCritical
            Exit Sub
        End If
    Else
        dblPorcentajeComision = CDec(txtPorcenAgente(0).Text)
    End If
     
    Set adoRegistro = New ADODB.Recordset
    adoComm.CommandText = "SELECT ValorParametro AS MontoMinComisiones FROM ParametroGeneral WHERE CodParametro = '31'"
    Set adoRegistro = adoComm.Execute

    If Not (adoRegistro.EOF) Then
        If Trim$(adoRegistro("MontoMinComisiones")) <> "" Then
            dblMontoMinComisiones = CDbl(adoRegistro("MontoMinComisiones"))
        Else
            MsgBox "Monto mínimo de comisión no está definido en Parámetros Globales.", vbCritical
            Exit Sub
        End If
    End If
     
    adoRegistro.Close: Set adoRegistro = Nothing

    strCodMonedaParEvaluacion = Trim$(strCodMoneda) & Trim$(strCodMonedaComision)

    If Trim$(strCodMoneda) <> strCodMonedaComision Then
        strCodMonedaParPorDefecto = ObtenerMonedaParPorDefecto(gstrCodClaseTipoCambioOperacionFondo, strCodMonedaParEvaluacion)
    Else
        strCodMonedaParPorDefecto = strCodMonedaParEvaluacion
    End If
            
    If strCodMonedaParPorDefecto = "0000" Then strCodMonedaParPorDefecto = strCodMonedaParEvaluacion
  
    If strCodMonedaComision <> Trim$(strCodMoneda) Then
        
        TCComision = 0#
        TCComision = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, dtpFechaOrden.Value, Mid$(strCodMonedaParPorDefecto, 1, 2), Mid$(strCodMonedaParPorDefecto, 3, 2)))

        If TCComision <> 0 Then
            If strCodMonedaParEvaluacion = strCodMonedaParPorDefecto Then
                dblsubTotal1 = txtValorNominal.Value / TCComision
            Else
                dblsubTotal1 = txtValorNominal.Value * TCComision
            End If

        Else
            dblsubTotal1 = 0#
        End If

    Else
        dblsubTotal1 = CDec(txtValorNominal.Text)
        
    End If
    
    dblTotalMEAnexo = (dblsubTotal1 + dblsubTotal2)
    dblTotalMEAnexoDesc = dblTotalMEAnexo * (CDbl(txtPorcenDctoValorNominal.Text) / 100)
    dblComisionOperacion = dblTotalMEAnexoDesc * (dblPorcentajeComision / 100)
    
    If dblComisionOperacion < dblMontoMinComisiones Then
        dblComisionOperacion = dblMontoMinComisiones
    End If
    
    Call AsignarComisionOperacion
    
End Sub

Private Sub AsignarComisionOperacion()

    Dim dblTipoCambio2 As Double

    'If dblComisionOperacion <> 0 Then
        If Trim$(strCodMoneda) <> Trim$(strCodMonedaComision) Then
            
            strCodMonedaParEvaluacion = Trim$(strCodMonedaComision) & Trim$(strCodMoneda)
            strCodMonedaParPorDefecto = ObtenerMonedaParPorDefecto(gstrCodClaseTipoCambioOperacionFondo, strCodMonedaParEvaluacion)

            If strCodMonedaParPorDefecto = "0000" Then strCodMonedaParPorDefecto = strCodMonedaParEvaluacion
      
            dblTipoCambio2 = 0#
            dblTipoCambio2 = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, dtpFechaOrden.Value, Mid$(strCodMonedaParPorDefecto, 1, 2), Mid$(strCodMonedaParPorDefecto, 3, 2)))
            
            If dblTipoCambio2 <> 0 Then
                If strCodMonedaParEvaluacion = strCodMonedaParPorDefecto Then
                    txtComisionAgente(0).Text = dblComisionOperacion / dblTipoCambio2
                Else
                    txtComisionAgente(0).Text = dblComisionOperacion * dblTipoCambio2
                End If

            Else
                txtComisionAgente(0).Text = 0#
            End If
        
        Else
            txtComisionAgente(0).Text = dblComisionOperacion
        End If
        
        If CDec(txtComisionAgente(0).Text) <> 0 Then
            Call CalculoTotal(0)
        End If
    
   ' End If

End Sub

Private Function ValidaRequisitosTab(intIndTab As Integer, intTabOrigen) As Boolean

    ValidaRequisitosTab = False

    Select Case intIndTab

        Case 2

            If CInt(txtDiasPlazo.Text) <= 0 Or cboMoneda.ListIndex <= 0 Then
                MsgBox "Verifique si la moneda y el plazo están ingresados.", vbCritical, Me.Caption
                Exit Function
            End If
     
            If cboEmisor.ListIndex <= 0 Then
                MsgBox "Debe seleccionar el Emisor.", vbCritical, Me.Caption

                If cboEmisor.Enabled And tabRFCortoPlazo.TabEnabled(1) = True Then
                    tabRFCortoPlazo.Tab = 1
                    cboEmisor.SetFocus
                End If

                Exit Function
            End If
    
'            If cboObligado.ListIndex <= 0 Then
'                MsgBox "Debe seleccionar el Obligado.", vbCritical, Me.Caption
'
'                If cboObligado.Enabled And tabRFCortoPlazo.TabEnabled(1) = True Then
'                    tabRFCortoPlazo.Tab = 1
'                   ' cboObligado.SetFocus
'                End If
'
'                Exit Function
'            End If
'
            If cboGestor.ListIndex <= 0 Then
                MsgBox "Debe seleccionar el Gestor.", vbCritical, Me.Caption

                If cboGestor.Enabled And tabRFCortoPlazo.TabEnabled(1) = True Then
                    tabRFCortoPlazo.Tab = 1
                    cboGestor.SetFocus
                End If

                Exit Function
            End If
    
            If cboLineaCliente.ListIndex <= 0 And cboLineaCliente.ListCount > 1 Then
                MsgBox "Debe especificar la línea a afectar.", vbCritical, Me.Caption

                If cboLineaCliente.Enabled And tabRFCortoPlazo.TabEnabled(1) = True Then
                    tabRFCortoPlazo.Tab = 1
                    cboLineaCliente.SetFocus
                End If

                Exit Function
            End If
    
            If cboResponsablePago.ListIndex <= 0 Then
                MsgBox "Debe especificar quién hará el pago al vencimiento.", vbCritical, Me.Caption

                If cboResponsablePago.Enabled And tabRFCortoPlazo.TabEnabled(1) = True Then
                    tabRFCortoPlazo.Tab = 1
                    cboResponsablePago.SetFocus
                End If

                Exit Function
            End If
    
        Case 3

    End Select

    ValidaRequisitosTab = True

End Function

Private Sub tdgConsulta_FormatText(ByVal ColIndex As Integer, _
                                   Value As Variant, _
                                   Bookmark As Variant)

    If ColIndex = 4 Then
        Call DarFormatoValor(Value, Decimales_Monto)
    End If
    
    If ColIndex = 6 Then
        Call DarFormatoValor(Value, Decimales_Precio)
    End If
    
    If ColIndex = 7 Then
        Call DarFormatoValor(Value, Decimales_Monto)
    End If
    
End Sub

Private Sub txtCantidad_Change()

    Call FormatoCajaTexto(txtCantidad, Decimales_Monto)
    
    Call txtPrecioUnitario1_Change '(0)
    
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtCantidad, Decimales_Monto)
    
End Sub

Private Sub txtComisionAgente_Change(Index As Integer)

    Call FormatoCajaTexto(txtComisionAgente(Index), Decimales_Monto)
    
    If strPersonalizaComision = "NO" Then
        If chkAplicar(Index).Value = Checked Then
            lblComisionIgv(0).Caption = (CDbl(txtPorcenIgv(0).Text) / 100) * CDbl(txtComisionAgente(0).Text)
        End If

    Else

        If chkAplicar(Index).Value = Checked Then
            ActualizaPorcentaje txtComisionAgente(Index), txtPorcenAgente(Index)
            lblComisionIgv(0).Caption = (CDbl(txtPorcenIgv(0).Text) / 100) * CDbl(txtComisionAgente(0).Text)
        End If
    
    End If

End Sub

Private Sub txtComisionAgente_KeyPress(Index As Integer, KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtComisionAgente(Index), Decimales_Monto)

    If KeyAscii = vbKeyReturn Then
        If strPersonalizaComision = "NO" Then
            lblComisionIgv(0).Caption = (CDbl(txtPorcenIgv(0).Text) / 100) * CDbl(txtComisionAgente(0).Text)
            Call CalculoTotal(Index)
        Else

            If chkAplicar(Index).Value Then
                ActualizaPorcentaje txtComisionAgente(Index), txtPorcenAgente(Index)
                lblComisionIgv(0).Caption = (CDbl(txtPorcenIgv(0).Text) / 100) * CDbl(txtComisionAgente(0).Text)
            End If

            Call CalculoTotal(Index)
        End If
    End If
    
End Sub

Private Sub txtComisionBolsa_Change(Index As Integer)

    Call FormatoCajaTexto(txtComisionBolsa(Index), Decimales_Monto)

    If chkAplicar(Index).Value Then
        ActualizaPorcentaje txtComisionBolsa(Index), lblPorcenBolsa(Index)
    End If
    
End Sub

Private Sub txtComisionBolsa_KeyPress(Index As Integer, KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtComisionBolsa(Index), Decimales_Monto)
    
    If KeyAscii = vbKeyReturn Then
        If chkAplicar(Index).Value Then
            ActualizaPorcentaje txtComisionBolsa(Index), lblPorcenBolsa(Index)
        End If

        Call CalculoTotal(Index)
    End If
    
End Sub

Private Sub txtComisionCavali_Change(Index As Integer)

    Call FormatoCajaTexto(txtComisionCavali(Index), Decimales_Monto)
    
    If chkAplicar(Index).Value Then
        ActualizaPorcentaje txtComisionCavali(Index), lblPorcenCavali(Index)
    End If
    
End Sub

Private Sub txtComisionCavali_KeyPress(Index As Integer, KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtComisionCavali(Index), Decimales_Monto)
    
    If KeyAscii = vbKeyReturn Then
        If chkAplicar(Index).Value Then
            ActualizaPorcentaje txtComisionCavali(Index), lblPorcenCavali(Index)
        End If

        Call CalculoTotal(Index)
    End If
    
End Sub

Private Sub txtComisionConasev_Change(Index As Integer)

    Call FormatoCajaTexto(txtComisionConasev(Index), Decimales_Monto)
    
    If chkAplicar(Index).Value Then
        ActualizaPorcentaje txtComisionConasev(Index), lblPorcenConasev(Index)
    End If
    
End Sub

Private Sub txtComisionConasev_KeyPress(Index As Integer, KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtComisionConasev(Index), Decimales_Monto)
    
    If KeyAscii = vbKeyReturn Then
        If chkAplicar(Index).Value Then
            ActualizaPorcentaje txtComisionConasev(Index), lblPorcenConasev(Index)
        End If

        Call CalculoTotal(Index)
    End If
        
End Sub

Private Sub txtComisionFondo_Change(Index As Integer)

    Call FormatoCajaTexto(txtComisionFondo(Index), Decimales_Monto)
    
    If chkAplicar(Index).Value Then
        ActualizaPorcentaje txtComisionFondo(Index), lblPorcenFondo(Index)
    End If
    
End Sub

Private Sub txtComisionFondo_KeyPress(Index As Integer, KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtComisionFondo(Index), Decimales_Monto)
    
    If KeyAscii = vbKeyReturn Then
        If chkAplicar(Index).Value Then
            ActualizaPorcentaje txtComisionFondo(Index), lblPorcenFondo(Index)
        End If

        Call CalculoTotal(Index)
    End If
    
End Sub

Private Sub CalculoTotal(Index As Integer)

    Dim curComImp As Currency, curMonTotal As Currency, curInteresCorrido As Currency
    Dim curIntImp As Currency, curImptoInteresCorrido As Currency

    If Not IsNumeric(txtComisionAgente(Index).Text) Or Not IsNumeric(txtComisionBolsa(Index).Text) Or Not IsNumeric(txtComisionConasev(Index).Text) Or Not IsNumeric(txtComisionCavali(Index).Text) Or Not IsNumeric(txtComisionFondo(Index).Text) Or Not IsNumeric(txtInteresCorrido(Index).Text) Then Exit Sub

    txtImptoInteres(Index).Text = Round(CCur(CDbl(lblIntAdelantado(Index).Caption) * (CDbl(txtPorcenIgvInt(Index).Value)) / 100), 2)
    txtImptoInteresAdic(Index).Text = Round(CCur(CDbl(txtIntAdicional(Index).Text) * (CDbl(txtPorcenIgvInt(Index).Value)) / 100), 2)
    curIntImp = CCur(txtImptoInteres(Index).Text) + CCur(txtImptoInteresAdic(Index).Text)
        
    If strCodCobroInteres = Codigo_Modalidad_Pago_Vencimiento Then  'No afecta al cálculo de desembolso
        lblIntAdelantado(Index).ForeColor = &H80000011
        txtIntAdicional(Index).ForeColor = &H80000011
        lblComisionIgvInt(Index).ForeColor = &H80000011
    Else
        lblIntAdelantado(Index).ForeColor = &H80000012
        txtIntAdicional(Index).ForeColor = &H80000012
        lblComisionIgvInt(Index).ForeColor = &H80000012
    End If
    
    curComImp = CCur(CCur(txtComisionAgente(Index).Text) + CCur(txtComisionBolsa(Index).Text) + CCur(txtComisionConasev(Index).Text) + CCur(txtComisionCavali(Index).Text) + CCur(txtComisionFondo(Index).Text)) * (CDbl(txtPorcenIgv(Index).Value) / 100)
   
    lblComisionIgv(Index).Caption = CStr(curComImp)
    lblComisionIgvInt(Index).Caption = CStr(curIntImp)
     
    If strCodCobroInteres = Codigo_Modalidad_Pago_Adelantado Then
        curComImp = CCur(txtComisionAgente(Index).Text) + CCur(txtComisionBolsa(Index).Text) + CCur(txtComisionConasev(Index).Text) + CCur(txtComisionCavali(Index).Text) + CCur(txtComisionFondo(Index).Text) + CCur(lblIntAdelantado(Index).Caption) + CCur(txtIntAdicional(Index).Text) + CCur(lblComisionIgv(Index).Caption) + CCur(lblComisionIgvInt(Index).Caption)
    Else
        curComImp = CCur(txtComisionAgente(Index).Text) + CCur(txtComisionBolsa(Index).Text) + CCur(txtComisionConasev(Index).Text) + CCur(txtComisionCavali(Index).Text) + CCur(txtComisionFondo(Index).Text) + CCur(lblComisionIgv(Index).Caption)
    End If

    lblComisionesResumen(Index).Caption = CStr(curComImp)
    
    curInteresCorrido = CCur(txtInteresCorrido(Index).Text)
    curImptoInteresCorrido = CCur(txtImptoInteresCorrido(Index).Text)
    
    If strCodTipoOrden = Codigo_Orden_Compra Or strCodTipoOrden = Codigo_Orden_Renovacion Or strCodTipoOrden = Codigo_Orden_Pacto Then  '*** Compra ***
        If Index = 0 Then
            curMonTotal = CCur(lblSubTotal(Index).Caption) - curComImp
        Else
            curMonTotal = CCur(lblSubTotal(Index).Caption) + curComImp
        End If
        

    ElseIf strCodTipoOrden = Codigo_Orden_Venta Then '*** Venta ***
        curMonTotal = CCur(lblSubTotal(Index).Caption) - curComImp
    End If
    
    curMonTotal = curMonTotal
    lblMontoTotal(Index).Caption = CStr(curMonTotal)
    
    If strCalcVcto = "D" Then
        If Trim$(txtValorNominalDcto.Text) <> "" Then
            txtMontoVencimiento1.Text = CDbl(txtValorNominalDcto.Text) * CCur(txtCantidad.Text)
        End If

    Else

        If Trim$(txtValorNominalDcto.Text) <> "" Then
            txtMontoVencimiento1.Text = CDbl(txtValorNominalDcto.Text) * CCur(txtCantidad.Text)
        End If
    End If
    
    If strCodCobroInteres = Codigo_Modalidad_Pago_Vencimiento Then
        txtMontoVencimiento1.Text = (CDbl(txtMontoVencimiento1.Value) + CDbl(lblIntAdelantado(Index).Caption) + CDbl(txtIntAdicional(Index).Text) + curIntImp)
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
        
        lblDiasPlazo.Caption = CStr(txtDiasPlazo.Text)
        lblFechaVencimiento.Caption = CStr(dtpFechaVencimiento.Value)
        lblFechaCupon.Caption = CStr(dtpFechaVencimiento.Value)
    
    End If
    
End Sub

Private Sub txtDiasPlazo_LostFocus()

    txtDiasPlazo_KeyPress (vbKeyReturn)
    cboEmisor_Click
    
End Sub

Private Sub AsignaComision(strTipoComision As String, _
                           dblValorComision As Double, _
                           ctrlValorComision As Control)
    
    If Not IsNumeric(lblSubTotal(ctrlValorComision.Index).Caption) Then Exit Sub
    
    If dblValorComision > 0 Then
        ctrlValorComision.Text = CStr(CCur(lblSubTotal(ctrlValorComision.Index)) * dblValorComision / 100)
    End If
            
End Sub

Private Sub ActualizaPorcentaje(ctrlComision As Control, ctrlPorcentaje As Control)

    If Not IsNumeric(ctrlComision) Or Not IsNumeric(lblSubTotal(ctrlComision.Index).Caption) Then Exit Sub
                
    If CCur(lblSubTotal(ctrlComision.Index)) = 0 Then
        ctrlPorcentaje = "0"
    'Else

'        If CCur(ctrlComision) > 0 Then
'            ctrlPorcentaje = CStr((CCur(ctrlComision) / CCur(lblSubTotal(ctrlComision.Index).Caption)) * 100)
'        Else
'            ctrlPorcentaje = "0"
'        End If
    End If
                
End Sub

Private Sub txtIntAdicional_KeyPress(Index As Integer, KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtIntAdicional(Index), Decimales_Monto)
    
    If KeyAscii = vbKeyReturn Then
        Call CalculoTotal(Index)
    End If

End Sub

Private Sub txtInteresCorrido_Change(Index As Integer)

    Call FormatoCajaTexto(txtInteresCorrido(Index), Decimales_Monto)
    
    If Trim$(txtInteresCorrido(Index).Text) <> Valor_Caracter Then
        lblInteresesResumen(Index).Caption = CStr(CCur(txtInteresCorrido(Index).Text))
        txtImptoInteresCorrido(Index).Text = (txtInteresCorrido(Index).Text * txtPorcenIgvInt(Index).Text / 100)
    End If
    
End Sub

Private Sub txtInteresCorrido_KeyPress(Index As Integer, KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtInteresCorrido(Index), Decimales_Monto)
    
    If KeyAscii = vbKeyReturn Then
        txtImptoInteresCorrido(Index).Text = (txtInteresCorrido(Index).Text * txtPorcenIgvInt(Index).Text / 100)
        Call CalculoTotal(Index)
    End If
    
End Sub

Private Sub ActualizaComision(ctrlPorcentaje As Control, ctrlComision As Control)

    If Not IsNumeric(lblSubTotal(ctrlComision.Index).Caption) Or Not IsNumeric(ctrlPorcentaje) Then Exit Sub
        
    If CDbl(ctrlPorcentaje) > 0 Then
        ctrlComision = CStr(CCur(lblSubTotal(ctrlComision.Index).Caption) * CDbl(ctrlPorcentaje) / 100)
    Else
        ctrlComision = "0"
    End If
        
End Sub

Private Sub txtMontoTotalRecibido_Change()

    If txtMontoTotalRecibido.Text = "" Then txtMontoTotalRecibido.Text = 0
    If CDbl(txtMontoTotalRecibido.Text) >= CDbl(txtDeudaTotal.Text) Then
        txtMontoTotalCancel.Text = txtDeudaTotal.Text
        txtSaldoDeuda.Text = 0
    Else
        txtMontoTotalCancel.Text = txtMontoTotalRecibido.Text
        txtSaldoDeuda.Text = CDbl(txtDeudaTotal.Text) - CDbl(txtMontoTotalCancel.Text)
    End If
    
End Sub

Private Sub txtMontoTotalRecibido_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        
        If CDbl(txtMontoTotalRecibido.Text) >= CDbl(txtDeudaTotal.Text) Then
            txtMontoTotalCancel.Text = txtDeudaTotal.Text
            txtSaldoDeuda.Text = 0
        Else
            txtMontoTotalCancel.Text = txtMontoTotalRecibido.Text
            txtSaldoDeuda.Text = CDbl(txtDeudaTotal.Text) - CDbl(txtMontoTotalCancel.Text)
        End If

    End If

End Sub

Private Sub txtNemonico_Change()

    txtDescripOrden = Trim$(cboTipoInstrumentoOrden.Text) & " - " & Trim$(txtNemonico.Text)
    
End Sub

Private Sub txtNemonico_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    
End Sub

Private Sub txtNemonicoCancel_Change()

    txtDescripOrdenCancel = Trim$(cboTipoInstrumentoOrden.Text) & " - " & Trim$(txtNemonicoCancel.Text)

End Sub

Private Sub txtNumAnexo_Change()

    txtNemonico.Text = GenerarNemonico(strCodTipoInstrumentoOrden, strCodClaseInstrumento, strCodEmisor, txtNumAnexo.Text)
    txtDescripOrden.Text = Trim$(cboTipoInstrumentoOrden.Text) & " - " & txtNemonico.Text

End Sub

Private Sub txtNumDocDsctoCancel_Change()

    txtNemonicoCancel.Text = GenerarNemonico(strCodTipoInstrumentoOrden, strCodClaseInstrumento, strCodEmisor, txtNumDocDsctoCancel.Text)
    txtDescripOrdenCancel.Text = Trim$(cboTipoInstrumentoOrden.Text) & " - " & txtNemonicoCancel.Text

End Sub

Private Sub txtNumDocDsctoCancel_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        txtNemonicoCancel.Text = GenerarNemonico(strCodTipoInstrumentoOrden, strCodClaseInstrumento, strCodEmisor, txtNumDocDsctoCancel.Text)
        txtDescripOrdenCancel.Text = Trim$(cboTipoInstrumentoOrden.Text) & " - " & txtNemonicoCancel.Text
    End If
    
End Sub

Private Sub txtNumOperacionOrig_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        txtNumOperacionOrig.Text = Format$(txtNumOperacionOrig.Text, "0000000000")
        Call CargarSolicitud(strCodFondoDescuento, gstrCodAdministradora, txtNumOperacionOrig.Text, 0)
    End If

End Sub

Private Sub txtPorcenAgente_Change(Index As Integer)
     If strPersonalizaComision <> "NO" Then
            'Call ValidaCajaTexto(KeyAscii, "M", txtPorcenAgente(Index), Decimales_Tasa)

            If chkAplicar(Index).Value Then
                ActualizaComision txtPorcenAgente(Index), txtComisionAgente(Index)
            End If
                        
            Call CalcularComision
            Call AsignarComisionOperacion

            Call CalculoTotal(Index)
        End If
'    If strPersonalizaComision <> "NO" Then
'        Call FormatoCajaTexto(txtPorcenAgente(index), Decimales_Tasa)
'    End If

End Sub

Private Sub txtPorcenAgente_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then

        If strPersonalizaComision <> "NO" Then
            'Call ValidaCajaTexto(KeyAscii, "M", txtPorcenAgente(Index), Decimales_Tasa)

            If chkAplicar(Index).Value Then
                ActualizaComision txtPorcenAgente(Index), txtComisionAgente(Index)
            End If
                        
            Call CalcularComision
            Call AsignarComisionOperacion

            Call CalculoTotal(Index)
        End If

    End If
        
End Sub

Private Sub txtPorcenDctoValorNominal_Change()

    Call txtValorNominal_Change

End Sub

Private Sub txtPorcenIgv_Change(Index As Integer)

    If chkAplicar(Index).Value Then
        ActualizaComision txtPorcenIgv(Index), lblComisionIgv(Index)

        If Trim$(txtPorcenIgv(0).Text) <> Null Then
            lblComisionIgv(0).Caption = (CDbl(txtPorcenIgv(0).Text) / 100) * CDbl(txtComisionAgente(0).Text)
        End If
    End If

    Call CalculoTotal(Index)
    
End Sub

Private Sub txtPorcenIgv_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If chkAplicar(Index).Value Then
            ActualizaComision txtPorcenIgv(Index), lblComisionIgv(Index)

            If Trim$(txtPorcenIgv(0).Text) <> Null Then
                lblComisionIgv(0).Caption = (CDbl(txtPorcenIgv(0).Text) / 100) * CDbl(txtComisionAgente(0).Text)
            End If
        End If

        Call CalculoTotal(Index)
    End If

End Sub

Private Sub txtPorcenIgvInt_Change(Index As Integer)

    If chkAplicar(Index).Value Then
        ActualizaComision txtPorcenIgvInt(Index), lblComisionIgvInt(Index)
    End If

    Call CalculoTotal(Index)

End Sub

Private Sub txtPrecioUnitario_Change(Index As Integer)

    Call FormatoCajaTexto(txtPrecioUnitario(Index), Decimales_Precio)

    If Not IsNumeric(txtCantidad.Text) Or Not IsNumeric(txtValorNominal.Text) Or Not IsNumeric(txtPrecioUnitario(Index).Text) Or Not IsNumeric(txtDiasPlazo.Text) Then Exit Sub
    If Not (CCur(txtCantidad.Text) > 0 And CDbl(txtValorNominal.Text) > 0 And CDbl(txtPrecioUnitario(Index).Text) > 0 And CInt(txtDiasPlazo.Text) > 0) Then Exit Sub

    lblSubTotal(Index).Caption = CDbl(txtValorNominal.Text) * CCur(txtCantidad.Text) * CDbl(txtPrecioUnitario(Index).Text) / 100

    If txtPrecioUnitario(Index).Tag = "0" Then
        txtTirBruta1.Tag = "1"
        txtTirBruta1.Text = ((CDbl(txtMontoVencimiento1.Value) / (CDbl(txtPrecioUnitario(0).Text) / 100 * CCur(txtCantidad.Text) * CDbl(txtValorNominal.Text))) ^ (360 / CInt(txtDiasPlazo.Text)) - 1) * 100
    Else
        txtPrecioUnitario(Index).Tag = "0"
    End If

    lblPrecioResumen(Index).Caption = CStr(txtPrecioUnitario(Index).Text)
    
End Sub

Private Sub txtPrecioUnitario_KeyPress(Index As Integer, KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtPrecioUnitario(Index), Decimales_Precio)
    
End Sub

Private Sub txtPrecioUnitario1_Change()
        
    If Not IsNumeric(txtCantidad.Text) Or Not IsNumeric(txtDiasPlazo.Text) Or Not IsNumeric(txtValorNominalDcto.Text) Then Exit Sub
    
    If Not (CCur(txtCantidad.Text) > 0 And CDbl(txtValorNominalDcto.Text) > 0 And txtPrecioUnitario1.Value > 0 And CInt(txtDiasPlazo.Text) > 0) Then Exit Sub
        
    If chkDiasAdicional.Visible = True And chkDiasAdicional.Value = Checked Then
        datFechaVctoAdicional = DateAdd("d", intDiasAdicionales, CVDate(dtpFechaVencimiento.Value))

        If Not EsDiaUtil(datFechaVctoAdicional) Then
            datFechaVctoAdicional = ProximoDiaUtil(datFechaVctoAdicional)
        End If

        lblFechaVencimientoAdic.Caption = datFechaVctoAdicional
        lblFechaVencimientoAdic.Visible = True
        lblDiasAdic(0).Caption = "( " & CStr(DateDiff("d", dtpFechaVencimiento.Value, datFechaVctoAdicional)) & " días )"
    Else
        datFechaVctoAdicional = dtpFechaVencimiento.Value
        lblFechaVencimientoAdic.Caption = datFechaVctoAdicional
        lblFechaVencimientoAdic.Visible = False
        lblDiasAdic(0).Caption = ""
    End If
    
    If strCalcVcto = "V" Then
        
        lblSubTotal(0).Caption = CDbl(txtValorNominalDcto.Text)

        If strCodCobroInteres = Codigo_Modalidad_Pago_Adelantado And (chkDiasAdicional.Visible = True And chkDiasAdicional.Value = Checked) Then
            txtIntAdicional(0).Text = CalculoInteresDescuento(CDbl(txtTasa.Text), strCodTipoTasa, "01", strCodBaseAnual, txtValorNominalDcto, CStr(dtpFechaOrden.Value), CStr(datFechaVctoAdicional)) - CDbl(lblIntAdelantado(0).Caption)
        Else
            txtIntAdicional(0).Text = 0#
        End If

    Else
        lblIntAdelantado(0).Caption = 0#
        txtIntAdicional(0).Text = 0#
        lblSubTotal(0).Caption = CDbl(txtValorNominalDcto.Text) * CCur(txtCantidad.Text) * txtPrecioUnitario1.Value / 100
    End If

    If txtPrecioUnitario1.Tag = "0" Then
        txtTirBruta1.Tag = "1"
        txtTirBruta1.Text = ((1 - (txtPrecioUnitario1.Value / 100) + 1) ^ (intBaseCalculo / CInt(txtDiasPlazo.Text)) - 1) * 100

    Else
        txtPrecioUnitario1.Tag = "0"
    End If

    lblPrecioResumen(0).Caption = CStr(txtPrecioUnitario1.Value)
    
    If strCalcVcto = "D" Then
        txtMontoVencimiento1.Text = txtValorNominalDcto.Text * CCur(txtCantidad.Text)
    Else
        txtMontoVencimiento1.Text = txtValorNominalDcto.Text * CCur(txtCantidad.Text)
    End If
    
    If chkInteresCorrido(0).Value = Checked Then
        txtInteresCorrido(0).Text = CalculoInteresDescuento(CDbl(txtTasa.Text), strCodTipoTasa, "01", strCodBaseAnual, txtValorNominalDcto, CStr(dtpFechaLiquidacion.Value), CStr(dtpFechaOrden.Value))
        txtImptoInteresCorrido(0).Text = (txtInteresCorrido(0).Text * txtPorcenIgvInt(0).Text / 100)
    End If
    
    If strCodCobroInteres = Codigo_Modalidad_Pago_Vencimiento Then
        txtMontoVencimiento1.Text = (CDbl(txtMontoVencimiento1.Value) + CDbl(lblIntAdelantado(0).Caption) + CDbl(txtIntAdicional(0).Text) + ((CDbl(lblIntAdelantado(0).Caption) + CDbl(txtIntAdicional(0).Text)) * (CDbl(txtPorcenIgvInt(0).Value) / 100)))
    End If

End Sub

Private Sub txtTasa_Change()

    Call FormatoCajaTexto(txtTasa, Decimales_Tasa)
    
    Call txtPrecioUnitario1_Change '(0)
    
    Call CalculoTotal(0)
    
End Sub

Private Sub txtTasa_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtTasa, Decimales_Tasa)
    
End Sub

Private Sub txtTipoCambio_Change()

    Call FormatoCajaTexto(txtTipoCambio, Decimales_TipoCambio)
    
End Sub

Private Sub txtTipoCambio_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtTipoCambio, Decimales_TipoCambio)
    
End Sub

Private Sub txtTirBruta1_Change()
    
    If Not (txtTirBruta1.Value <> 0 And CInt(txtDiasPlazo.Text) > 0 And CCur(txtCantidad.Text) > 0 And CDbl(txtValorNominalDcto.Text) > 0) Then Exit Sub
    
    If txtTirBruta1.Tag = "0" Then 'indica cambio directo en la pantalla
        txtPrecioUnitario1.Tag = "1"
        
        txtPrecioUnitario1.Text = (CDbl(txtValorNominalDcto.Text) * (1 - ((1 + 0.01 * txtTirBruta1.Value) ^ (CInt(txtDiasPlazo.Text) / intBaseCalculo) - 1))) / (CDbl(txtValorNominalDcto.Text) * CDbl(txtCantidad.Text)) * 100

    Else
        txtTirBruta1.Tag = "0"
    End If

End Sub

Private Sub txtTirNeta_Change()

    Call FormatoCajaTexto(txtTirNeta, Decimales_Tasa)

End Sub

Private Sub txtTirNeta_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtTirNeta, Decimales_Tasa)

End Sub

Private Sub txtValorNominal_Change()
  
    If Not IsNumeric(txtCantidad.Text) Then Exit Sub
    
    txtValorNominalDcto.Text = CStr(txtPorcenDctoValorNominal.Value / 100 * txtValorNominal.Value)
        
    If chkAplicar(0).Value Then Call CalcularComision
    Call txtPrecioUnitario1_Change '(0)
    Call CalculoTotal(0)
   
End Sub

Private Sub txtValorNominalDcto_Change()

    Call FormatoCajaTexto(txtValorNominalDcto, Decimales_Monto)

End Sub

Private Function FP_ReqFormasPagoOK() As Boolean

    FP_ReqFormasPagoOK = False
 
    If cboTipoInstrumentoOrden.ListIndex <= 0 Then
        MsgBox "Debe seleccionar el Tipo de Instrumento de Corto Plazo.", vbCritical, Me.Caption

        If cboTipoInstrumentoOrden.Enabled And tabRFCortoPlazo.TabEnabled(1) = True Then
            tabRFCortoPlazo.Tab = 1
            cboTipoInstrumentoOrden.SetFocus
        End If

        Exit Function
    End If
    
    If cboClaseInstrumento.ListIndex <= 0 Then
        MsgBox "Debe seleccionar la Clase de Instrumento de Corto Plazo.", vbCritical, Me.Caption

        If cboClaseInstrumento.Enabled And tabRFCortoPlazo.TabEnabled(1) = True Then
            tabRFCortoPlazo.Tab = 1
            cboClaseInstrumento.SetFocus
        End If

        Exit Function
    End If
        
    If cboTipoOrden.ListIndex < 0 Then
        MsgBox "Debe seleccionar el tipo de orden.", vbCritical, Me.Caption

        If cboTipoOrden.Enabled And tabRFCortoPlazo.TabEnabled(1) = True Then
            tabRFCortoPlazo.Tab = 1
            cboTipoOrden.SetFocus
        End If

        Exit Function
    End If

    If CInt(txtDiasPlazo.Text) = 0 And blnCancelaPrepago = False Then
        MsgBox "Debe indicar el número de días de plazo.", vbCritical, Me.Caption

        If txtDiasPlazo.Enabled And tabRFCortoPlazo.TabEnabled(1) = True Then
            tabRFCortoPlazo.Tab = 1
            txtDiasPlazo.SetFocus
        End If

        Exit Function
    End If
    
    If cboMoneda.ListIndex <= 0 And blnCancelaPrepago = False Then
        MsgBox "Debe seleccionar la Moneda.", vbCritical, Me.Caption

        If cboMoneda.Enabled And tabRFCortoPlazo.TabEnabled(1) = True Then
            tabRFCortoPlazo.Tab = 1
            cboMoneda.SetFocus
        End If

        Exit Function
    End If
    
    If CDbl(txtTasa.Text) = 0 And strCodSubClaseInstrumento = "001" And blnCancelaPrepago = False Then
        MsgBox "Debe indicar la Tasa Facial.", vbCritical, Me.Caption

        If txtTasa.Enabled And tabRFCortoPlazo.TabEnabled(1) = True Then
            tabRFCortoPlazo.Tab = 1
            txtTasa.SetFocus
        End If

        Exit Function
    End If
        
    If CCur(txtCantidad.Text) = 0 And blnCancelaPrepago = False Then
        MsgBox "Debe indicar el monto.", vbCritical, Me.Caption

        If txtCantidad.Enabled And tabRFCortoPlazo.TabEnabled(2) = True Then
            tabRFCortoPlazo.Tab = 2
            txtCantidad.SetFocus
        End If

        Exit Function
    End If
    
    If CCur(txtMontoVencimiento1.Value) = 0 And blnCancelaPrepago = False Then
        MsgBox "Debe calcular el Valor al Vencimiento.", vbCritical, Me.Caption

        If cmdCalculo.Enabled And tabRFCortoPlazo.TabEnabled(2) = True Then
            tabRFCortoPlazo.Tab = 2
            cmdCalculo.SetFocus
        End If

        Exit Function
    End If
    
    If CCur(txtPrecioUnitario1.Value) = 0 And blnCancelaPrepago = False Then
        MsgBox "Debe ingresar el %Precio. ", vbCritical, Me.Caption

        If txtPrecioUnitario1.Enabled And tabRFCortoPlazo.TabEnabled(2) = True Then
            tabRFCortoPlazo.Tab = 2
            txtPrecioUnitario1.SetFocus
        End If

        Exit Function
    End If
        
    FP_ReqFormasPagoOK = True
  
    tabRFCortoPlazo.TabEnabled(3) = True
    tabRFCortoPlazo.Tab = 3

End Function

Public Function CalculoInteresDescuento(numPorcenTasa As Double, _
                                        strCodTipoTasa As String, _
                                        strCodPeriodoTasa As String, _
                                        strCodBaseCalculo As String, _
                                        numMontoBaseCalculo As Double, _
                                        datFechaInicial As Date, _
                                        datFechaFinal As Date) As Double

    Dim intNumPeriodoAnualTasa As Integer
    Dim intDiasProvision       As Integer
    Dim intDiasBaseAnual       As Integer
    Dim numPorcenTasaAnual     As Double
    Dim numMontoCalculoInteres As Double
    Dim adoConsulta            As ADODB.Recordset
        
    With adoComm
        Set adoConsulta = New ADODB.Recordset
    
        .CommandText = "SELECT ValorParametro FROM AuxiliarParametro WHERE CodTipoParametro='TIPFRE' AND CodParametro='" & strCodPeriodoTasa & "'"
        Set adoConsulta = .Execute
    
        If Not adoConsulta.EOF Then
            intNumPeriodoAnualTasa = CInt(360 / adoConsulta("ValorParametro"))     '*** Numero del periodos por año de la tasa ***
        End If

        adoConsulta.Close: Set adoConsulta = Nothing
    End With
   
    Select Case strCodBaseCalculo

        Case Codigo_Base_30_360:
            intDiasBaseAnual = 360
            intDiasProvision = DateDiff("d", datFechaInicial, datFechaFinal)

        Case Codigo_Base_Actual_365:
            intDiasBaseAnual = 365
            intDiasProvision = DateDiff("d", datFechaInicial, datFechaFinal) + 1

        Case Codigo_Base_Actual_360:
            intDiasBaseAnual = 360
            intDiasProvision = DateDiff("d", datFechaInicial, datFechaFinal) + 1

        Case Codigo_Base_30_365:
            intDiasBaseAnual = 365
            intDiasProvision = Dias360(datFechaInicial, datFechaFinal, True)
    End Select

    Select Case strCodTipoTasa

        Case Codigo_Tipo_Tasa_Efectiva:
            numPorcenTasaAnual = (1 + (numPorcenTasa / 100)) ^ (intNumPeriodoAnualTasa) - 1
            numMontoCalculoInteres = Round(numMontoBaseCalculo * ((((1 + numPorcenTasaAnual)) ^ (intDiasProvision / intDiasBaseAnual)) - 1), 2) 'adoRegistro("MontoDevengo") + curMontoRenta

        Case Codigo_Tipo_Tasa_Nominal:
            numPorcenTasaAnual = (numPorcenTasa / 100) * intNumPeriodoAnualTasa
            numMontoCalculoInteres = Round(numMontoBaseCalculo * ((numPorcenTasaAnual * (intDiasProvision / intDiasBaseAnual))), 2)

        Case Codigo_Tipo_Tasa_Flat:
            numPorcenTasaAnual = numPorcenTasa / 100
            numMontoCalculoInteres = Round(numMontoBaseCalculo * (numPorcenTasaAnual), 2)
    End Select

    CalculoInteresDescuento = numMontoCalculoInteres

End Function

Public Sub CargarSolicitud(strpCodFondoOrden As String, _
                           strpCodAdministradora As String, _
                           strpNumOperacionOrig As String, _
                           intpIndCartera As Integer)
    
    If intpIndCartera = 1 Then Form_Load
    
    Dim intRegistro      As Integer
    
    Dim adoOperacionOrig As ADODB.Recordset
    
    Set adoOperacionOrig = New ADODB.Recordset

    With adoComm
    
        .CommandText = "SELECT TOP(1) ISO.CodFondo          ,ISO.CodAdministradora  ,ISO.NumSolicitud   ,ISO.FechaSolicitud    ,ISO.CodTitulo" & _
           ",ISO.EstadoSolicitud   ,ISO.CodFile            ,ISO.CodAnalitica   ,ISO.CodDetalleFile    ,ISO.CodSubDetalleFile" & _
           ",ISO.TipoSolicitud     ,ISO.DescripSolicitud   ,ISO.CodEmisor ,ISO.CodComisionista, ISO.NumSecuencialComisionistaCondicion       ,ISO.FechaConfirmacion ,ISO.FechaVencimiento" & _
           ",ISO.FechaLiquidacion  ,ISO.FechaEmision       ,ISO.CodMoneda      ,ISO.ValorTipoCambio   ,ISO.MontoSolicitud" & _
           ",ISO.MontoAprobado     ,ISO.TipoTasa           ,ISO.BaseAnual      ,ISO.TasaInteres       ,ISO.Observacion " & _
           ",ISO.MontoConsumido    ,ISO.CodAnalitica,       sum(ISCT.MontoInteresCuota) as InteresVencido , isnull(IICD.NumDesembolso,0) as NumDesembolso, " & _
           " isnull(IICD.FechaDesembolso,'" & gstrFechaActual & "') as FechaDesembolso, isnull(IICD.ValorDesembolso,0) as ValorDesembolso " & _
           "FROM InversionSolicitud ISO " & _
           " left join InstrumentoInversionCalendarioDesembolso IICD on(IICD.CodTitulo = ISO.CodTitulo and EstadoDesembolso = '01') " & _
           " join InversionOperacionCalendarioCuota ISCT on (ISO.CodFondo = ISCT.CodFondo and ISO.CodAdministradora = ISCT.CodAdministradora and " & _
           " ISO.NumSolicitud= ISCT.NumOperacionOrig and ISO.CodFile = ISCT.CodFile and ISO.CodAnalitica = ISCT.CodAnalitica and ISO.CodTitulo = ISCT.CodTitulo and ISCT.NumDesembolso = isnull(IICD.NumDesembolso,0)) " & _
           "WHERE ISO.CodFondo = '" & strCodFondoDescuento & "' AND ISO.CodAdministradora = '" & strpCodAdministradora & "' AND ISO.NumSolicitud='" & strpNumOperacionOrig & "' " & _
           "group by ISO.CodFondo,ISO.CodAdministradora,ISO.NumSolicitud,ISO.FechaSolicitud,ISO.CodTitulo,ISO.EstadoSolicitud,ISO.CodFile," & _
           "ISO.CodAnalitica,ISO.CodDetalleFile,ISO.CodSubDetalleFile,ISO.TipoSolicitud,ISO.DescripSolicitud,ISO.CodEmisor,ISO.CodComisionista, ISO.NumSecuencialComisionistaCondicion  ,ISO.FechaConfirmacion," & _
           "ISO.FechaVencimiento,ISO.FechaLiquidacion,ISO.FechaEmision,ISO.CodMoneda,ISO.ValorTipoCambio,ISO.MontoSolicitud,ISO.MontoAprobado,ISO.TipoTasa," & _
           "ISO.BaseAnual,ISO.TasaInteres,ISO.Observacion,ISO.MontoConsumido,ISO.CodAnalitica,IICD.NumDesembolso, IICD.FechaDesembolso, IICD.ValorDesembolso" & _
           " order by IICD.NumDesembolso "

        
        Set adoOperacionOrig = .Execute

        If Not adoOperacionOrig.EOF Then

            intRegistro = ObtenerItemLista(arrFondoOrden(), adoOperacionOrig.Fields("CodFondo"))

            If intRegistro >= 0 Then cboFondoOrden.ListIndex = intRegistro
            
            intRegistro = ObtenerItemLista(arrTipoInstrumentoOrden(), adoOperacionOrig.Fields("CodFile"))

            If intRegistro >= 0 Then cboTipoInstrumentoOrden.ListIndex = intRegistro
                                        
            intRegistro = ObtenerItemLista(arrClaseInstrumento(), adoOperacionOrig.Fields("CodDetalleFile"))

            If intRegistro >= 0 Then cboClaseInstrumento.ListIndex = intRegistro
                
            intRegistro = ObtenerItemLista(arrSubClaseInstrumento(), adoOperacionOrig.Fields("CodSubDetalleFile"))

            If intRegistro >= 0 Then cboSubClaseInstrumento.ListIndex = intRegistro
                
            intRegistro = ObtenerItemLista(arrEmisor(), adoOperacionOrig.Fields("CodEmisor"))

            If intRegistro >= 0 Then cboEmisor.ListIndex = intRegistro
            
            strCodComisionista = adoOperacionOrig.Fields("CodComisionista")
            
            numSecCondicion = adoOperacionOrig.Fields("NumSecuencialComisionistaCondicion")

            txtTasa.Text = adoOperacionOrig.Fields("TasaInteres")
                
            intRegistro = ObtenerItemLista(arrBaseAnual(), adoOperacionOrig.Fields("BaseAnual"))

            If intRegistro >= 0 Then cboBaseAnual.ListIndex = intRegistro
                
            intRegistro = ObtenerItemLista(arrTipoTasa(), adoOperacionOrig.Fields("TipoTasa"))

            If intRegistro >= 0 Then cboTipoTasa.ListIndex = intRegistro

            intRegistro = ObtenerItemLista(arrMoneda(), adoOperacionOrig.Fields("CodMoneda"))

            If intRegistro >= 0 Then cboMoneda.ListIndex = intRegistro
                
            strCodTitulo = Trim$(adoOperacionOrig("CodTitulo"))
                
            strCodAnalitica = Trim$(adoOperacionOrig("CodAnalitica"))

            txtDescripOrden.Text = adoOperacionOrig.Fields("DescripSolicitud")

            txtNumAnexo.Text = adoOperacionOrig.Fields("Observacion")
            txtMontoSolicitud.Text = adoOperacionOrig.Fields("MontoAprobado")
            txtMontoConsumido.Text = adoOperacionOrig.Fields("MontoConsumido")
            txtMontoDesembolso.Text = adoOperacionOrig("ValorDesembolso")
            txtNumeroDesembolso.Text = adoOperacionOrig("NumDesembolso")
                
            dtpFechaVencimiento.Value = traerCampo("InstrumentoInversionCondicionesFinancieras", "FechaVencimiento", "CodTitulo", adoOperacionOrig.Fields("CodTitulo")) ', " CodFondo = '" & strpCodFondoOrden & "' AND CodAdminisradora = '" & strpCodAdministradora & "'")
            dtpFechaVencimiento_Change
            lblIntAdelantado(0).Caption = Trim$(adoOperacionOrig("InteresVencido"))
        End If

        adoOperacionOrig.Close: Set adoOperacionOrig = Nothing

    End With

End Sub

Public Function ObtenerTotalDeuda(strpCodFondoOrden As String, _
                                  strpCodAdministradora As String, _
                                  strpNumOperacionOrig As String, _
                                  gstrpLogin As String) As Double

    Dim adoConsulta       As ADODB.Recordset
    Dim dblMontoDeuda     As Double
    Dim strFechaPagoCuota As String

    ObtenerTotalDeuda = 0

    strFechaPagoCuota = Convertyyyymmdd(dtpFechaOrdenCancel.Value)

    Set adoConsulta = New ADODB.Recordset
    
    With adoComm

        .CommandText = "{ call up_ACCalcularDeudaTotal ('" & strpCodFondoOrden & "','" & strpCodAdministradora & "','" & strpNumOperacionOrig & "','" & strFechaPagoCuota & "','" & strFechaPagoCuota & "','" & gstrpLogin & "' ) }"

        Set adoConsulta = .Execute

        If Not adoConsulta.EOF Then
            If adoConsulta("Deuda") <> "" Then
                dblMontoDeuda = CDbl(adoConsulta("Deuda"))
            Else
                dblMontoDeuda = 0
            End If
        End If
    
        adoConsulta.Close: Set adoConsulta = Nothing

    End With
    
    ObtenerTotalDeuda = dblMontoDeuda

End Function

Public Function ObtenerInteresesAdicionales(strpCodFondoOrden As String, _
                                            strpCodAdministradora As String, _
                                            strpNumOperacionOrig As String, _
                                            gstrpLogin As String) As Double

    Dim adoConsulta       As ADODB.Recordset
    Dim dblInteresesAdic  As Double
    Dim strFechaPagoCuota As String

    ObtenerInteresesAdicionales = 0

    strFechaPagoCuota = Convertyyyymmdd(dtpFechaOrdenCancel.Value)

    Set adoConsulta = New ADODB.Recordset
    
    With adoComm

        .CommandText = "{ call up_ACCalcularInteresesAdicionales ('" & strpCodFondoOrden & "','" & strpCodAdministradora & "','" & strpNumOperacionOrig & "','" & strFechaPagoCuota & "','" & strFechaPagoCuota & "','" & gstrpLogin & "' ) }"

        Set adoConsulta = .Execute

        If Not adoConsulta.EOF Then
            If adoConsulta("DeudaInteresAdicional") <> "" Then
                dblInteresesAdic = CDbl(adoConsulta("DeudaInteresAdicional"))
            Else
                dblInteresesAdic = 0
            End If
        End If
    
        adoConsulta.Close: Set adoConsulta = Nothing

    End With
    
    ObtenerInteresesAdicionales = dblInteresesAdic

End Function

Public Sub HabilitaCombos(ByVal pBloquea As Boolean)

    cboFondoOrden.Enabled = pBloquea
    cboTipoInstrumentoOrden.Enabled = pBloquea
    cboClaseInstrumento.Enabled = pBloquea
    cboSubClaseInstrumento.Enabled = pBloquea
    cboTipoOrden.Enabled = pBloquea
    cboTitulo.Enabled = pBloquea
    cboEmisor.Enabled = pBloquea

    If (strCodTipoOrden <> Codigo_Orden_Compra) And (strCodTipoOrden <> Codigo_Orden_Renovacion) Then
        cboObligado.Enabled = pBloquea
    End If

    cboGestor.Enabled = pBloquea
    cboOperacion.Enabled = pBloquea
    cboOrigen.Enabled = pBloquea
    cboLineaCliente.Enabled = pBloquea

End Sub

Public Sub mostrarForm(ByVal strNumSolicitud As String)

    Load Me
    
    Adicionar
    
    txtNumOperacionOrig.Text = strNumSolicitud
    txtNumOperacionOrig_KeyPress 13
    
    If txtNumeroDesembolso.Value = 0 Then
        txtValorNominal.Text = txtMontoSolicitud.Value - txtMontoConsumido.Value
    Else
        txtValorNominal.Text = txtMontoDesembolso.Value
    End If
    
        
    Me.Show
End Sub

