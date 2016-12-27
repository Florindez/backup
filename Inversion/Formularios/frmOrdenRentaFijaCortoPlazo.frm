VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{830D5A06-9B70-4F8C-98B6-7A19C4E7760B}#1.0#0"; "TAMControls.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmOrdenRentaFijaCortoPlazo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ordenes - Al Contado Renta Fija Corto Plazo"
   ClientHeight    =   9795
   ClientLeft      =   1500
   ClientTop       =   1680
   ClientWidth     =   14505
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
   Icon            =   "frmOrdenRentaFijaCortoPlazo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9795
   ScaleWidth      =   14505
   Begin MSAdodcLib.Adodc adoConsulta 
      Height          =   330
      Left            =   10290
      Top             =   9390
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
   Begin TabDlg.SSTab tabRFCortoPlazo 
      Height          =   8955
      Left            =   90
      TabIndex        =   49
      Top             =   120
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   15796
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Lista"
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tdgConsulta"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraCriterio"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Datos Orden Inversión"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraResumen"
      Tab(1).Control(1)=   "fraDatosBasicos"
      Tab(1).Control(2)=   "fraDatosTitulo"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Negociación"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraComisionMontoFL2"
      Tab(2).Control(1)=   "fraComisionMontoFL1"
      Tab(2).Control(2)=   "fraDatosNegociacion"
      Tab(2).Control(3)=   "fraPosicion"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Garantias"
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame1"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame1 
         Caption         =   "Garantia"
         Height          =   1785
         Left            =   -74610
         TabIndex        =   197
         Top             =   660
         Width           =   8265
         Begin VB.TextBox txtGarantia 
            Height          =   345
            Left            =   2130
            TabIndex        =   212
            Top             =   450
            Width           =   4275
         End
         Begin VB.CheckBox chkFiador 
            Caption         =   "Fiador"
            Height          =   225
            Left            =   420
            TabIndex        =   211
            Top             =   960
            Width           =   1335
         End
         Begin VB.CheckBox chkGarantia 
            Caption         =   "Garantia"
            Height          =   225
            Left            =   420
            TabIndex        =   199
            Top             =   540
            Width           =   1335
         End
         Begin VB.ComboBox cboFiador 
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
            Left            =   2130
            Style           =   2  'Dropdown List
            TabIndex        =   198
            Top             =   930
            Width           =   4305
         End
      End
      Begin VB.Frame fraComisionMontoFL2 
         Caption         =   "Comisiones y Montos - Plazo (FL2)"
         Height          =   360
         Left            =   -67020
         TabIndex        =   156
         Top             =   3750
         Width           =   3255
         Begin VB.CheckBox chkAplicar 
            Caption         =   "Aplicar"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   37
            Top             =   720
            Width           =   975
         End
         Begin VB.CommandButton Command1 
            Caption         =   "#"
            Height          =   375
            Left            =   480
            TabIndex        =   45
            ToolTipText     =   "Calcular Valor al Vencimiento y TIRs de la orden"
            Top             =   4485
            Width           =   375
         End
         Begin VB.TextBox txtInteresCorrido 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   1
            Left            =   4290
            MaxLength       =   45
            TabIndex        =   44
            Top             =   3465
            Width           =   2025
         End
         Begin VB.TextBox txtComisionAgente 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   1
            Left            =   4290
            MaxLength       =   45
            TabIndex        =   39
            Top             =   1185
            Width           =   2025
         End
         Begin VB.TextBox txtComisionBolsa 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   1
            Left            =   4290
            MaxLength       =   45
            TabIndex        =   40
            Top             =   1530
            Width           =   2025
         End
         Begin VB.TextBox txtComisionCavali 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   1
            Left            =   4290
            MaxLength       =   45
            TabIndex        =   41
            Top             =   1890
            Width           =   2025
         End
         Begin VB.TextBox txtComisionFondo 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   1
            Left            =   4290
            MaxLength       =   45
            TabIndex        =   42
            Top             =   2250
            Width           =   2025
         End
         Begin VB.TextBox txtComisionConasev 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   1
            Left            =   4290
            MaxLength       =   45
            TabIndex        =   43
            Top             =   2610
            Width           =   2025
         End
         Begin VB.TextBox txtPorcenAgente 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   1
            Left            =   2625
            MaxLength       =   45
            TabIndex        =   38
            Top             =   1185
            Width           =   1340
         End
         Begin VB.TextBox txtPrecioUnitario 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Index           =   1
            Left            =   2625
            MaxLength       =   45
            TabIndex        =   36
            Top             =   360
            Width           =   1340
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   5
            X1              =   360
            X2              =   6300
            Y1              =   4200
            Y2              =   4200
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "TIR Neta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   81
            Left            =   2880
            TabIndex        =   181
            Top             =   4320
            Width           =   660
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "TIR Bruta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   80
            Left            =   1320
            TabIndex        =   180
            Top             =   4320
            Width           =   690
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Valor al Vencimiento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   79
            Left            =   4440
            TabIndex        =   179
            Top             =   4320
            Width           =   1440
         End
         Begin VB.Label lblPorcenConasev 
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
            Index           =   1
            Left            =   2625
            TabIndex        =   178
            Tag             =   "0"
            ToolTipText     =   "Porcentaje de Comisión Conasev"
            Top             =   2610
            Width           =   1335
         End
         Begin VB.Label lblPorcenFondo 
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
            Index           =   1
            Left            =   2625
            TabIndex        =   177
            Tag             =   "0"
            ToolTipText     =   "Porcentaje de Comisión Fondo Liquidación"
            Top             =   2250
            Width           =   1335
         End
         Begin VB.Label lblPorcenCavali 
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
            Index           =   1
            Left            =   2625
            TabIndex        =   176
            Tag             =   "0"
            ToolTipText     =   "Porcentaje de Comisión Cavali"
            Top             =   1890
            Width           =   1335
         End
         Begin VB.Label lblPorcenBolsa 
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
            Index           =   1
            Left            =   2625
            TabIndex        =   175
            Tag             =   "0"
            ToolTipText     =   "Porcentaje de Comisión BVL"
            Top             =   1530
            Width           =   1335
         End
         Begin VB.Label lblMontoTotal 
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
            Index           =   1
            Left            =   4290
            TabIndex        =   174
            Tag             =   "0.00"
            ToolTipText     =   "Monto Total de la Orden"
            Top             =   3825
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
         Begin VB.Label lblPorcenIgv 
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
            Index           =   1
            Left            =   2625
            TabIndex        =   173
            Tag             =   "0.00"
            ToolTipText     =   "Porcentaje de Comisión IGV"
            Top             =   2970
            Width           =   1335
         End
         Begin VB.Label lblComisionIgv 
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
            Index           =   1
            Left            =   4290
            TabIndex        =   172
            Tag             =   "0.00"
            ToolTipText     =   "Monto de Comisión IGV"
            Top             =   2970
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
         Begin VB.Label lblSubTotal 
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
            Index           =   1
            Left            =   4290
            TabIndex        =   171
            Tag             =   "0.00"
            ToolTipText     =   "SubTotal de la Orden"
            Top             =   720
            Width           =   2025
         End
         Begin VB.Label Label3 
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
            Left            =   2625
            TabIndex        =   170
            Tag             =   "0.00"
            Top             =   4590
            Width           =   1335
         End
         Begin VB.Label Label2 
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
            Left            =   1080
            TabIndex        =   169
            Tag             =   "0.00"
            Top             =   4590
            Width           =   1335
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
            Left            =   4290
            TabIndex        =   168
            Tag             =   "0.00"
            Top             =   4590
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   4365
            TabIndex        =   167
            Top             =   240
            Width           =   1845
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto Total"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   76
            Left            =   2640
            TabIndex        =   166
            Top             =   3840
            Width           =   855
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Interés Corrido"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   73
            Left            =   2640
            TabIndex        =   165
            Top             =   3480
            Width           =   1020
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "SubTotal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   72
            Left            =   2640
            TabIndex        =   164
            Top             =   735
            Width           =   645
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "IGV"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   62
            Left            =   390
            TabIndex        =   163
            Top             =   3030
            Width           =   270
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisión Conasev"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   60
            Left            =   390
            TabIndex        =   162
            Top             =   2670
            Width           =   1305
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisión Fondo Garantía"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   59
            Left            =   390
            TabIndex        =   161
            Top             =   2295
            Width           =   1800
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisión Cavali"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   58
            Left            =   390
            TabIndex        =   160
            Top             =   1935
            Width           =   1110
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisión BVL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   57
            Left            =   390
            TabIndex        =   159
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisión SAB"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   56
            Left            =   390
            TabIndex        =   158
            Top             =   1200
            Width           =   990
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Precio (%)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   55
            Left            =   390
            TabIndex        =   157
            Top             =   375
            Width           =   705
         End
      End
      Begin VB.Frame fraResumen 
         Caption         =   "Resumen Negociación"
         Height          =   1935
         Left            =   -74730
         TabIndex        =   123
         Top             =   5400
         Width           =   13905
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Analítica"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   84
            Left            =   10200
            TabIndex        =   188
            Top             =   380
            Width           =   630
         End
         Begin VB.Label lblAnalitica 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "???-????????"
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
            Left            =   11400
            TabIndex        =   187
            Tag             =   "0.00"
            Top             =   360
            Width           =   2025
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Precio"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   54
            Left            =   5280
            TabIndex        =   155
            Top             =   1055
            Width           =   450
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Precio"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   53
            Left            =   360
            TabIndex        =   154
            Top             =   1050
            Width           =   450
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "SubTotal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   61
            Left            =   390
            TabIndex        =   153
            Top             =   3000
            Width           =   645
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisiones"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   63
            Left            =   390
            TabIndex        =   152
            Top             =   3330
            Width           =   795
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Intereses Corridos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   64
            Left            =   390
            TabIndex        =   151
            Top             =   3660
            Width           =   1260
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto Total"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   65
            Left            =   360
            TabIndex        =   150
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "SubTotal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   66
            Left            =   5190
            TabIndex        =   149
            Top             =   3000
            Width           =   645
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisiones"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   67
            Left            =   5190
            TabIndex        =   148
            Top             =   3330
            Width           =   795
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Intereses Corridos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   68
            Left            =   5190
            TabIndex        =   147
            Top             =   3660
            Width           =   1260
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto Total"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   69
            Left            =   5250
            TabIndex        =   146
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label lblPrecioResumen 
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
            Index           =   0
            Left            =   2400
            TabIndex        =   145
            Tag             =   "0.00"
            Top             =   1035
            Width           =   2025
         End
         Begin VB.Label lblSubTotalResumen 
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
            Index           =   0
            Left            =   2310
            TabIndex        =   144
            Tag             =   "0.00"
            Top             =   2985
            Width           =   2025
         End
         Begin VB.Label lblComisionesResumen 
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
            Index           =   0
            Left            =   2310
            TabIndex        =   143
            Tag             =   "0.00"
            Top             =   3315
            Width           =   2025
         End
         Begin VB.Label lblInteresesResumen 
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
            Index           =   0
            Left            =   2310
            TabIndex        =   142
            Tag             =   "0.00"
            Top             =   3645
            Width           =   2025
         End
         Begin VB.Label lblTotalResumen 
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
            Index           =   0
            Left            =   2400
            TabIndex        =   141
            Tag             =   "0.00"
            Top             =   1425
            Width           =   2025
         End
         Begin VB.Label lblPrecioResumen 
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
            Index           =   1
            Left            =   7320
            TabIndex        =   140
            Tag             =   "0.00"
            Top             =   1035
            Width           =   2025
         End
         Begin VB.Label lblSubTotalResumen 
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
            Index           =   1
            Left            =   7230
            TabIndex        =   139
            Tag             =   "0.00"
            Top             =   2985
            Width           =   2025
         End
         Begin VB.Label lblComisionesResumen 
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
            Index           =   1
            Left            =   7230
            TabIndex        =   138
            Tag             =   "0.00"
            Top             =   3315
            Width           =   2025
         End
         Begin VB.Label lblInteresesResumen 
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
            Index           =   1
            Left            =   7230
            TabIndex        =   137
            Tag             =   "0.00"
            Top             =   3645
            Width           =   2025
         End
         Begin VB.Label lblTotalResumen 
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
            Index           =   1
            Left            =   7320
            TabIndex        =   136
            Tag             =   "0.00"
            Top             =   1425
            Width           =   2025
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Contado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   70
            Left            =   360
            TabIndex        =   135
            Top             =   720
            Width           =   600
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Plazo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   71
            Left            =   5280
            TabIndex        =   134
            Top             =   720
            Width           =   390
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tir Bruta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   74
            Left            =   10200
            TabIndex        =   133
            Top             =   960
            Width           =   600
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tir Neta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   75
            Left            =   10200
            TabIndex        =   132
            Top             =   1320
            Width           =   570
         End
         Begin VB.Label lblTirBrutaResumen 
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
            Left            =   11400
            TabIndex        =   131
            Tag             =   "0.00"
            Top             =   960
            Width           =   2025
         End
         Begin VB.Label lblTirNetaResumen 
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
            Left            =   11400
            TabIndex        =   130
            Tag             =   "0.00"
            Top             =   1320
            Width           =   2025
         End
         Begin VB.Line Line4 
            BorderColor     =   &H80000015&
            X1              =   4800
            X2              =   4800
            Y1              =   360
            Y2              =   1620
         End
         Begin VB.Label lblCantidadResumen 
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
            Left            =   2400
            TabIndex        =   129
            Tag             =   "0.00"
            Top             =   360
            Width           =   2025
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad Facial"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   77
            Left            =   360
            TabIndex        =   128
            Top             =   375
            Width           =   1095
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Vencimiento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   78
            Left            =   5280
            TabIndex        =   127
            Top             =   380
            Width           =   870
         End
         Begin VB.Label lblVencimientoResumen 
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
            Left            =   7320
            TabIndex        =   126
            Tag             =   "0.00"
            Top             =   360
            Width           =   2025
         End
         Begin VB.Label lblDescripMonedaResumen 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Nuevos Soles (S/.)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   2400
            TabIndex        =   125
            Top             =   720
            Width           =   2025
         End
         Begin VB.Label lblDescripMonedaResumen 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Nuevos Soles (S/.)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   7320
            TabIndex        =   124
            Top             =   720
            Width           =   1845
         End
         Begin VB.Line Line3 
            BorderColor     =   &H80000015&
            X1              =   9720
            X2              =   9720
            Y1              =   360
            Y2              =   1680
         End
      End
      Begin VB.Frame fraComisionMontoFL1 
         Caption         =   "Comisiones y Montos - Plazo (FL1)"
         Height          =   5160
         Left            =   -74850
         TabIndex        =   89
         Top             =   3720
         Width           =   6735
         Begin TAMControls.TAMTextBox txtTirBruta1 
            Height          =   315
            Left            =   390
            TabIndex        =   213
            Top             =   4560
            Width           =   1965
            _ExtentX        =   3466
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
            Container       =   "frmOrdenRentaFijaCortoPlazo.frx":0442
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
         Begin VB.TextBox txtTirNeta 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Left            =   2610
            MaxLength       =   45
            TabIndex        =   210
            Top             =   5340
            Width           =   1365
         End
         Begin VB.TextBox txtPrecioUnitario 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Index           =   0
            Left            =   240
            MaxLength       =   45
            TabIndex        =   26
            Top             =   5160
            Width           =   1340
         End
         Begin VB.TextBox txtPorcenAgente 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   2625
            MaxLength       =   45
            TabIndex        =   28
            Top             =   1185
            Width           =   1340
         End
         Begin VB.TextBox txtComisionConasev 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   4290
            MaxLength       =   45
            TabIndex        =   33
            Top             =   2610
            Width           =   2025
         End
         Begin VB.TextBox txtComisionFondo 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   4290
            MaxLength       =   45
            TabIndex        =   32
            Top             =   2250
            Width           =   2025
         End
         Begin VB.TextBox txtComisionCavali 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   4290
            MaxLength       =   45
            TabIndex        =   31
            Top             =   1890
            Width           =   2025
         End
         Begin VB.TextBox txtComisionBolsa 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   4290
            MaxLength       =   45
            TabIndex        =   30
            Top             =   1530
            Width           =   2025
         End
         Begin VB.TextBox txtComisionAgente 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   4290
            MaxLength       =   45
            TabIndex        =   29
            Top             =   1185
            Width           =   2025
         End
         Begin VB.TextBox txtInteresCorrido 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   4290
            MaxLength       =   45
            TabIndex        =   34
            Top             =   3465
            Width           =   2025
         End
         Begin VB.CommandButton cmdCalculo 
            Caption         =   "#"
            Height          =   375
            Left            =   510
            TabIndex        =   35
            ToolTipText     =   "Calcular Valor al Vencimiento y TIRs de la orden"
            Top             =   5310
            Width           =   375
         End
         Begin VB.CheckBox chkAplicar 
            Caption         =   "Aplicar"
            Height          =   255
            Index           =   0
            Left            =   390
            TabIndex        =   27
            Top             =   720
            Width           =   975
         End
         Begin TAMControls.TAMTextBox txtMontoVencimiento1 
            Height          =   315
            Left            =   4290
            TabIndex        =   214
            Top             =   4560
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
            Container       =   "frmOrdenRentaFijaCortoPlazo.frx":045E
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
            Left            =   2640
            TabIndex        =   215
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
            Container       =   "frmOrdenRentaFijaCortoPlazo.frx":047A
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
            Caption         =   "Precio (%)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   18
            Left            =   390
            TabIndex        =   119
            Top             =   375
            Width           =   705
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisión SAB"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   24
            Left            =   390
            TabIndex        =   113
            Top             =   1200
            Width           =   990
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisión BVL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   26
            Left            =   390
            TabIndex        =   112
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisión Cavali"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   27
            Left            =   390
            TabIndex        =   111
            Top             =   1935
            Width           =   1110
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisión Fondo Garantía"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   32
            Left            =   390
            TabIndex        =   110
            Top             =   2295
            Width           =   1800
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Comisión Conasev"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   33
            Left            =   390
            TabIndex        =   109
            Top             =   2670
            Width           =   1305
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "IGV"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   34
            Left            =   390
            TabIndex        =   108
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   25
            Left            =   2640
            TabIndex        =   107
            Top             =   735
            Width           =   645
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Interés Corrido"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   35
            Left            =   420
            TabIndex        =   106
            Top             =   3480
            Width           =   2160
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto Total"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   36
            Left            =   420
            TabIndex        =   105
            Top             =   3840
            Width           =   855
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   4365
            TabIndex        =   104
            Top             =   360
            Width           =   1845
         End
         Begin VB.Label lblMontoVencimiento 
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
            Left            =   4320
            TabIndex        =   103
            Tag             =   "0.00"
            Top             =   5370
            Width           =   2025
         End
         Begin VB.Label lblTirBruta 
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
            Left            =   1080
            TabIndex        =   102
            Tag             =   "0.00"
            Top             =   5340
            Width           =   1335
         End
         Begin VB.Label lblTirNeta 
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
            Left            =   2610
            TabIndex        =   101
            Tag             =   "0.00"
            Top             =   4560
            Width           =   1425
         End
         Begin VB.Label lblSubTotal 
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
            Index           =   0
            Left            =   4290
            TabIndex        =   100
            Tag             =   "0.00"
            ToolTipText     =   "SubTotal de la Orden"
            Top             =   720
            Width           =   2025
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   0
            X1              =   2580
            X2              =   6300
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Label lblComisionIgv 
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
            Index           =   0
            Left            =   4290
            TabIndex        =   99
            Tag             =   "0.00"
            ToolTipText     =   "Monto de Comisión IGV"
            Top             =   2970
            Width           =   2025
         End
         Begin VB.Label lblPorcenIgv 
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
            Index           =   0
            Left            =   2625
            TabIndex        =   98
            Tag             =   "0.00"
            ToolTipText     =   "Porcentaje de Comisión IGV"
            Top             =   2970
            Width           =   1335
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   1
            X1              =   360
            X2              =   6300
            Y1              =   3345
            Y2              =   3345
         End
         Begin VB.Label lblMontoTotal 
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
            Index           =   0
            Left            =   4290
            TabIndex        =   97
            Tag             =   "0.00"
            ToolTipText     =   "Monto Total de la Orden"
            Top             =   3825
            Width           =   2025
         End
         Begin VB.Label lblPorcenBolsa 
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
            Index           =   0
            Left            =   2625
            TabIndex        =   96
            Tag             =   "0"
            ToolTipText     =   "Porcentaje de Comisión BVL"
            Top             =   1530
            Width           =   1335
         End
         Begin VB.Label lblPorcenCavali 
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
            Index           =   0
            Left            =   2625
            TabIndex        =   95
            Tag             =   "0"
            ToolTipText     =   "Porcentaje de Comisión Cavali"
            Top             =   1890
            Width           =   1335
         End
         Begin VB.Label lblPorcenFondo 
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
            Index           =   0
            Left            =   2625
            TabIndex        =   94
            Tag             =   "0"
            ToolTipText     =   "Porcentaje de Comisión Fondo Liquidación"
            Top             =   2250
            Width           =   1335
         End
         Begin VB.Label lblPorcenConasev 
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
            Index           =   0
            Left            =   2625
            TabIndex        =   93
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   37
            Left            =   4620
            TabIndex        =   92
            Top             =   4320
            Width           =   1440
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "TIR Bruta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   38
            Left            =   1110
            TabIndex        =   91
            Top             =   4320
            Width           =   690
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "TIR Neta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   39
            Left            =   2970
            TabIndex        =   90
            Top             =   4320
            Width           =   660
         End
         Begin VB.Line linSeparador 
            BorderColor     =   &H80000015&
            Index           =   2
            X1              =   360
            X2              =   6300
            Y1              =   4200
            Y2              =   4200
         End
      End
      Begin VB.Frame fraDatosNegociacion 
         Caption         =   "Negociación"
         Height          =   3165
         Left            =   -74820
         TabIndex        =   87
         Top             =   510
         Width           =   9135
         Begin VB.TextBox txtValorNominalDcto 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Left            =   2160
            MaxLength       =   45
            TabIndex        =   217
            Top             =   2280
            Width           =   1900
         End
         Begin VB.ComboBox cboConceptoCosto 
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
            Left            =   6570
            Style           =   2  'Dropdown List
            TabIndex        =   202
            Top             =   690
            Width           =   2295
         End
         Begin VB.ComboBox cboNegociacion 
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
            Left            =   6570
            Style           =   2  'Dropdown List
            TabIndex        =   200
            Top             =   300
            Width           =   2295
         End
         Begin VB.TextBox txtTasa 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Left            =   2160
            MaxLength       =   45
            TabIndex        =   23
            Top             =   330
            Width           =   1900
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
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   1095
            Width           =   1900
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
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   690
            Width           =   1900
         End
         Begin VB.TextBox txtTipoCambio 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Left            =   6570
            MaxLength       =   45
            TabIndex        =   25
            Top             =   2640
            Width           =   1830
         End
         Begin VB.TextBox txtCantidad 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Left            =   2160
            MaxLength       =   45
            TabIndex        =   24
            Top             =   2640
            Width           =   1900
         End
         Begin TAMControls.TAMTextBox txtPorcenDctoValorNominal 
            Height          =   315
            Left            =   2160
            TabIndex        =   219
            Top             =   1920
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
            Container       =   "frmOrdenRentaFijaCortoPlazo.frx":0496
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
            TabIndex        =   220
            Top             =   1560
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
            Container       =   "frmOrdenRentaFijaCortoPlazo.frx":04B2
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
            Caption         =   "Valor Nominal Dcto."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   90
            Left            =   360
            TabIndex        =   218
            Top             =   2310
            Width           =   1410
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "% V.Nominal Dcto."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   89
            Left            =   360
            TabIndex        =   216
            Top             =   1950
            Width           =   1320
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Concepto Costo Neg."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   82
            Left            =   4770
            TabIndex        =   203
            Top             =   720
            Width           =   1530
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Mecanismo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   31
            Left            =   4770
            TabIndex        =   201
            Top             =   390
            Width           =   810
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Plazo (Días)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   83
            Left            =   4800
            TabIndex        =   186
            Top             =   2325
            Width           =   870
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Vencimiento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   50
            Left            =   4800
            TabIndex        =   185
            Top             =   1935
            Width           =   870
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Emisión"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   48
            Left            =   4800
            TabIndex        =   184
            Top             =   1575
            Width           =   540
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Liquidación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   47
            Left            =   4800
            TabIndex        =   183
            Top             =   1215
            Width           =   810
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Valor Nominal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   12
            Left            =   360
            TabIndex        =   182
            Top             =   1590
            Width           =   975
         End
         Begin VB.Label lblDiasPlazo 
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
            Left            =   6570
            TabIndex        =   122
            Tag             =   "0.00"
            ToolTipText     =   "Días de Plazo del Título de la Orden"
            Top             =   2280
            Width           =   1815
         End
         Begin VB.Label lblFechaVencimiento 
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
            Left            =   6570
            TabIndex        =   121
            Tag             =   "0.00"
            ToolTipText     =   "Fecha de Vencimiento del Título de la Orden"
            Top             =   1920
            Width           =   1815
         End
         Begin VB.Label lblFechaEmision 
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
            Left            =   6570
            TabIndex        =   120
            Tag             =   "0.00"
            ToolTipText     =   "Fecha Emisión"
            Top             =   1560
            Width           =   1815
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000015&
            X1              =   4470
            X2              =   4470
            Y1              =   330
            Y2              =   2940
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tasa Facial"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   11
            Left            =   360
            TabIndex        =   118
            Top             =   375
            Width           =   825
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Base Anual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   40
            Left            =   360
            TabIndex        =   117
            Top             =   1110
            Width           =   810
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Tasa"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   41
            Left            =   360
            TabIndex        =   116
            Top             =   735
            Width           =   720
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Cambio"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   17
            Left            =   4800
            TabIndex        =   115
            Top             =   2685
            Width           =   885
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad Facial"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   15
            Left            =   360
            TabIndex        =   114
            Top             =   2640
            Width           =   1095
         End
         Begin VB.Label lblFechaLiquidacion 
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
            Left            =   6570
            TabIndex        =   88
            Tag             =   "0.00"
            ToolTipText     =   "Fecha de Liquidación"
            Top             =   1200
            Width           =   1815
         End
      End
      Begin VB.Frame fraPosicion 
         Caption         =   "Datos Posición"
         Height          =   3135
         Left            =   -65520
         TabIndex        =   76
         Top             =   540
         Width           =   4695
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Cupón Vigente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   52
            Left            =   480
            TabIndex        =   86
            Top             =   380
            Width           =   1050
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Clasificación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   5
            Left            =   480
            TabIndex        =   85
            Top             =   740
            Width           =   885
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Base - Tasa %"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   51
            Left            =   480
            TabIndex        =   84
            Top             =   1100
            Width           =   1020
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Stock Nominal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   16
            Left            =   480
            TabIndex        =   83
            Top             =   1460
            Width           =   1035
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   49
            Left            =   480
            TabIndex        =   82
            Top             =   1820
            Width           =   585
         End
         Begin VB.Label lblFechaCupon 
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
            Left            =   2160
            TabIndex        =   81
            Tag             =   "0.00"
            Top             =   360
            Width           =   2025
         End
         Begin VB.Label lblClasificacion 
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
            Left            =   2160
            TabIndex        =   80
            Tag             =   "0.00"
            Top             =   720
            Width           =   2025
         End
         Begin VB.Label lblBaseTasaCupon 
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
            Left            =   2160
            TabIndex        =   79
            Tag             =   "0.00"
            Top             =   1080
            Width           =   2025
         End
         Begin VB.Label lblStockNominal 
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
            Left            =   2160
            TabIndex        =   78
            Tag             =   "0.00"
            Top             =   1440
            Width           =   2025
         End
         Begin VB.Label lblMoneda 
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
            Left            =   2160
            TabIndex        =   77
            Tag             =   "0.00"
            ToolTipText     =   "Moneda del Título"
            Top             =   1800
            Width           =   2025
         End
      End
      Begin VB.Frame fraDatosBasicos 
         Caption         =   "Datos Básicos"
         Height          =   2400
         Left            =   -74760
         TabIndex        =   68
         Top             =   450
         Width           =   13935
         Begin VB.ComboBox cboOperacion 
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
            TabIndex        =   206
            Top             =   1440
            Width           =   4185
         End
         Begin VB.ComboBox cboOrigen 
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
            Left            =   9315
            Style           =   2  'Dropdown List
            TabIndex        =   204
            Top             =   1800
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
            Left            =   9300
            Style           =   2  'Dropdown List
            TabIndex        =   195
            Top             =   1080
            Width           =   4185
         End
         Begin VB.ComboBox cboObligado 
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
            TabIndex        =   193
            Top             =   720
            Width           =   4185
         End
         Begin VB.ComboBox cboSubClaseInstrumento 
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
            Left            =   2370
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1455
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
            Left            =   2370
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   360
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
            Left            =   2370
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   720
            Width           =   4185
         End
         Begin VB.ComboBox cboTipoOrden 
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
            Left            =   2370
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1830
            Width           =   4185
         End
         Begin VB.ComboBox cboTitulo 
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
            TabIndex        =   13
            Top             =   360
            Width           =   4185
         End
         Begin VB.ComboBox cboClaseInstrumento 
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
            Left            =   2370
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1095
            Width           =   4185
         End
         Begin VB.ComboBox cboEmisor 
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
            TabIndex        =   12
            Top             =   360
            Width           =   4185
         End
         Begin VB.CheckBox chkTitulo 
            Height          =   255
            Left            =   13560
            TabIndex        =   46
            ToolTipText     =   "Seleccionar Título"
            Top             =   360
            Width           =   255
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Liquidación Operación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   28
            Left            =   7170
            TabIndex        =   207
            Top             =   1500
            Width           =   1590
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Mercado Negociación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   4
            Left            =   7170
            TabIndex        =   205
            Top             =   1875
            Width           =   1575
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Gestor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   88
            Left            =   7170
            TabIndex        =   196
            Top             =   1125
            Width           =   465
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Obligado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   87
            Left            =   7170
            TabIndex        =   194
            Top             =   765
            Width           =   630
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Clase"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   29
            Left            =   360
            TabIndex        =   74
            Top             =   1114
            Width           =   390
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fondo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   73
            Top             =   380
            Width           =   450
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Instrumento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   72
            Top             =   747
            Width           =   825
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Orden de"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   71
            Top             =   1850
            Width           =   660
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Emisor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   6
            Left            =   7170
            TabIndex        =   70
            Top             =   405
            Width           =   465
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "SubClase"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   42
            Left            =   360
            TabIndex        =   69
            Top             =   1481
            Width           =   675
         End
      End
      Begin VB.Frame fraCriterio 
         Caption         =   "Criterios de Búsqueda"
         Height          =   1815
         Left            =   240
         TabIndex        =   58
         Top             =   420
         Width           =   13935
         Begin VB.CommandButton cmdEnviar 
            Caption         =   "En&viar"
            Height          =   375
            Left            =   12200
            TabIndex        =   190
            ToolTipText     =   "Enviar a BackOffice"
            Top             =   1200
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
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1200
            Width           =   5145
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
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   780
            Width           =   5145
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
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   360
            Width           =   5145
         End
         Begin MSComCtl2.DTPicker dtpFechaOrdenDesde 
            Height          =   285
            Left            =   9600
            TabIndex        =   3
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
            Format          =   176095233
            CurrentDate     =   38785
         End
         Begin MSComCtl2.DTPicker dtpFechaOrdenHasta 
            Height          =   285
            Left            =   11955
            TabIndex        =   4
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
            Format          =   176095233
            CurrentDate     =   38785
         End
         Begin MSComCtl2.DTPicker dtpFechaLiquidacionDesde 
            Height          =   285
            Left            =   9600
            TabIndex        =   5
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
            Format          =   176095233
            CurrentDate     =   38785
         End
         Begin MSComCtl2.DTPicker dtpFechaLiquidacionHasta 
            Height          =   285
            Left            =   11955
            TabIndex        =   6
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
            Format          =   176095233
            CurrentDate     =   38785
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Estado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   23
            Left            =   480
            TabIndex        =   67
            Top             =   1220
            Width           =   495
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Instrumento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   22
            Left            =   480
            TabIndex        =   66
            Top             =   800
            Width           =   825
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   21
            Left            =   11280
            TabIndex        =   65
            Top             =   380
            Width           =   420
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   20
            Left            =   8880
            TabIndex        =   64
            Top             =   380
            Width           =   465
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fondo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   19
            Left            =   480
            TabIndex        =   63
            Top             =   380
            Width           =   450
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Orden"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   43
            Left            =   7200
            TabIndex        =   62
            Top             =   380
            Width           =   930
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Liquidación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   44
            Left            =   7200
            TabIndex        =   61
            Top             =   800
            Width           =   1305
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   45
            Left            =   8880
            TabIndex        =   60
            Top             =   800
            Width           =   465
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   46
            Left            =   11280
            TabIndex        =   59
            Top             =   800
            Width           =   420
         End
      End
      Begin VB.Frame fraDatosTitulo 
         Caption         =   "Datos de la Orden"
         Height          =   2445
         Left            =   -74760
         TabIndex        =   50
         Top             =   2850
         Width           =   13935
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
            Height          =   765
            Left            =   2370
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   208
            Top             =   1470
            Width           =   11130
         End
         Begin VB.TextBox txtNemonico 
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
            Height          =   285
            Left            =   11880
            MaxLength       =   15
            TabIndex        =   20
            Top             =   705
            Width           =   1575
         End
         Begin VB.TextBox txtDiasPlazo 
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
            Left            =   8400
            TabIndex        =   18
            Top             =   705
            Width           =   1280
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
            Left            =   2370
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   720
            Width           =   2640
         End
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
            Left            =   2370
            MaxLength       =   45
            TabIndex        =   16
            Top             =   1100
            Width           =   4170
         End
         Begin MSComCtl2.DTPicker dtpFechaOrden 
            Height          =   315
            Left            =   2370
            TabIndex        =   47
            Top             =   330
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
            Format          =   176095233
            CurrentDate     =   38776
         End
         Begin MSComCtl2.DTPicker dtpFechaLiquidacion 
            Height          =   315
            Left            =   5160
            TabIndex        =   14
            Top             =   330
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
            Format          =   176095233
            CurrentDate     =   38776
         End
         Begin MSComCtl2.DTPicker dtpFechaEmision 
            Height          =   315
            Left            =   8400
            TabIndex        =   48
            Top             =   345
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
            Format          =   176095233
            CurrentDate     =   38776
         End
         Begin MSComCtl2.UpDown updDiasPlazo 
            Height          =   315
            Left            =   9675
            TabIndex        =   19
            Top             =   705
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   556
            _Version        =   393216
            BuddyControl    =   "txtDiasPlazo"
            BuddyDispid     =   196700
            OrigLeft        =   3360
            OrigTop         =   3960
            OrigRight       =   3615
            OrigBottom      =   4245
            Max             =   360
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.DTPicker dtpFechaVencimiento 
            Height          =   315
            Left            =   11880
            TabIndex        =   17
            Top             =   345
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
            Format          =   176095233
            CurrentDate     =   38776
         End
         Begin MSComCtl2.DTPicker dtpFechaPago 
            Height          =   315
            Left            =   8400
            TabIndex        =   192
            Top             =   1095
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
            Format          =   176095233
            CurrentDate     =   38776
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Instrucciones"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   30
            Left            =   360
            TabIndex        =   209
            Top             =   1470
            Width           =   945
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Pago"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   86
            Left            =   7080
            TabIndex        =   191
            Top             =   1120
            Width           =   375
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Nemónico"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   85
            Left            =   10320
            TabIndex        =   189
            Top             =   725
            Width           =   720
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Emisión"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   8
            Left            =   7080
            TabIndex        =   57
            Top             =   360
            Width           =   540
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   360
            TabIndex        =   56
            Top             =   750
            Width           =   585
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Plazo (DIAS)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   10
            Left            =   7080
            TabIndex        =   55
            Top             =   720
            Width           =   900
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Vencimiento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   9
            Left            =   10320
            TabIndex        =   54
            Top             =   360
            Width           =   870
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Descripción"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   7
            Left            =   360
            TabIndex        =   53
            Top             =   1120
            Width           =   840
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Liquidación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   14
            Left            =   4140
            TabIndex        =   52
            Top             =   375
            Width           =   810
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Orden"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   13
            Left            =   360
            TabIndex        =   51
            Top             =   375
            Width           =   435
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmOrdenRentaFijaCortoPlazo.frx":04CE
         Height          =   5835
         Left            =   240
         OleObjectBlob   =   "frmOrdenRentaFijaCortoPlazo.frx":04E8
         TabIndex        =   75
         Top             =   2520
         Width           =   13905
      End
   End
   Begin TAMControls2.ucBotonEdicion2 cmdOpcion 
      Height          =   735
      Left            =   420
      TabIndex        =   222
      Top             =   9090
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
   Begin TAMControls2.ucBotonEdicion2 cmdAccion 
      Height          =   735
      Left            =   6180
      TabIndex        =   223
      Top             =   9090
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
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   12570
      TabIndex        =   221
      Top             =   9090
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
      Caption0        =   "&Salir"
      Tag0            =   "9"
      Visible0        =   0   'False
      ToolTipText0    =   "Salir"
      UserControlWidth=   1200
   End
End
Attribute VB_Name = "frmOrdenRentaFijaCortoPlazo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Ordenes de Instrumentos de Renta Fija Corto Plazo"
Option Explicit

Dim arrFondo()              As String, arrFondoOrden()              As String
Dim arrTipoInstrumento()    As String, arrTipoInstrumentoOrden()    As String
Dim arrEstado()             As String, arrTipoOrden()               As String
Dim arrOperacion()          As String, arrNegociacion()             As String
Dim arrEmisor()             As String, arrMoneda()                  As String
Dim arrObligado()           As String, arrGestor()                  As String
Dim arrBaseAnual()          As String, arrTipoTasa()                As String
Dim arrOrigen()             As String, arrClaseInstrumento()        As String
Dim arrTitulo()             As String, arrSubClaseInstrumento()     As String
Dim arrConceptoCosto()      As String, arrFiador()                 As String

Dim strCodFondo             As String, strCodFondoOrden             As String
Dim strCodTipoInstrumento   As String, strCodTipoInstrumentoOrden   As String
Dim strCodEstado            As String, strCodTipoOrden              As String
Dim strCodOperacion         As String, strCodNegociacion            As String
Dim strCodEmisor            As String, strCodMoneda                 As String
Dim strCodObligado          As String, strCodGestor                 As String
Dim strCodBaseAnual         As String, strCodTipoTasa               As String
Dim strCodOrigen            As String, strCodClaseInstrumento       As String
Dim strCodTitulo            As String, strCodSubClaseInstrumento    As String
Dim strCodConcepto          As String, strCodReportado              As String
Dim strCodGarantia          As String, strCodAgente                 As String
Dim strEstado               As String, strSQL                       As String
Dim strCodFiador           As String, strIndGarantia               As String

Dim strCodFile              As String, strCodAnalitica              As String
Dim strCodGrupo             As String, strCodCiiu                   As String
Dim strEstadoOrden          As String, strCodCategoria              As String
Dim strCodRiesgo            As String, strCodSubRiesgo              As String
Dim strCalcVcto             As String, strCodSector                 As String
Dim strCodTipoCostoBolsa    As String, strCodTipoCostoConasev       As String
Dim strCodTipoCostoFondo    As String, strCodTipoCavali             As String
Dim strIndCuponCero         As String, strIndPacto                  As String
Dim strIndNegociable        As String, strCodigosFile               As String
Dim strCodIndiceInicial     As String, strCodIndiceFinal            As String
Dim strCodTipoAjuste        As String, strCodPeriodoPago            As String
Dim dblTipoCambio           As Double, dblTasaCuponNormal           As Double
Dim dblComisionBolsa        As Double, dblComisionConasev           As Double
Dim dblComisionFondo        As Double, dblComisionCavali            As Double
Dim intBaseCalculo          As Integer, dblFactorDiarioNormal       As Double

Dim SwCalculo               As Boolean


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
        With tabRFCortoPlazo
            .TabEnabled(0) = False
            'ACR:01/06/2009
            '.TabEnabled(2) = False
            'ACR:01/06/2009
            .Tab = 1
        End With
        'Call Habilita
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

    Dim dblTir As Double
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

    If DateDiff("d", dtpFechaEmision, dtpFechaVencimiento) < 0 Then
        MsgBox "La Fecha de vencimiento debe ser posterior a la Fecha de Emisión.", vbCritical, Me.Caption
        lblMontoVencimiento.Caption = "0"
    Else
    
        Dim intNumDias30    As Integer
        
        '*** Hallar los días 30/360,30/365 ***
        intNumDias30 = Dias360(dtpFechaEmision.Value, dtpFechaVencimiento.Value, True)
        
        If Not IsNumeric(txtDiasPlazo.Text) Then txtDiasPlazo.Text = "0"
        lblMontoVencimiento.Caption = CStr(ValorVencimiento(CCur(txtCantidad.Text), CDbl(txtTasa.Text), intBaseCalculo, CInt(txtDiasPlazo.Text), intNumDias30, strCodTipoTasa, strCodBaseAnual))

'Inicio ACR: 29/06/2009
'        If strCalcVcto = "D" Then
'            If Not IsNumeric(txtDiasPlazo.Text) Then txtDiasPlazo.Text = "0"
'            lblMontoVencimiento.Caption = CCur(txtCantidad.Text)
'        Else
'            Dim intNumDias30    As Integer
'
'            '*** Hallar los días 30/360,30/365 ***
'            intNumDias30 = Dias360(dtpFechaEmision.Value, dtpFechaVencimiento.Value, True)
'
'            If Not IsNumeric(txtDiasPlazo.Text) Then txtDiasPlazo.Text = "0"
'            lblMontoVencimiento.Caption = CStr(ValorVencimiento(CCur(txtCantidad.Text), CDbl(txtTasa.Text), intBaseCalculo, CInt(txtDiasPlazo.Text), intNumDias30, strCodTipoTasa, strCodBaseAnual))
'        End If
'Fin ACR: 29/06/2009

    End If
    lblVencimientoResumen.Caption = lblMontoVencimiento.Caption

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
    
    If DateDiff("d", dtpFechaEmision, dtpFechaVencimiento) < 0 Then
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
'            txtPrecioUnitario(0).Text = CStr(ValorActual(CCur(lblMontoVencimiento.Caption), CDbl(txtTasa.Text), intBaseCalculo, CInt(txtDiasPlazo.Text)) / CCur(lblMontoVencimiento.Caption) * 100)
        Else
            If Not IsNumeric(txtDiasPlazo.Text) Then txtDiasPlazo.Text = "0"
            txtPrecioUnitario(0).Text = "100"
        End If
    End If

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

    Dim adoRecord   As ADODB.Recordset
    Dim strSQL      As String
    Dim intRegistro As Integer
    
    Select Case strModo
        Case Reg_Adicion
        
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
            
            cboObligado.ListIndex = -1
            If cboObligado.ListCount > 0 Then cboObligado.ListIndex = 0
            
            cboGestor.ListIndex = -1
            If cboGestor.ListCount > 0 Then cboGestor.ListIndex = 0
            
            chkGarantia.Value = vbUnchecked
                        
            intRegistro = ObtenerItemLista(arrOrigen(), Codigo_Negociacion_Local)
            If intRegistro >= 0 Then cboOrigen.ListIndex = intRegistro
            
            dtpFechaOrden.Value = gdatFechaActual
            dtpFechaLiquidacion.Value = dtpFechaOrden.Value
            lblFechaLiquidacion.Caption = CStr(dtpFechaOrden.Value)
            
            intRegistro = ObtenerItemLista(arrMoneda(), strCodMoneda)
            If intRegistro >= 0 Then cboMoneda.ListIndex = intRegistro
            
            txtDiasPlazo.Text = "0"
            lblDiasPlazo.Caption = "0"
            
            txtInteresCorrido(0).Text = "0"
            
            txtTasa.Text = "0"
            
            txtDescripOrden.Text = Valor_Caracter
            txtNemonico.Text = Valor_Caracter
            txtObservacion.Text = Valor_Caracter
            txtPrecioUnitario(0).Text = "100"
            txtPrecioUnitario(1).Text = "0"
            txtValorNominal.Text = "1"
            txtCantidad.Text = "0"


            lblAnalitica.Caption = "??? - ????????"
            lblStockNominal.Caption = "0"
            lblClasificacion.Caption = Valor_Caracter

            dtpFechaEmision.Value = gdatFechaActual
            dtpFechaVencimiento.Value = dtpFechaEmision.Value
            dtpFechaPago.Value = dtpFechaVencimiento.Value
            lblFechaEmision.Caption = CStr(dtpFechaEmision.Value)
            lblFechaVencimiento.Caption = CStr(dtpFechaVencimiento.Value)
            
            
            cboBaseAnual.ListIndex = -1
            If cboBaseAnual.ListCount > 0 Then cboBaseAnual.ListIndex = 0
            
            cboTipoTasa.ListIndex = -1
            If cboTipoTasa.ListCount > 0 Then cboTipoTasa.ListIndex = 0
                                    
            chkAplicar(0).Value = vbUnchecked
            chkAplicar(1).Value = vbUnchecked
            
            lblSubTotal(0).Caption = "0"
            lblSubTotal(1).Caption = "0"
            
            Call IniciarComisiones
            
            txtInteresCorrido(0).Text = "0"
            txtInteresCorrido(1).Text = "0"
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
            
            cboFondoOrden.SetFocus
            
            txtMontoVencimiento1.Text = "0"
            txtTirBruta1.Text = "0"
            
            SwCalculo = False 'prepara la interfase para el ingreso de datos

            txtTirBruta1.Tag = 0         'indica cambio directo en la pantalla
            txtPrecioUnitario(0).Tag = 0
            txtMontoVencimiento1.Tag = 0

    End Select
    
End Sub


Public Sub Cancelar()

    cmdOpcion.Visible = True
    With tabRFCortoPlazo
        .TabEnabled(0) = True
        .Tab = 0
    End With
    Call Buscar
    
End Sub

Public Sub Eliminar()

    If strEstado = Reg_Consulta Or strEstado = Reg_Edicion Then
        Dim strMensaje  As String
        
        strMensaje = "Se procederá a eliminar la ORDEN " & tdgConsulta.Columns(1) & " por la " & _
            tdgConsulta.Columns(3) & vbNewLine & vbNewLine & vbNewLine & "¿ Seguro de continuar ?"
        
        If MsgBox(strMensaje, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
    
            '*** Anular Orden ***
            adoComm.CommandText = "UPDATE InversionOrden SET EstadoOrden='" & Estado_Orden_Anulada & "' WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodTitulo='" & Trim(tdgConsulta.Columns(2)) & "' AND NumOrden='" & Trim(tdgConsulta.Columns(1)) & "'"
                
            adoConn.Execute adoComm.CommandText
            
            '*** Anular Título si corresponde ***
            adoComm.CommandText = "UPDATE InstrumentoInversion SET IndVigente='' WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodTitulo='" & Trim(tdgConsulta.Columns(2)) & "'"
                
            adoConn.Execute adoComm.CommandText
            
            MsgBox Mensaje_Eliminacion_Exitosa, vbExclamation, Me.Caption
            
            tabRFCortoPlazo.TabEnabled(0) = True
            tabRFCortoPlazo.Tab = 0
            Call Buscar
            
            Exit Sub
        End If
    End If
    
End Sub

Public Sub Grabar()

    Dim adoRegistro         As ADODB.Recordset
    Dim strFechaOrden       As String, strFechaLiquidacion      As String
    Dim strFechaEmision     As String, strFechaVencimiento      As String
    Dim strFechaPago        As String
    Dim strMensaje          As String, strIndTitulo             As String
    Dim intRegistro         As Integer, intAccion               As Integer
    Dim lngNumError         As Long
    Dim dblTasaInteres      As Double
    
    'On Error GoTo CtrlError
    
    If strEstado = Reg_Consulta Then Exit Sub
    
    If strEstado = Reg_Adicion Then
        If TodoOK() Then
            strEstadoOrden = Estado_Orden_Ingresada
            
            '*** Validación del Límite de Inversión con respecto al Activo ***
'            If strCodTipoOrden = Codigo_Orden_Compra Then
'                Me.MousePointer = vbHourglass
'                intRes = ValidLimActivo(strCodFile, strCodFondoOrden, Convertyyyymmdd(dtpFechaOrden.Value), strCodMoneda, CCur(lblMontoTotal.Caption), CDbl(txtTipoCambio.Text), strCodEmisor)
'                Me.MousePointer = vbDefault
'
'                If intRes = 0 Then
'                    strEstadoOrden = Estado_Orden_PorAutorizar
'                End If
'            End If

            '*** Validación del Límite de Línea de Crédito Vigente ***
'            If strCodTipoOrden = Codigo_Orden_Compra Then
'                Me.MousePointer = vbHourglass
'                intRes = ValidLimCobertura(strCodEmisor, Convertyyyymmdd(dtpFechaOrden.Text), lblDescripMoneda.Tag, CCur(lblMontoTotal.Caption), CDbl(txtTipoCambio.Text))
'                Me.MousePointer = vbDefault
'
'                If intRes = 0 Then
'                    strEstadoOrden = Estado_Orden_PorAutorizar
'                End If
'            End If

        
            strMensaje = "Para proceder al Registro de la Orden Confirme lo siquiente : " & Chr(vbKeyReturn) & Chr(vbKeyReturn) & _
                "Fecha de Emisión" & Space(8) & ">" & Space(2) & CStr(dtpFechaEmision.Value) & Chr(vbKeyReturn) & _
                "Fecha de Vencimiento" & Space(1) & ">" & Space(2) & CStr(dtpFechaVencimiento.Value) & Chr(vbKeyReturn) & Chr(vbKeyReturn) & _
                "Fecha de Operación" & Space(4) & ">" & Space(2) & CStr(dtpFechaOrden.Value) & Chr(vbKeyReturn) & _
                "Fecha de Liquidación" & Space(3) & ">" & Space(2) & CStr(dtpFechaLiquidacion.Value) & Chr(vbKeyReturn) & Chr(vbKeyReturn) & _
                "Fecha de Pago" & Space(12) & ">" & Space(2) & CStr(dtpFechaPago.Value) & Chr(vbKeyReturn) & Chr(vbKeyReturn) & _
                "Nominal" & Space(24) & ">" & Space(2) & txtValorNominal.Text & Chr(vbKeyReturn) & _
                "Cantidad" & Space(22) & ">" & Space(2) & txtCantidad.Text & Chr(vbKeyReturn) & _
                "Precio Unitario (%)" & Space(6) & ">" & Space(2) & txtPrecioUnitario(0).Text & Chr(vbKeyReturn) & Chr(vbKeyReturn) & _
                "Monto Total" & Space(17) & ">" & Space(2) & Trim(lblDescripMoneda(0).Caption) & Space(1) & lblMontoTotal(0).Caption & Chr(vbKeyReturn) & Chr(vbKeyReturn) & _
                "Tir Neta" & Space(23) & ">" & Space(2) & lblTirNeta.Caption & Chr(vbKeyReturn) & _
                Chr(vbKeyReturn) & Chr(vbKeyReturn) & _
                "¿ Seguro de continuar ?"

            If MsgBox(strMensaje, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
               Me.Refresh: Exit Sub
            End If

        
            Me.MousePointer = vbHourglass
            
            strFechaOrden = Convertyyyymmdd(dtpFechaOrden.Value)
            strFechaLiquidacion = Convertyyyymmdd(dtpFechaLiquidacion.Value)
            strFechaEmision = Convertyyyymmdd(dtpFechaEmision.Value)
            strFechaVencimiento = Convertyyyymmdd(dtpFechaVencimiento.Value)
            strFechaPago = Convertyyyymmdd(dtpFechaPago.Value)
            
            Set adoRegistro = New ADODB.Recordset
            '*** Guardar Orden de Inversion ***
            With adoComm
                strIndTitulo = Valor_Caracter
                                
                If strCodTipoOrden = Codigo_Orden_Pacto Then
                    strIndTitulo = Valor_Caracter
                    strCodAnalitica = NumAleatorio(8)
                    strCodTitulo = NumAleatorio(15)
                    strCodTipoTasa = Codigo_Tipo_Tasa_Efectiva
                    strCodBaseAnual = Codigo_Base_Actual_365
                    strCodRiesgo = "00" ' Sin Clasificacion
                    strCodReportado = Valor_Caracter
                    strCodFile = Left(Trim(lblAnalitica.Caption), 3)
                ElseIf strCodTipoOrden = Codigo_Orden_Compra Then
                    If chkTitulo.Value Then
                        strIndTitulo = "X"
                        strCodTitulo = strCodGarantia
                    Else
                        strCodAnalitica = ObtenerNuevaAnalitica(strCodFile)
                        strCodTitulo = NumAleatorio(15)
                        
                    End If
                Else
                    strIndTitulo = Valor_Indicador
                    strCodTitulo = strCodGarantia
                    strCodGarantia = Valor_Caracter
                    strCodMoneda = lblMoneda.Tag
                    strFechaVencimiento = Convertyyyymmdd(Valor_Fecha)
                    strCodReportado = Valor_Caracter
                End If
                
                If strCalcVcto = "V" Then
                    dblTasaInteres = CDbl(txtTasa.Text)
                Else
                    dblTasaInteres = txtTirBruta1.Value
                End If
                                                                        
'                .CommandText = "BEGIN TRAN ProcOrden"
'                adoConn.Execute .CommandText
                
                .CommandText = "{ call up_IVAdicInversionOrden('" & strCodFondoOrden & "','" & _
                    gstrCodAdministradora & "','','" & strFechaOrden & "','" & _
                    strCodTitulo & "','" & Trim(txtNemonico.Text) & "','" & gstrPeriodoActual & "','" & gstrMesActual & "','" & _
                    "','" & strEstadoOrden & "','" & strCodAnalitica & "','" & strCodFile & "','" & _
                    strCodAnalitica & "','" & strCodClaseInstrumento & "','" & strCodSubClaseInstrumento & "','" & strCodTipoOrden & "','" & _
                    strCodOperacion & "','" & strCodNegociacion & "','" & strCodOrigen & "','" & Trim(txtDescripOrden.Text) & "','" & strCodEmisor & "','" & _
                    strCodAgente & "','" & strCodGarantia & "','" & strFechaPago & "','" & strFechaVencimiento & "','" & strFechaLiquidacion & "','" & _
                    strFechaEmision & "','" & strCodMoneda & "'," & CDec(txtCantidad.Text) & "," & CDec(txtTipoCambio.Text) & "," & _
                    txtValorNominal.Value & "," & txtPorcenDctoValorNominal.Value & "," & CDec(txtValorNominalDcto.Text) & "," & txtPrecioUnitario1.Value & "," & _
                    txtPrecioUnitario1.Value & "," & CDec(lblSubTotal(0).Caption) & "," & _
                    CDec(txtInteresCorrido(0).Text) & "," & CDec(txtComisionAgente(0).Text) & "," & CDec(txtComisionCavali(0).Text) & "," & _
                    CDec(txtComisionConasev(0).Text) & "," & CDec(txtComisionBolsa(0).Text) & "," & CDec(txtComisionFondo(0).Text) & ",0,0,0," & _
                    CDec(lblComisionIgv(0).Caption) & "," & CDec(lblMontoTotal(0).Caption) & "," & CDec(txtPrecioUnitario(1).Text) & "," & CDec(txtPrecioUnitario(1).Text) & "," & _
                    CDec(lblSubTotal(1).Caption) & "," & CDec(txtInteresCorrido(1).Text) & "," & CDec(txtComisionAgente(1).Text) & "," & _
                    CDec(txtComisionCavali(1).Text) & "," & CDec(txtComisionConasev(1).Text) & "," & CDec(txtComisionBolsa(1).Text) & "," & _
                    CDec(txtComisionFondo(1).Text) & ",0,0,0," & CDec(lblComisionIgv(1).Caption) & "," & CDec(lblMontoTotal(1).Caption) & "," & _
                    txtMontoVencimiento1.Value & "," & CInt(txtDiasPlazo.Text) & ",'','','','','','" & strCodReportado & "','" & strCodEmisor & "','','" & strCodObligado & "','" & strCodGestor & "','" & strCodFiador & "',0,'','','" & strIndTitulo & "','" & _
                    strCodTipoTasa & "','" & strCodBaseAnual & "'," & dblTasaInteres & "," & CDec(lblTirBrutaResumen.Caption) & "," & CDec(lblTirBrutaResumen.Caption) & "," & CDec(lblTirNetaResumen.Caption) & ",'" & _
                    strCodRiesgo & "','" & strCodSubRiesgo & "','" & Trim(txtObservacion.Text) & "') }"
                adoConn.Execute .CommandText

                
'                .CommandText = "COMMIT TRAN ProcOrden"
'                adoConn.Execute .CommandText
                                                                                                      
            End With
                                                                                    
            Me.MousePointer = vbDefault
        
            MsgBox Mensaje_Adicion_Exitosa, vbExclamation
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
            cmdOpcion.Visible = True
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
'    adoComm.CommandText = "ROLLBACK TRAN ProcOrden"
'    adoConn.Execute adoComm.CommandText
        
End Sub

Public Sub Modificar()

End Sub

Private Function TodoOK() As Boolean
        
    Dim adoRegistro         As ADODB.Recordset
    Dim strFechaDesde       As String, strFechaHasta        As String
    
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
        
        '*** Buscar en Títulos ***
        adoComm.CommandText = "SELECT Nemotecnico FROM InstrumentoInversion " & _
            "WHERE CodFile='" & strCodFile & "' AND Nemotecnico='" & Trim(txtNemonico.Text) & "' AND IndVigente='X'"
        Set adoRegistro = adoComm.Execute
        
        If Not adoRegistro.EOF Then
            MsgBox "Nemónico YA EXISTE...por favor verificar.", vbCritical, Me.Caption
            If txtNemonico.Enabled Then txtNemonico.SetFocus
            adoRegistro.Close: Set adoRegistro = Nothing
            Exit Function
        End If
        adoRegistro.Close
        
        strFechaDesde = Convertyyyymmdd(dtpFechaOrden.Value)
        strFechaHasta = Convertyyyymmdd(DateAdd("d", 1, dtpFechaOrden.Value))
        
        '*** Buscar en Ordenes del día ***
        adoComm.CommandText = "SELECT Nemotecnico FROM InversionOrden " & _
            "WHERE (FechaOrden>='" & strFechaDesde & "' AND FechaOrden<'" & strFechaHasta & "') AND " & _
            "CodFile='" & strCodFile & "' AND Nemotecnico='" & Trim(txtNemonico.Text) & "' AND EstadoOrden<>'" & Estado_Orden_Anulada & "'"
        Set adoRegistro = adoComm.Execute
        
        If Not adoRegistro.EOF Then
            MsgBox "Nemónico YA EXISTE...por favor verificar.", vbCritical, Me.Caption
            If txtNemonico.Enabled Then txtNemonico.SetFocus
            adoRegistro.Close: Set adoRegistro = Nothing
            Exit Function
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End If
        
    If Trim(txtDescripOrden.Text) = Valor_Caracter Then
        MsgBox "Debe indicar la Descripción de la ORDEN.", vbCritical, Me.Caption
        If txtDescripOrden.Enabled Then txtDescripOrden.SetFocus
        Exit Function
    End If
        
    If CVDate(dtpFechaEmision.Value) > CVDate(dtpFechaVencimiento.Value) Then
        MsgBox "La Fecha de Vencimiento debe ser mayor a la Fecha de Emisión.", vbCritical, Me.Caption
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
    
    If CDbl(txtTipoCambio.Text) = 0 Then
        MsgBox "Debe indicar el Tipo de Cambio.", vbCritical, Me.Caption
        If txtTipoCambio.Enabled Then txtTipoCambio.SetFocus
        Exit Function
    End If
    
    '*** Validación de STOCK ***
    If strCodTipoOrden = Codigo_Orden_Venta Then
        If CCur(txtValorNominal.Text) > CCur(lblStockNominal.Caption) Then
            MsgBox "Stock insuficiente para Registrar la Orden de Venta.", vbCritical, Me.Caption
            If txtValorNominal.Enabled Then txtValorNominal.SetFocus
            Exit Function
        End If
'    Else
'        If CCur(lblMontoVencimiento.Caption) = 0 Then
'            MsgBox "Debe calcular el Valor al Vencimiento.", vbCritical, Me.Caption
'            If cmdCalculo.Enabled Then cmdCalculo.SetFocus
'            Exit Function
'        End If
    End If
    
    If txtTirBruta1.Value = 0 Then
        MsgBox "Debe calcular la Tir Bruta.", vbCritical, Me.Caption
        If cmdCalculo.Enabled Then cmdCalculo.SetFocus
        Exit Function
    End If
    
'    If CDbl(lblTirNeta.Caption) = 0 Then
'        MsgBox "Debe calcular la Tir Neta.", vbCritical, Me.Caption
'        If cmdCalculo.Enabled Then cmdCalculo.SetFocus
'        Exit Function
'    End If
    
    '*** Si todo paso OK ***
    TodoOK = True
  
End Function
Public Sub Imprimir()

End Sub

Public Sub SubImprimir(Index As Integer)

    Dim frmReporte              As frmVisorReporte
    Dim aReportParamS(), aReportParamF(), aReportParamFn()
    Dim strFechaDesde           As String, strFechaHasta        As String
    Dim strSeleccionRegistro    As String

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
                            
                aReportParamF(0) = Trim(cboFondo.Text)
                aReportParamF(1) = gstrNombreEmpresa & Space(1)
                            
                aReportParamS(0) = strCodFondo
                aReportParamS(1) = gstrCodAdministradora
                aReportParamS(2) = Convertyyyymmdd(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10))
                aReportParamS(3) = Convertyyyymmdd(DateAdd("d", 1, CVDate(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10))))
                aReportParamS(4) = strCodMoneda
                aReportParamS(5) = strCodTipoInstrumento
            End If
        Case 3
            gstrNameRepo = "InversionCuponValorizacion"
            
'            strSeleccionRegistro = "{InversionOrden.FechaOrden} IN 'Fch1' TO 'Fch2'"
'            gstrSelFrml = strSeleccionRegistro
'            frmRangoFecha.Show vbModal
                        
            If gstrSelFrml <> "0" Then
                Set frmReporte = New frmVisorReporte
    
                ReDim aReportParamS(2)
                ReDim aReportParamFn(5)
                ReDim aReportParamF(5)
                            
                aReportParamFn(0) = "Usuario"
                aReportParamFn(1) = "FechaDesde"
                aReportParamFn(2) = "FechaHasta"
                aReportParamFn(3) = "Hora"
                aReportParamFn(4) = "Fondo"
                aReportParamFn(5) = "NombreEmpresa"
                            
                aReportParamF(0) = gstrLogin
                aReportParamF(1) = "" 'Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10)
                aReportParamF(2) = "" 'Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10)
                aReportParamF(3) = Format(Time(), "hh:mm:ss")
                aReportParamF(4) = Trim(cboFondo.Text)
                aReportParamF(5) = gstrNombreEmpresa & Space(1)
                            
                aReportParamS(0) = strCodFondo
                aReportParamS(1) = gstrCodAdministradora
                aReportParamS(2) = tdgConsulta.Columns("NumOrden") 'Convertyyyymmdd(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10))
                'aReportParamS(3) = Convertyyyymmdd(DateAdd("d", 1, CVDate(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10))))
                'aReportParamS(4) = strCodMoneda
                'aReportParamS(5) = strCodTipoInstrumento
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

Public Sub Salir()

    Unload Me
    
End Sub

Private Sub cboBaseAnual_Click()

    strCodBaseAnual = Valor_Caracter
    If cboBaseAnual.ListIndex < 0 Then Exit Sub
    
    strCodBaseAnual = Trim(arrBaseAnual(cboBaseAnual.ListIndex))
    
    '*** Base de Cálculo ***
    intBaseCalculo = 365
    Select Case strCodBaseAnual
        Case Codigo_Base_30_360: intBaseCalculo = 360
        Case Codigo_Base_Actual_365: intBaseCalculo = 365
        Case Codigo_Base_Actual_360: intBaseCalculo = 360
        Case Codigo_Base_30_365: intBaseCalculo = 365
    End Select
    
    txtValorNominal_Change
    
    'lblTirBruta.Caption = "0": lblTirNeta.Caption = "0"
    'lblMontoVencimiento.Caption = "0"
    
End Sub


Private Sub cboClaseInstrumento_Click()

    strCodClaseInstrumento = Valor_Caracter
    If cboClaseInstrumento.ListIndex < 0 Then Exit Sub
    
    strCodClaseInstrumento = Trim(arrClaseInstrumento(cboClaseInstrumento.ListIndex))
    
    If strCodClaseInstrumento = "001" Then strCalcVcto = "V"
    If strCodClaseInstrumento = "002" Then strCalcVcto = "D"
    
    '*** SubClase de Instrumento ***
    strSQL = "SELECT CodSubDetalleFile CODIGO,DescripSubDetalleFile DESCRIP FROM InversionSubDetalleFile WHERE " & _
        "CodDetalleFile='" & strCodClaseInstrumento & "' AND CodFile='" & strCodTipoInstrumentoOrden & "' AND IndVigente='X' ORDER BY DescripSubDetalleFile"
        
    CargarControlLista strSQL, cboSubClaseInstrumento, arrSubClaseInstrumento(), Sel_Defecto

    cboSubClaseInstrumento.ListIndex = 0
    cboSubClaseInstrumento.Enabled = True




End Sub


Private Sub cboConceptoCosto_Click()

    Dim adoRegistro As ADODB.Recordset
    
    strCodConcepto = Valor_Caracter
    If cboConceptoCosto.ListIndex < 0 Then Exit Sub
    
    strCodConcepto = Trim(arrConceptoCosto(cboConceptoCosto.ListIndex))
    
    strCodTipoCostoBolsa = Valor_Caracter: strCodTipoCostoConasev = Valor_Caracter
    strCodTipoCavali = Valor_Caracter: strCodTipoCostoFondo = Valor_Caracter
    dblComisionBolsa = 0: dblComisionConasev = 0
    dblComisionCavali = 0: dblComisionFondo = 0
        
    With adoComm
        Set adoRegistro = New ADODB.Recordset
                
        .CommandText = "SELECT CodCosto,TipoCosto,ValorCosto FROM CostoNegociacion WHERE TipoOperacion='" & strCodConcepto & "' AND TipoValor='" & Codigo_Valor_RentaFija & "' ORDER BY CodCosto"
        Set adoRegistro = .Execute

        Do Until adoRegistro.EOF
            Select Case Trim(adoRegistro("CodCosto"))
                Case Codigo_Costo_Bolsa
                    strCodTipoCostoBolsa = Trim(adoRegistro("TipoCosto"))
                    dblComisionBolsa = CDbl(adoRegistro("ValorCosto"))
                Case Codigo_Costo_Conasev
                    strCodTipoCostoConasev = Trim(adoRegistro("TipoCosto"))
                    dblComisionConasev = CDbl(adoRegistro("ValorCosto"))
                Case Codigo_Costo_Cavali
                    strCodTipoCavali = Trim(adoRegistro("TipoCosto"))
                    dblComisionCavali = CDbl(adoRegistro("ValorCosto"))
                Case Codigo_Costo_FLiquidacion
                    strCodTipoCostoFondo = Trim(adoRegistro("TipoCosto"))
                    dblComisionFondo = CDbl(adoRegistro("ValorCosto"))
           End Select
           adoRegistro.MoveNext
        Loop
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
End Sub


Private Sub cboEmisor_Click()

    Dim adoRegistro     As ADODB.Recordset
    
    strCodTitulo = Valor_Caracter: strCodGrupo = Valor_Caracter: strCodCiiu = Valor_Caracter
    strCodEmisor = Valor_Caracter: strCodAnalitica = Valor_Caracter
    lblAnalitica = strCodTipoInstrumentoOrden & "-??????": txtValorNominal.Text = "1"
    lblStockNominal = "0": strCodGrupo = Valor_Caracter
    
    If cboEmisor.ListIndex < 0 Then Exit Sub

    strCodEmisor = Left(Trim(arrEmisor(cboEmisor.ListIndex)), 8)
    strCodGrupo = Mid(Trim(arrEmisor(cboEmisor.ListIndex)), 9, 3)
    strCodCiiu = Right(Trim(arrEmisor(cboEmisor.ListIndex)), 4)

    '*** Validar Limites ***
    If strCodTipoInstrumentoOrden = Valor_Caracter Then Exit Sub
    If Not PosicionLimites() Then Exit Sub

    'txtDescripOrden = Trim(cboTipoInstrumentoOrden.Text) & " - " & Trim(cboEmisor.Text)
    strCodTitulo = strCodFondoOrden & strCodFile & strCodAnalitica
    
    With adoComm
        Set adoRegistro = New ADODB.Recordset
                        
        '*** Categoría del instrumento emitido por el emisor ***
        .CommandText = "SELECT CodCategoriaRiesgo,CodRiesgoFinal,CodSubRiesgoFinal FROM EmisionInstitucionPersona " & _
            "WHERE CodEmisor='" & strCodEmisor & "' AND CodFile='" & strCodTipoInstrumentoOrden & "' AND " & _
            "CodDetalleFile='" & strCodClaseInstrumento & "'"
            
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            strCodRiesgo = Trim(adoRegistro("CodRiesgoFinal"))
            strCodSubRiesgo = Trim(adoRegistro("CodSubRiesgoFinal"))
        Else
            If strCodEmisor <> Valor_Caracter Then
                MsgBox "La Clasificación de Riesgo no está definida...", vbCritical, Me.Caption
                Exit Sub
            End If
        End If
        adoRegistro.Close
        
        '*** Obtener el Riesgo ***
        .CommandText = "SELECT ValorParametro FROM AuxiliarParametro WHERE CodTipoParametro='TIPRIE' AND CodParametro='" & strCodRiesgo & "'"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            strCodCategoria = Trim(adoRegistro("ValorParametro"))
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
        
        lblClasificacion.Caption = strCodCategoria & Space(1) & strCodSubRiesgo
    End With
    
End Sub

Private Function PosicionLimites() As Boolean

    PosicionLimites = False
        
    If cboTipoInstrumentoOrden.ListIndex <= 0 Then
        MsgBox "Debe seleccionar el Tipo de Instrumento.", vbCritical, Me.Caption
        cboEmisor.ListIndex = -1: cboTitulo.ListIndex = -1
        If cboTipoInstrumentoOrden.Enabled Then cboTipoInstrumentoOrden.SetFocus
        Exit Function
    End If

'    If strCodTipoOrden = Codigo_Orden_Compra Then ValidLimites strCodEmisor, Convertyyyymmdd(dtpFechaOrden.Value), CDbl(txtTipoCambio.Text), strCodFile, strCodFondoOrden

    '*** Si todo pasó OK ***
    PosicionLimites = True
    
End Function

Private Sub cboEstado_Click()

    strCodEstado = Valor_Caracter
    If cboEstado.ListIndex < 0 Then Exit Sub
    
    strCodEstado = Trim(arrEstado(cboEstado.ListIndex))
    
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
            dtpFechaOrdenDesde.Value = gdatFechaActual
            dtpFechaOrdenHasta.Value = dtpFechaOrdenDesde.Value
            strCodMoneda = Trim(adoRegistro("CodMoneda"))
                        
            frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    '*** Tipo de Instrumento ***
    strSQL = "SELECT FIF.CodFile CODIGO,DescripFile DESCRIP " & _
        "FROM FondoInversionFile FIF JOIN InversionFile IVF ON(IVF.CodFile=FIF.CodFile) " & _
        "WHERE TipoValor='" & Codigo_Valor_RentaFija & "' AND TipoPlazo='" & Codigo_Valor_CortoPlazo & "' AND IndInstrumento='X' AND IndVigente='X' AND " & _
        "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' ORDER BY DescripFile"
    CargarControlLista strSQL, cboTipoInstrumento, arrTipoInstrumento(), Sel_Todos
    
    If cboTipoInstrumento.ListCount > 0 Then cboTipoInstrumento.ListIndex = 0
        
End Sub

Private Sub cboFondoOrden_Click()

    Dim adoRegistro As ADODB.Recordset
    
    strCodFondoOrden = Valor_Caracter
    If cboFondoOrden.ListIndex < 0 Then Exit Sub
    
    strCodFondoOrden = Trim(arrFondoOrden(cboFondoOrden.ListIndex))

    Set adoRegistro = New ADODB.Recordset
    
    With adoComm
        '*** Fecha Vigente, Moneda, Tipo de Cambio ***
        .CommandText = "{ call up_ACSelDatosParametro(23,'" & strCodFondoOrden & "','" & gstrCodAdministradora & "') }"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            gdatFechaActual = CVDate(adoRegistro("FechaCuota"))
            gstrFechaActual = Convertyyyymmdd(gdatFechaActual)
            dtpFechaOrden.Value = gdatFechaActual
            dtpFechaLiquidacion.Value = dtpFechaOrden.Value
            dtpFechaEmision.Value = dtpFechaOrden.Value
            dtpFechaVencimiento.Value = dtpFechaEmision.Value
            strCodMoneda = Trim(adoRegistro("CodMoneda"))
            txtTipoCambio.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, dtpFechaOrden.Value, Codigo_Moneda_Local, strCodMoneda))
            If CDbl(txtTipoCambio.Text) = 0 Then txtTipoCambio.Text = CStr(ObtenerTipoCambioMoneda(gstrCodClaseTipoCambioOperacionFondo, gstrValorTipoCambioOperacion, DateAdd("d", -1, dtpFechaOrden.Value), Codigo_Moneda_Local, strCodMoneda))
            dblTipoCambio = CDbl(txtTipoCambio.Text)
'            txtTipoCambio.Text = CStr(dblTipoCambio)
                        
            frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
            Me.Refresh
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    '*** Tipo de Instrumento ***
    strSQL = "SELECT FIF.CodFile CODIGO,DescripFile DESCRIP " & _
        "FROM FondoInversionFile FIF JOIN InversionFile IVF ON(IVF.CodFile=FIF.CodFile) " & _
        "WHERE TipoValor='" & Codigo_Valor_RentaFija & "' AND TipoPlazo='" & Codigo_Valor_CortoPlazo & "' AND IndInstrumento='X' AND IndVigente='X' AND " & _
        "CodFondo='" & strCodFondoOrden & "' AND CodAdministradora='" & gstrCodAdministradora & "' ORDER BY DescripFile"
    CargarControlLista strSQL, cboTipoInstrumentoOrden, arrTipoInstrumentoOrden(), Sel_Defecto
        
End Sub

Private Sub cboFiador_Click()

    strCodFiador = Valor_Caracter
    If cboFiador.ListIndex < 0 Then Exit Sub
    
    strCodFiador = Trim(arrFiador(cboFiador.ListIndex))

End Sub

Private Sub cboGestor_Click()
    
    strCodGestor = Valor_Caracter
    If cboGestor.ListIndex < 0 Then Exit Sub
    
    strCodGestor = Trim(arrGestor(cboGestor.ListIndex))

End Sub

Private Sub cboNegociacion_Click()

    strCodNegociacion = Valor_Caracter
    If cboNegociacion.ListIndex < 0 Then Exit Sub
    
    strCodNegociacion = Trim(arrNegociacion(cboNegociacion.ListIndex))
            
    cboConceptoCosto.ListIndex = -1
    If cboConceptoCosto.ListCount > 0 Then cboConceptoCosto.ListIndex = 0
    
    cboConceptoCosto.Enabled = False
    If strCodNegociacion = Codigo_Mecanismo_Rueda Then cboConceptoCosto.Enabled = True
     
End Sub

Private Sub cboObligado_Click()

    strCodObligado = Valor_Caracter
    If cboObligado.ListIndex < 0 Then Exit Sub
    
    strCodObligado = Trim(arrObligado(cboObligado.ListIndex))

End Sub

Private Sub cboSubClaseInstrumento_Click()

    strCodSubClaseInstrumento = Valor_Caracter
    If cboSubClaseInstrumento.ListIndex < 0 Then Exit Sub
    
    strCodSubClaseInstrumento = Trim(arrSubClaseInstrumento(cboSubClaseInstrumento.ListIndex))
    
    If strCodSubClaseInstrumento = "001" Or strCalcVcto = "V" Then 'Al vencimiento
        'txtPrecioUnitario(0).Enabled = False
        strCalcVcto = "V"
        txtCantidad.Enabled = True
        txtCantidad.Text = "1"
        cboTipoTasa.Enabled = True
        txtTasa.Enabled = True
        txtPorcenDctoValorNominal.Text = "100"
        txtPrecioUnitario1.Text = "100"
        txtPorcenDctoValorNominal.Enabled = False
        txtPrecioUnitario1.Enabled = False
    End If
    
    If strCodSubClaseInstrumento = "002" Or strCalcVcto = "D" Then 'Al descuento
        'txtPrecioUnitario(0).Enabled = True
        strCalcVcto = "D"
        txtCantidad.Text = "1"
        txtCantidad.Enabled = False
        cboTipoTasa.Enabled = False
        txtTasa.Enabled = False
        txtPorcenDctoValorNominal.Text = "100"
        txtPrecioUnitario1.Text = "100"
        txtPorcenDctoValorNominal.Enabled = True
        txtPrecioUnitario1.Enabled = True
    End If
    
End Sub

Private Sub cboOrigen_Click()

    strCodOrigen = Valor_Caracter
    If cboOrigen.ListIndex < 0 Then Exit Sub
    
    strCodOrigen = Trim(arrOrigen(cboOrigen.ListIndex))
    
End Sub

Private Sub cboTipoInstrumento_Click()

    strCodTipoInstrumento = Valor_Caracter
    If cboTipoInstrumento.ListIndex < 0 Then Exit Sub
    
    strCodTipoInstrumento = Trim(arrTipoInstrumento(cboTipoInstrumento.ListIndex))
    
End Sub

Private Sub cboTipoInstrumentoOrden_Click()
    
    Dim adoRegistro As ADODB.Recordset
    Dim strFecha    As String
    
    strCodTipoInstrumentoOrden = Valor_Caracter
    strIndPacto = Valor_Caracter: strIndNegociable = Valor_Caracter
    If cboTipoInstrumentoOrden.ListIndex < 0 Then Exit Sub
    
    strCodTipoInstrumentoOrden = Trim(arrTipoInstrumentoOrden(cboTipoInstrumentoOrden.ListIndex))

'    Set adoRegistro = New ADODB.Recordset
'    With adoComm
'        .CommandText = "SELECT IndPacto,IndNegociable FROM InversionFile WHERE CodFile='" & strCodTipoInstrumentoOrden & "'"
'        Set adoRegistro = .Execute
'
'        If Not adoRegistro.EOF Then
'            strIndPacto = Trim(adoRegistro("IndPacto"))
'            strIndNegociable = Trim(adoRegistro("IndNegociable"))
'        End If
'        adoRegistro.Close: Set adoRegistro = Nothing
'    End With

    Set adoRegistro = New ADODB.Recordset
    With adoComm
        .CommandText = "SELECT DescripFile,DescripInicial FROM InversionFile WHERE CodFile='" & strCodTipoInstrumentoOrden & "'"
        Set adoRegistro = .Execute
            
        If Not adoRegistro.EOF Then
            strFecha = Format(Day(gdatFechaActual), "00") & Format(Month(gdatFechaActual), "00") & Format(Year(gdatFechaActual), "0000")
            txtNemonico.Text = Trim(adoRegistro("DescripInicial")) & strFecha & GetTickCount
            txtDescripOrden = Trim(cboTipoInstrumentoOrden.Text) & " - " & Trim(txtNemonico.Text)
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    '*** Tipo de Orden ***
    strSQL = "SELECT IFTON.CodTipoOperacion CODIGO,DescripTipoOperacion DESCRIP " & _
        "FROM InversionFileTipoOperacionNegociacion IFTON JOIN TipoOperacionNegociacion TON ON(TON.CodTipoOperacion=IFTON.CodTipoOperacion)" & _
        "WHERE IFTON.CodFile='" & strCodTipoInstrumentoOrden & "' ORDER BY DescripTipoOperacion"
    CargarControlLista strSQL, cboTipoOrden, arrTipoOrden(), Valor_Caracter

    If cboTipoOrden.ListCount > 0 Then cboTipoOrden.ListIndex = 0
    
    lblAnalitica.Caption = strCodTipoInstrumentoOrden & " - ????????"
    strCodFile = strCodTipoInstrumentoOrden

    '*** Clase de Instrumento ***
    strSQL = "SELECT CodDetalleFile CODIGO,DescripDetalleFile DESCRIP FROM InversionDetalleFile WHERE CodFile='" & strCodTipoInstrumentoOrden & "' AND IndVigente='X' ORDER BY DescripDetalleFile"
    CargarControlLista strSQL, cboClaseInstrumento, arrClaseInstrumento(), Sel_Defecto
    
    If cboClaseInstrumento.ListCount > 0 Then
        cboClaseInstrumento.ListIndex = 0
        cboClaseInstrumento.Enabled = True
    End If
            
End Sub

Private Sub cboMoneda_Click()
    
    lblDescripMoneda(0).Caption = "S/.": lblDescripMoneda(0).Tag = Codigo_Moneda_Local
    lblDescripMoneda(1).Caption = "S/.": lblDescripMoneda(1).Tag = Codigo_Moneda_Local
    lblDescripMonedaResumen(0) = "S/.": lblDescripMonedaResumen(0).Tag = Codigo_Moneda_Local
    lblDescripMonedaResumen(1) = "S/.": lblDescripMonedaResumen(1).Tag = Codigo_Moneda_Local
    
    If cboMoneda.ListIndex < 0 Then Exit Sub
    
    strCodMoneda = Trim(arrMoneda(cboMoneda.ListIndex))
        
    lblDescripMoneda(0).Caption = ObtenerSignoMoneda(strCodMoneda)
    lblDescripMoneda(0).Tag = strCodMoneda
    lblDescripMoneda(1).Caption = ObtenerSignoMoneda(strCodMoneda)
    lblDescripMoneda(1).Tag = strCodMoneda
    lblDescripMonedaResumen(0).Caption = ObtenerSignoMoneda(strCodMoneda)
    lblDescripMonedaResumen(0).Tag = strCodMoneda
    lblDescripMonedaResumen(1).Caption = ObtenerSignoMoneda(strCodMoneda)
    lblDescripMonedaResumen(1).Tag = strCodMoneda
    lblMoneda.Caption = ObtenerDescripcionMoneda(strCodMoneda)
    
End Sub

Private Sub cboOperacion_Click()

    strCodOperacion = Valor_Caracter
    If cboOperacion.ListIndex < 0 Then Exit Sub
    
    strCodOperacion = Trim(arrOperacion(cboOperacion.ListIndex))
    
End Sub

Public Sub CargarComisiones(ByVal strCodComision As String, Index As Integer)
     
     Call AplicarCostos(Index)
     
End Sub

Private Sub cboTipoOrden_Click()

    strCodTipoOrden = Valor_Caracter
    If cboTipoOrden.ListIndex < 0 Then Exit Sub

    strCodTipoOrden = Trim(arrTipoOrden(cboTipoOrden.ListIndex))

    Me.MousePointer = vbHourglass
    Select Case strCodTipoOrden
        Case Codigo_Orden_Compra
            chkTitulo.Enabled = True
            cboTitulo.Visible = False: cboEmisor.Visible = True
            lblDescrip(6) = "Emisor"
            
            If chkTitulo.Value Then
                strSQL = "SELECT CodTitulo CODIGO,(Nemotecnico + ' ' + DescripTitulo)DESCRIP FROM InstrumentoInversion " & _
                    "WHERE CodFile='" & strCodFile & "' AND CodDetalleFile='" & strCodClaseInstrumento & "' AND IndVigente='X' ORDER BY DescripTitulo"
                CargarControlLista strSQL, cboTitulo, arrTitulo(), Sel_Defecto
            
                If cboTitulo.ListCount > 0 Then cboTitulo.ListIndex = 0
            End If

            fraComisionMontoFL2.Visible = False

        Case Codigo_Orden_Venta
            chkTitulo.Enabled = False
            cboTitulo.Visible = True: cboEmisor.Visible = False
            lblDescrip(6) = "Título"
            
            strSQL = "SELECT InstrumentoInversion.CodTitulo CODIGO," & _
                "(RTRIM(InstrumentoInversion.Nemotecnico) + ' ' + RTRIM(InstrumentoInversion.DescripTitulo)) DESCRIP FROM InstrumentoInversion,InversionKardex " & _
                "WHERE SaldoFinal > 0 AND IndUltimoMovimiento='X' AND InstrumentoInversion.CodFile=InversionKardex.CodFile AND " & _
                "InstrumentoInversion.CodAnalitica=InversionKardex.CodAnalitica AND InversionKardex.CodFile='" & strCodFile & "' AND " & _
                "InstrumentoInversion.CodFondo='" & strCodFondoOrden & "' AND InversionKardex.CodFondo='" & strCodFondoOrden & "' " & _
                "ORDER BY InstrumentoInversion.Nemotecnico"
            CargarControlLista strSQL, cboTitulo, arrTitulo(), Sel_Defecto
        
            If cboTitulo.ListCount > 0 Then cboTitulo.ListIndex = 0

            fraComisionMontoFL2.Visible = False
            
        Case Codigo_Orden_Pacto
            chkTitulo.Enabled = True
            cboTitulo.Visible = False: cboEmisor.Visible = True
            lblDescrip(6) = "Emisor"
        
            If chkTitulo.Value Then
                strSQL = "SELECT CodTitulo CODIGO,(Nemotecnico + ' ' + DescripTitulo)DESCRIP FROM InstrumentoInversion " & _
                    "WHERE CodFile='" & strCodFile & "' AND IndVigente='X' ORDER BY DescripTitulo"
                CargarControlLista strSQL, cboTitulo, arrTitulo(), Sel_Defecto
            
                If cboTitulo.ListCount > 0 Then cboTitulo.ListIndex = 0
            End If
            
            fraComisionMontoFL2.Visible = True
                            
    End Select
    Me.MousePointer = vbDefault
    
End Sub

Private Sub cboTipoTasa_Click()

    strCodTipoTasa = Valor_Caracter
    If cboTipoTasa.ListIndex < 0 Then Exit Sub
    
    strCodTipoTasa = Trim(arrTipoTasa(cboTipoTasa.ListIndex))
    
End Sub

Private Sub cboTitulo_Click()

    Dim adoRegistro     As ADODB.Recordset
    Dim intRegistro     As Integer
    
    strCodGarantia = Valor_Caracter: txtDescripOrden = Valor_Caracter: strCodAnalitica = Valor_Caracter
    lblAnalitica = strCodTipoInstrumentoOrden & "-????????": txtValorNominal.Text = "1"
    lblStockNominal = "0"
    strCodEmisor = Valor_Caracter: strCodGrupo = Valor_Caracter
    If cboTitulo.ListIndex < 0 Then Exit Sub

    strCodGarantia = Trim(arrTitulo(cboTitulo.ListIndex))

    With adoComm
        Set adoRegistro = New ADODB.Recordset

        .CommandText = "SELECT CodAnalitica,ValorNominal,CodMoneda,CodEmisor,CodGrupo,FechaEmision,FechaVencimiento," & _
            "TasaInteres,CodRiesgo,CodSubRiesgo,CodTipoTasa,BaseAnual,Nemotecnico " & _
            "FROM InstrumentoInversion WHERE CodTitulo='" & strCodGarantia & "'"
        Set adoRegistro = .Execute

        If Not adoRegistro.EOF Then
            strCodAnalitica = Trim(adoRegistro("CodAnalitica"))
            lblAnalitica.Caption = strCodFile & "-" & strCodAnalitica
                        
            intRegistro = ObtenerItemLista(arrMoneda(), adoRegistro("CodMoneda"))
            If intRegistro >= 0 Then cboMoneda.ListIndex = intRegistro
            
            dtpFechaEmision.Value = adoRegistro("FechaEmision")
            dtpFechaVencimiento.Value = adoRegistro("FechaVencimiento")
            dtpFechaVencimiento_Change
            txtNemonico.Text = Trim(adoRegistro("Nemotecnico"))
            
            intRegistro = ObtenerItemLista(arrTipoTasa(), adoRegistro("CodTipoTasa"))
            If intRegistro >= 0 Then cboTipoTasa.ListIndex = intRegistro
            
            intRegistro = ObtenerItemLista(arrBaseAnual(), adoRegistro("BaseAnual"))
            If intRegistro >= 0 Then cboBaseAnual.ListIndex = intRegistro
            
            txtTasa.Text = adoRegistro("TasaInteres")
            txtValorNominal.Text = CStr(adoRegistro("ValorNominal"))
                
            strCodEmisor = Trim(adoRegistro("CodEmisor")): strCodGrupo = Trim(adoRegistro("CodGrupo"))
            strCodRiesgo = Trim(adoRegistro("CodRiesgo"))
            strCodSubRiesgo = Trim(adoRegistro("CodSubRiesgo"))
            lblMoneda.Caption = ObtenerDescripcionMoneda(adoRegistro("CodMoneda"))
            lblBaseTasaCupon.Caption = "360" & Space(1) & "-" & Space(1) & Trim(txtTasa.Text) & "%"
            If adoRegistro("BaseAnual") = Codigo_Base_Actual_Actual Then lblBaseTasaCupon.Caption = "365" & Space(1) & "-" & Space(1) & Trim(txtTasa.Text) & "%"
            If adoRegistro("BaseAnual") = Codigo_Base_Actual_365 Then lblBaseTasaCupon.Caption = "365" & Space(1) & "-" & Space(1) & Trim(txtTasa.Text) & "%"
            If adoRegistro("BaseAnual") = Codigo_Base_30_365 Then lblBaseTasaCupon.Caption = "365" & Space(1) & "-" & Space(1) & Trim(txtTasa.Text) & "%"
            
            cboMoneda.Enabled = False
            cboTipoTasa.Enabled = False
            cboBaseAnual.Enabled = False
            dtpFechaVencimiento.Enabled = False
            txtDiasPlazo.Enabled = False
            txtValorNominal.Enabled = False
            txtTasa.Enabled = False
            txtNemonico.Enabled = False
        End If
        adoRegistro.Close

        .CommandText = "SELECT FechaPago " & _
            "FROM InstrumentoInversionCalendario WHERE CodTitulo='" & strCodGarantia & "'"
        Set adoRegistro = .Execute

        If Not adoRegistro.EOF Then
            dtpFechaPago.Value = adoRegistro("FechaPago")
            dtpFechaPago_Change
            dtpFechaPago.Enabled = False
        End If
        adoRegistro.Close
        
        '*** Obtener el Riesgo ***
        .CommandText = "SELECT ValorParametro FROM AuxiliarParametro WHERE CodTipoParametro='TIPRIE' AND CodParametro='" & strCodRiesgo & "'"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            strCodCategoria = Trim(adoRegistro("ValorParametro"))
        End If
        adoRegistro.Close
        
        lblClasificacion.Caption = strCodCategoria & Space(1) & strCodSubRiesgo
        
        '*** Validar Limites ***
        If Not PosicionLimites() Then Exit Sub

        .CommandText = "SELECT SaldoFinal,ValorPromedio FROM InversionKardex WHERE CodAnalitica='" & strCodAnalitica & "' AND " & _
            "CodFile='" & strCodFile & "' AND CodFondo='" & strCodFondoOrden & "' AND " & _
            "CodAdministradora='" & gstrCodAdministradora & "' AND IndUltimoMovimiento='X' AND SaldoFinal > 0"
            
        Set adoRegistro = .Execute

        If Not adoRegistro.EOF Then
            lblStockNominal.Caption = CStr(adoRegistro("SaldoFinal"))
        End If
        adoRegistro.Close: Set adoRegistro = Nothing

    End With

    txtDescripOrden = Trim(cboTipoInstrumentoOrden.Text) & " - " & Left(cboTitulo.Text, 15)
        
End Sub

Private Sub Check1_Click()

End Sub

Private Sub chkAplicar_Click(Index As Integer)

    If chkAplicar(Index).Value Then
        Call AplicarCostos(Index)
    Else
        Call IniciarComisiones
        Call CalculoTotal(Index)
    End If
    
End Sub

Private Sub chkFiador_Click()

    If chkFiador.Value = vbChecked Then
        cboFiador.Visible = True
        cboFiador.ListIndex = -1
    Else
        cboFiador.Visible = False
        cboFiador.ListIndex = -1
    End If

End Sub

Private Sub chkGarantia_Click()

    If chkGarantia.Value = vbChecked Then
        txtGarantia.Visible = True
        txtGarantia.Text = ""
    Else
        txtGarantia.Visible = False
        txtGarantia.Text = ""
    End If

End Sub

Private Sub chkTitulo_Click()

    If chkTitulo.Value Then
        cboTitulo.Visible = True: cboEmisor.Visible = False
        lblDescrip(6) = "Título"
        
        Me.MousePointer = vbHourglass
        Select Case strCodTipoOrden
            Case Codigo_Orden_Compra, Codigo_Orden_Pacto
'                strSQL = "SELECT CodTitulo CODIGO,(Nemotecnico + ' ' + DescripTitulo)DESCRIP FROM InstrumentoInversion " & _
'                    "WHERE CodFile='" & strCodFile & "' AND (CodFondo='" & strCodFondo & "' OR CodFondo='') AND " & _
'                    "(CodAdministradora='" & gstrCodAdministradora & "' OR CodAdministradora='') AND IndVigente='X' AND IndInversion='' " & _
'                    "ORDER BY DescripTitulo"
                strSQL = "SELECT CodTitulo CODIGO,(Nemotecnico + ' ' + DescripTitulo)DESCRIP FROM InstrumentoInversion " & _
                    "WHERE CodFile='" & strCodFile & "' AND CodDetalleFile='" & strCodClaseInstrumento & "' AND IndVigente='X' " & _
                    "ORDER BY DescripTitulo"
                CargarControlLista strSQL, cboTitulo, arrTitulo(), Sel_Defecto
                            
            Case Codigo_Orden_Venta
                strSQL = "SELECT InstrumentoInversion.CodTitulo CODIGO," & _
                    "(RTRIM(InstrumentoInversion.CodTitulo) + ' ' + RTRIM(InstrumentoInversion.Nemotecnico) + ' ' + RTRIM(InstrumentoInversion.DescripTitulo)) DESCRIP " & _
                    "FROM InstrumentoInversion,InversionKardex " & _
                    "WHERE SaldoFinal > 0 AND IndUltimo='X' AND InstrumentoInversion.CodFile=InversionKardex.CodFile AND " & _
                    "InstrumentoInversion.CodAnalitica=InversionKardex.CodAnalitica AND InversionKardex.CodFile='" & strCodFile & "' AND " & _
                    "(InstrumentoInversion.CodFondo='" & strCodFondoOrden & "' OR InstrumentoInversion.CodFondo='') AND " & _
                    "(InstrumentoInversion.CodAdministradora='" & gstrCodAdministradora & "' OR InstrumentoInversion.CodAdministradora='') AND " & _
                    "InversionKardex.CodFondo='" & strCodFondoOrden & "' " & _
                    "ORDER BY InstrumentoInversion.Nemotecnico"
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
    End If
        
End Sub

Private Sub cmdCalculo_Click()

    'Call CalcularValorVencimiento
    'Call CalcularPrecio
    Call CalcularTirBruta
    'Call CalcularTirNeta
    
End Sub

Private Sub cmdEnviar_Click()

    Dim strFechaDesde       As String, strFechaHasta        As String
    Dim intRegistro         As Integer, intContador         As Integer
    Dim datFecha            As Date
    
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
            adoComm.CommandText = "UPDATE InversionOrden SET EstadoOrden='" & Estado_Orden_Enviada & "'," & _
                "UsuarioEdicion='" & gstrLogin & "',FechaEdicion='" & strFechaDesde & Space(1) & Format(Time, "hh:mm") & "' " & _
                "WHERE NumOrden='" & Trim(tdgConsulta.Columns(1)) & "' AND CodFondo='" & strCodFondo & "' AND " & _
                "CodAdministradora='" & gstrCodAdministradora & "' AND EstadoOrden='" & Estado_Orden_Ingresada & "'"
        ElseIf strCodEstado = Estado_Orden_Enviada Then
            adoComm.CommandText = "UPDATE InversionOrden SET EstadoOrden='" & Estado_Orden_Ingresada & "'," & _
                "UsuarioEdicion='" & gstrLogin & "',FechaEdicion='" & strFechaDesde & Space(1) & Format(Time, "hh:mm") & "' " & _
                "WHERE NumOrden='" & Trim(tdgConsulta.Columns(1)) & "' AND CodFondo='" & strCodFondo & "' AND " & _
                "CodAdministradora='" & gstrCodAdministradora & "' AND EstadoOrden='" & Estado_Orden_Enviada & "'"
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

Private Sub dtpFechaEmision_Change()

    lblFechaEmision.Caption = CStr(dtpFechaEmision.Value)
    
    If cboTitulo.ListIndex > 0 Then
        lblDescrip(35).Caption = "Interés Corrido (" & DateDiff("d", dtpFechaEmision.Value, dtpFechaOrden.Value) & " días)"
        Call txtPrecioUnitario_Change(0) 'para un posible recalculo de intereses corridos (solo paa papeles que figuran en el maestro de titulos (pe. bonos)
    End If
    
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
    End If
    
End Sub

Private Sub dtpFechaVencimiento_Change()

    If dtpFechaVencimiento.Value < dtpFechaOrden.Value Then
        dtpFechaVencimiento.Value = dtpFechaOrden.Value
    End If
    
    If dtpFechaVencimiento.Value < dtpFechaEmision.Value Then
        dtpFechaVencimiento.Value = dtpFechaEmision.Value
    End If
    
    If dtpFechaVencimiento.Value < dtpFechaLiquidacion.Value Then
        dtpFechaVencimiento.Value = dtpFechaLiquidacion.Value
    End If
    
'    If Not EsDiaUtil(dtpFechaVencimiento.Value) Then
'        MsgBox "La Fecha no es un día útil...se cambiará por una fecha correcta !", vbInformation, Me.Caption
'        dtpFechaVencimiento.Value = ProximoDiaUtil(dtpFechaVencimiento.Value)
'    End If
    
    txtDiasPlazo.Text = CStr(DateDiff("d", dtpFechaLiquidacion.Value, dtpFechaVencimiento.Value))
    Call CalculoTotal(0)
    dtpFechaPago.Value = dtpFechaVencimiento.Value
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
    
    For intCont = 0 To (lblDescrip.Count - 1)
        Call FormatoEtiqueta(lblDescrip(intCont), vbLeftJustify)
    Next
            
End Sub
Public Sub Buscar()

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
    
    strSQL = "SELECT NumOrden,FechaOrden,FechaLiquidacion,CodTitulo,Nemotecnico,EstadoOrden,CodFile,CodAnalitica,TipoOrden,IOR.CodMoneda," & _
        "(RTRIM(DescripTipoOperacion) + SPACE(1) + DescripOrden) DescripOrden,CantOrden,ValorNominal,PrecioUnitarioMFL1,MontoTotalMFL1, CodSigno DescripMoneda " & _
        "FROM InversionOrden IOR JOIN TipoOperacionNegociacion TON ON(TON.CodTipoOperacion=IOR.TipoOrden) " & _
        "JOIN Moneda MON ON(MON.CodMoneda=IOR.CodMoneda) " & _
        "WHERE CodAdministradora='" & gstrCodAdministradora & "' AND CodFondo='" & strCodFondo & "' "
        
    If strCodTipoInstrumento <> Valor_Caracter Then
        strSQL = strSQL & "AND CodFile='" & strCodTipoInstrumento & "' "
    Else
        strSQL = strSQL & "AND CodFile IN " & strCodigosFile & " "
    End If

    If Not IsNull(dtpFechaOrdenDesde.Value) Or Not IsNull(dtpFechaOrdenHasta.Value) Then
        strSQL = strSQL & "AND (FechaOrden >='" & strFechaOrdenDesde & "' AND FechaOrden <'" & strFechaOrdenHasta & "') "
    End If
    
    If Not IsNull(dtpFechaLiquidacionDesde.Value) Or Not IsNull(dtpFechaLiquidacionHasta.Value) Then
        strSQL = strSQL & "AND (FechaLiquidacion >='" & strFechaLiquidacionDesde & "' AND FechaLiquidacion <'" & strFechaLiquidacionHasta & "') "
    End If
    
    If strCodEstado <> Valor_Caracter Then
        strSQL = strSQL & "AND EstadoOrden='" & strCodEstado & "' "
    End If
    
    strSQL = strSQL & "ORDER BY NumOrden"
    
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
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo3").Text = "Valorización Diaria"
    
    
End Sub
Private Sub CargarListas()

    Dim intRegistro As Integer
    
    '*** Fondos ***
    strSQL = "{ call up_ACSelDatosParametro(29,'" & gstrCodAdministradora & "') }"
    CargarControlLista strSQL, cboFondo, arrFondo(), Valor_Caracter
    CargarControlLista strSQL, cboFondoOrden, arrFondoOrden(), Valor_Caracter
        
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
            
    '*** Estados de la Orden ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='ESTORD' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboEstado, arrEstado(), Sel_Todos
    
    intRegistro = ObtenerItemLista(arrEstado(), Estado_Orden_Ingresada)
    If intRegistro >= 0 Then cboEstado.ListIndex = intRegistro
        
    '*** Tipo de Orden ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPORD' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboTipoOrden, arrTipoOrden(), Valor_Caracter
    
    '*** Emisor ***
    strSQL = "SELECT (CodPersona + CodGrupo + CodCiiu) CODIGO,DescripPersona DESCRIP FROM InstitucionPersona WHERE TipoPersona='" & Codigo_Tipo_Persona_Emisor & "' AND IndVigente='X' ORDER BY DescripPersona"
    CargarControlLista strSQL, cboEmisor, arrEmisor(), Sel_Defecto

    '*** Obligado ***
    strSQL = "SELECT (CodPersona) CODIGO,DescripPersona DESCRIP FROM InstitucionPersona WHERE TipoPersona='" & Codigo_Tipo_Persona_Emisor & "' AND IndVigente='X' ORDER BY DescripPersona"
    CargarControlLista strSQL, cboObligado, arrObligado(), Sel_Defecto

    '*** Gestor ***
    strSQL = "SELECT (CodPersona) CODIGO,DescripPersona DESCRIP FROM InstitucionPersona WHERE TipoPersona='" & Codigo_Tipo_Persona_Emisor & "' AND IndVigente='X' ORDER BY DescripPersona"
    CargarControlLista strSQL, cboGestor, arrGestor(), Sel_Defecto
    
    '*** Fiador ***
    strSQL = "SELECT (CodPersona) CODIGO,DescripPersona DESCRIP FROM InstitucionPersona WHERE TipoPersona='" & Codigo_Tipo_Persona_Emisor & "' AND IndVigente='X' ORDER BY DescripPersona"
    CargarControlLista strSQL, cboFiador, arrFiador(), Sel_Defecto

    '*** Mercado de Negociación ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='MDONEG' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboOrigen, arrOrigen(), Valor_Caracter
            
    '*** Moneda ***
    strSQL = "{ call up_ACSelDatos(2) }"
    CargarControlLista strSQL, cboMoneda, arrMoneda(), Sel_Defecto
    
    '*** Base de Cálculo ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='BASANU' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboBaseAnual, arrBaseAnual(), Valor_Caracter
    
    '*** Tipo Tasa ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='NATTAS' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboTipoTasa, arrTipoTasa(), ""
    
    '*** Tipo Liquidación Operación ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPLIQ' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboOperacion, arrOperacion(), Valor_Caracter
    
    '*** Mecanismos de Negociación ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='MECNEG' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboNegociacion, arrNegociacion(), Valor_Caracter

    '*** Conceptos de Costos de Negociación ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPCCO' AND ValorParametro='RF' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboConceptoCosto, arrConceptoCosto(), Sel_Defecto
    
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
            Call Imprimir
        Case vSave
            Call Grabar
        Case vCancel
            Call Cancelar
        Case vExit
            Call Salir
        
    End Select
    
End Sub

Private Sub InicializarValores()
    
    Dim adoRegistro As ADODB.Recordset
    
    '*** Valores Iniciales ***
    strEstado = Reg_Defecto
    tabRFCortoPlazo.Tab = 0

    SwCalculo = True 'indica cambio directo en la pantalla (false)

    dtpFechaOrdenDesde.Value = gdatFechaActual
    dtpFechaOrdenHasta.Value = dtpFechaOrdenDesde.Value
    dtpFechaLiquidacionDesde.Value = Null
    dtpFechaLiquidacionHasta.Value = dtpFechaLiquidacionDesde.Value
    
    lblPorcenIgv(0).Caption = CStr(gdblTasaIgv)
    lblPorcenIgv(1).Caption = CStr(gdblTasaIgv)
    
    Set adoRegistro = New ADODB.Recordset
    With adoComm
        .CommandText = "SELECT CodFile FROM InversionFile " & _
            "WHERE TipoValor='" & Codigo_Valor_RentaFija & "' AND TipoPlazo='" & Codigo_Valor_CortoPlazo & "' AND IndInstrumento='X' AND IndVigente='X' " & _
            "ORDER BY DescripFile"
        Set adoRegistro = .Execute
                
        strCodigosFile = Valor_Caracter
        Do While Not adoRegistro.EOF
            If strCodigosFile <> Valor_Caracter Then strCodigosFile = strCodigosFile & ",'"
            
            strCodigosFile = strCodigosFile & Trim(adoRegistro("CodFile")) & "'"
        
            adoRegistro.MoveNext
        Loop
        adoRegistro.Close: Set adoRegistro = Nothing
                
        strCodigosFile = "('" & strCodigosFile & ",'009')"
    End With
    
    '*** Ancho por defecto de las columnas de la grilla ***
    tdgConsulta.Columns(1).Width = tdgConsulta.Width * 0.01 * 8
    tdgConsulta.Columns(8).Width = tdgConsulta.Width * 0.01 * 12
    tdgConsulta.Columns(3).Width = tdgConsulta.Width * 0.01 * 32
    tdgConsulta.Columns(0).Width = tdgConsulta.Width * 0.01 * 6
    tdgConsulta.Columns(4).Width = tdgConsulta.Width * 0.01 * 13
    tdgConsulta.Columns(6).Width = tdgConsulta.Width * 0.01 * 11
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
                
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

Private Sub lblPorcenIgv_Change(Index As Integer)

    Call FormatoMillarEtiqueta(lblPorcenIgv(Index), Decimales_Monto)
    
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
    'If Not (CDbl(txtPorcenAgente(Index).Text) > 0 And CDbl(lblPorcenBolsa(Index).Caption) > 0 And CDbl(lblPorcenCavali(Index).Caption) > 0 And CDbl(lblPorcenFondo(Index).Caption) > 0 And CDbl(lblPorcenConasev(Index).Caption) > 0) Then Exit Sub
    
    'Calcula comisiones
    txtComisionAgente(Index).Text = CDbl((lblSubTotal(Index).Caption)) * CDbl(txtPorcenAgente(Index).Text) / 100
    txtComisionBolsa(Index).Text = CDbl((lblSubTotal(Index).Caption)) * CDbl(lblPorcenBolsa(Index).Caption) / 100
    txtComisionCavali(Index).Text = CDbl((lblSubTotal(Index).Caption)) * CDbl(lblPorcenCavali(Index).Caption) / 100
    txtComisionFondo(Index).Text = CDbl((lblSubTotal(Index).Caption)) * CDbl(lblPorcenFondo(Index).Caption) / 100
    txtComisionConasev(Index).Text = CDbl((lblSubTotal(Index).Caption)) * CDbl(lblPorcenConasev(Index).Caption) / 100
    
    
    If Not IsNumeric(txtTasa.Text) Or Not IsNumeric(txtCantidad.Text) Then Exit Sub
    
    'Calcula interes corrido
    If strCalcVcto = "V" Then
        txtInteresCorrido(Index).Text = CStr(CalculoInteresCorrido(strCodGarantia, CCur(txtValorNominalDcto.Text) * CCur(txtCantidad.Text), dtpFechaEmision.Value, dtpFechaOrden.Value, strCodIndiceFinal, strCodTipoAjuste, strCodTipoTasa, strCodPeriodoPago, strCodIndiceInicial, strCodBaseAnual, intBaseCalculo))
    Else
        '*** Calculando factores ***
        'ACR: Inicio Comentarios temporales: 04/06/2009
        'dblTasaCuponNormal = FactorAnualNormal(CDbl(txtTasa.Text), CInt(txtDiasPlazo.Text), intBaseCalculo, strCodTipoTasa, Valor_Indicador, Valor_Caracter, 0, CInt(txtDiasPlazo.Text), 1)
        'dblFactorDiarioNormal = FactorDiarioNormal(dblTasaCuponNormal, CInt(txtDiasPlazo.Text), strCodTipoTasa, Valor_Indicador, CInt(txtDiasPlazo.Text))
        'txtInteresCorrido(Index).Text = CStr(CalculoInteresCorrido(strCodGarantia, CCur(txtCantidad.Text), dtpFechaEmision.Value, dtpFechaOrden.Value, strCodIndiceFinal, strCodTipoAjuste, strCodTipoTasa, strCodPeriodoPago, strCodIndiceInicial, strCodBaseAnual, intBaseCalculo, dblFactorDiarioNormal))
        'ACR: Final Comentarios temporales: 04/06/2009
        txtInteresCorrido(Index).Text = "0"
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

    Select Case tabRFCortoPlazo.Tab
        Case 1, 2
            cmdAccion.Visible = True
            If PreviousTab = 0 And strEstado = Reg_Consulta Then tabRFCortoPlazo.Tab = 0
            If strEstado = Reg_Defecto Then tabRFCortoPlazo.Tab = 0
            If tabRFCortoPlazo.Tab = 2 Then
                fraDatosNegociacion.Caption = "Negociación" & Space(1) & "-" & Space(1) & _
                    Trim(cboTipoOrden.Text) & Space(1) & Trim(Left(cboTitulo.Text, 15))
            End If
    
        Case 0
            cmdAccion.Visible = False
    End Select
    
End Sub

'Private Sub TAMTextBox1_Click()
'
'    If Not (CDbl(txtTirBruta.Text) > 0 And CInt(txtDiasPlazo.Text) > 0 And CCur(txtCantidad.Text) > 0 And CDbl(txtValorNominal.Text) > 0) Then Exit Sub
'
'    If txtTirBruta.Tag = "0" Then 'indica cambio directo en la pantalla
'        txtPrecioUnitario(0).Tag = "1"
'        txtPrecioUnitario(0).Text = (CDbl(txtMontoVencimiento.Text) / ((1 + 0.01 * CDbl(txtTirBruta.Text)) ^ (CInt(txtDiasPlazo.Text) / 360)) * (CDbl(txtValorNominal.Text) * CDbl(txtCantidad.Text)) * 100)
'        'txtMontoVencimiento.Tag = "1"
'        'txtMontoVencimiento.Text = ((1 + 0.01 * CDbl(txtTirBruta.Text)) ^ (CInt(txtDiasPlazo.Text) / 360)) * (CDbl(txtValorNominal.Text) * CDbl(txtCantidad.Text) * CDbl(txtPrecioUnitario(0).Text) / 100) * 100
'    Else
'        txtTirBruta.Tag = "0"
'    End If
'
'End Sub

Private Sub tdgConsulta_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)

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
    
    If chkAplicar(Index).Value Then
        ActualizaPorcentaje txtComisionAgente(Index), txtPorcenAgente(Index)
    End If
    
End Sub

Private Sub txtComisionAgente_KeyPress(Index As Integer, KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtComisionAgente(Index), Decimales_Monto)
    
    If KeyAscii = vbKeyReturn Then
        If chkAplicar(Index).Value Then
            ActualizaPorcentaje txtComisionAgente(Index), txtPorcenAgente(Index)
        End If
        Call CalculoTotal(Index)
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

Private Sub txtDiasPlazo_Change()

    Call FormatoCajaTexto(txtDiasPlazo, 0)
    
    If IsNumeric(txtDiasPlazo.Text) Then
        dtpFechaVencimiento.Value = DateAdd("d", txtDiasPlazo.Text, CVDate(dtpFechaOrden.Value))
    Else
        dtpFechaVencimiento.Value = dtpFechaOrden.Value
    End If

    Call CalculoTotal(0)
    dtpFechaPago.Value = dtpFechaVencimiento.Value
    dtpFechaPago_Change
    lblDiasPlazo.Caption = CStr(txtDiasPlazo.Text)
    lblFechaVencimiento.Caption = CStr(dtpFechaVencimiento.Value)
    lblFechaCupon.Caption = CStr(dtpFechaVencimiento.Value)
    
    'ACR: 01/06/2009
    'If CInt(txtDiasPlazo.Text) > 0 Then tabRFCortoPlazo.TabEnabled(2) = True
    'ACR: 01/06/2009
    
End Sub

Private Sub CalculoTotal(Index As Integer)

    Dim curComImp As Currency, curMonTotal As Currency, curInteresCorrido As Currency

    If Not IsNumeric(txtComisionAgente(Index).Text) Or Not IsNumeric(txtComisionBolsa(Index).Text) Or Not IsNumeric(txtComisionConasev(Index).Text) Or Not IsNumeric(txtComisionCavali(Index).Text) Or Not IsNumeric(txtComisionFondo(Index).Text) Or Not IsNumeric(txtInteresCorrido(Index).Text) Then Exit Sub
    
    curComImp = CCur(CCur(txtComisionAgente(Index).Text) + CCur(txtComisionBolsa(Index).Text) + CCur(txtComisionConasev(Index).Text) + CCur(txtComisionCavali(Index).Text) + CCur(txtComisionFondo(Index).Text)) * CDbl(lblPorcenIgv(Index).Caption)
    lblComisionIgv(Index).Caption = CStr(curComImp)

    curComImp = CCur(txtComisionAgente(Index).Text) + CCur(txtComisionBolsa(Index).Text) + CCur(txtComisionConasev(Index).Text) + CCur(txtComisionCavali(Index).Text) + CCur(txtComisionFondo(Index).Text) + CCur(lblComisionIgv(Index).Caption)

    lblComisionesResumen(Index).Caption = CStr(curComImp)
            
    curInteresCorrido = CCur(txtInteresCorrido(Index).Text)
    
    If strCodTipoOrden = Codigo_Orden_Compra Or strCodTipoOrden = Codigo_Orden_Pacto Then  '*** Compra ***
        If Index = 0 Then
            curMonTotal = CCur(lblSubTotal(Index).Caption) + curComImp
        Else
            curMonTotal = CCur(lblSubTotal(Index).Caption) - curComImp
        End If
    ElseIf strCodTipoOrden = Codigo_Orden_Venta Then '*** Venta ***
        curMonTotal = CCur(lblSubTotal(Index).Caption) - curComImp
    End If
    
    curMonTotal = curMonTotal + curInteresCorrido   'CCur(txtInteresCorrido(Index).Text)

    lblMontoTotal(Index).Caption = CStr(curMonTotal)
    
'    txtMontoVencimiento1.Tag = "1"
'    txtMontoVencimiento1.Text = CStr(curMonTotal)
    
End Sub

Private Sub txtDiasPlazo_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "N")
    
End Sub

Private Sub txtDiasPlazo_LostFocus()

    cboEmisor_Click
    'ACR:01/06/2009
    'If CInt(txtDiasPlazo.Text) > 0 Then tabRFCortoPlazo.TabEnabled(2) = True
    'ACR:01/06/2009
    
End Sub

Private Sub AsignaComision(strTipoComision As String, dblValorComision As Double, ctrlValorComision As Control)
    
    If Not IsNumeric(lblSubTotal(ctrlValorComision.Index).Caption) Then Exit Sub
    
    If dblValorComision > 0 Then
        ctrlValorComision.Text = CStr(CCur(lblSubTotal(ctrlValorComision.Index)) * dblValorComision / 100)
    End If
            
End Sub
Private Sub dtpFechaEmision_LostFocus()

'    Dim intRes As Integer
'
'    If Not LEsDiaUtil(dtpFechaEmision) Then
'       dtpFechaEmision.Text = LProxDiaUtil(dtpFechaEmision.Text)
'    End If
'
'    If CVDate(dtpFechaVencimiento.Text) > CVDate(dtpFechaEmision.Text) Then
'       MsgBox "Fecha de Emisión debe ser anterior a la Fecha de Vencimiento", vbCritical
'       dtpFechaEmision.Text = dtpFechaVencimiento.Text
'       dtpFechaEmision.SetFocus
'    End If
    
End Sub

Private Sub dtpFechaLiquidacion_LostFocus()

'    Dim intRes As Integer
'
'    If Not LEsDiaUtil(dtpFechaLiquidacion) Then
'       dtpFechaLiquidacion.Text = LProxDiaUtil(dtpFechaLiquidacion.Text)
'    End If
'
'    If CVDate(dtpFechaOrden.Text) > CVDate(dtpFechaLiquidacion.Text) Then
'       MsgBox "Fecha de Liquidación debe ser posterior a la Fecha de Operación", vbCritical
'       dtpFechaLiquidacion.Text = dtpFechaOrden.Text
'       dtpFechaLiquidacion.SetFocus
'    End If
    
End Sub

Private Sub dtpFechaVencimiento_LostFocus()

    txtDiasPlazo.Text = CStr(DateDiff("d", dtpFechaOrden.Value, dtpFechaVencimiento.Value))
    
End Sub

Private Sub ActualizaPorcentaje(ctrlComision As Control, ctrlPorcentaje As Control)

    If Not IsNumeric(ctrlComision) Or Not IsNumeric(lblSubTotal(ctrlComision.Index).Caption) Then Exit Sub
                
    If CCur(lblSubTotal(ctrlComision.Index)) = 0 Then
        ctrlPorcentaje = "0"
    Else
        If CCur(ctrlComision) > 0 Then
            ctrlPorcentaje = CStr((CCur(ctrlComision) / CCur(lblSubTotal(ctrlComision.Index).Caption)) * 100)
        Else
            ctrlPorcentaje = "0"
        End If
    End If
                
End Sub

Private Sub txtInteresCorrido_Change(Index As Integer)

    Call FormatoCajaTexto(txtInteresCorrido(Index), Decimales_Monto)
    
    If Trim(txtInteresCorrido(Index).Text) <> Valor_Caracter Then
        lblInteresesResumen(Index).Caption = CStr(CCur(txtInteresCorrido(Index).Text))
    End If
    
End Sub

Private Sub txtInteresCorrido_KeyPress(Index As Integer, KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtInteresCorrido(Index), Decimales_Monto)
    
    If KeyAscii = vbKeyReturn Then
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



Private Sub txtMontoVencimiento1_Change()

     '-3 de julio -cumple de Jorge Sousa
    
    'If Not IsNumeric(txtMontoVencimiento1.Text) Or Not IsNumeric(txtDiasPlazo.Text) Or Not IsNumeric(txtCantidad.Text) Or Not IsNumeric(txtValorNominal.Text) Then Exit Sub
    'If Not (txtMontoVencimiento1.Value > 0 And CInt(txtDiasPlazo.Text) > 0 And CDbl(txtValorNominal.Text) > 0) Then Exit Sub

    'txtCantidad.Text = CStr(txtMontoVencimiento1.Value)

    'Call txtTirBruta1_Change
    
'    If txtMontoVencimiento1.Tag = "0" Then 'indica cambio directo en la pantalla
'        txtTirBruta1.Tag = "1"
'        txtTirBruta1.Text = ((CDbl(txtMontoVencimiento1.Text) / (CDbl(txtPrecioUnitario(0).Text) / 100 * CCur(txtCantidad.Text) * CDbl(txtValorNominal.Text))) ^ (360 / CInt(txtDiasPlazo.Text)) - 1) * 100
'    Else
'        txtMontoVencimiento1.Tag = "0"
'    End If

End Sub

Private Sub txtNemonico_Change()

    txtDescripOrden = Trim(cboTipoInstrumentoOrden.Text) & " - " & Trim(txtNemonico.Text)
    
End Sub

Private Sub txtNemonico_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
End Sub

Private Sub txtPorcenAgente_Change(Index As Integer)

    Call FormatoCajaTexto(txtPorcenAgente(Index), Decimales_Tasa)
    
End Sub

Private Sub txtPorcenAgente_KeyPress(Index As Integer, KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtPorcenAgente(Index), Decimales_Tasa)
    
    If KeyAscii = vbKeyReturn Then
        If chkAplicar(Index).Value Then
            ActualizaComision txtPorcenAgente(Index), txtComisionAgente(Index)
        End If
        Call CalculoTotal(Index)
    End If
    
End Sub

Private Sub txtPorcenDctoValorNominal_Change()

    Call txtValorNominal_Change

End Sub

Private Sub txtPrecioUnitario_Change(Index As Integer)

    Call FormatoCajaTexto(txtPrecioUnitario(Index), Decimales_Precio)

    If Not IsNumeric(txtCantidad.Text) Or Not IsNumeric(txtValorNominal.Text) Or Not IsNumeric(txtPrecioUnitario(Index).Text) Or Not IsNumeric(txtDiasPlazo.Text) Then Exit Sub
    If Not (CCur(txtCantidad.Text) > 0 And CDbl(txtValorNominal.Text) > 0 And CDbl(txtPrecioUnitario(Index).Text) > 0 And CInt(txtDiasPlazo.Text) > 0) Then Exit Sub

    lblSubTotal(Index).Caption = CDbl(txtValorNominal.Text) * CCur(txtCantidad.Text) * CDbl(txtPrecioUnitario(Index).Text) / 100

    'Aca calcula la TIR, si no es cambio directo
    If txtPrecioUnitario(Index).Tag = "0" Then
        txtTirBruta1.Tag = "1"
        txtTirBruta1.Text = ((CDbl(txtMontoVencimiento1.Text) / (CDbl(txtPrecioUnitario(0).Text) / 100 * CCur(txtCantidad.Text) * CDbl(txtValorNominal.Text))) ^ (360 / CInt(txtDiasPlazo.Text)) - 1) * 100
    Else
        txtPrecioUnitario(Index).Tag = "0"
    End If

    lblPrecioResumen(Index).Caption = CStr(txtPrecioUnitario(Index).Text)
    
End Sub

Private Sub txtPrecioUnitario_KeyPress(Index As Integer, KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtPrecioUnitario(Index), Decimales_Precio)
    
End Sub

Private Sub txtPrecioUnitario_LostFocus(Index As Integer)

    '    If strCalcVcto = "D" Then
'        ReDim Array_Monto(1): ReDim Array_Dias(1)
'        Array_Monto(0) = (CCur(lblSubTotal.Caption) + txtInteresCorrido.Text) * -1
'        Array_Dias(0) = dtpFechaLiquidacion.Text
'        Array_Monto(1) = Format(txtValorNominal.Text, "0.00")
'        Array_Dias(1) = dtpFechaVencimiento.Text
'        txtTasa.Text = Format(TIR(Array_Monto(), Array_Dias(), (10 / 100)) * 100, "0.0000")
'        txtTasa.Text = CDbl(Format(((1 + CDbl(txtTasa.Text) * 0.01) ^ (360 / 365) - 1) * 100, "0.0000"))
'    End If

End Sub

Private Sub txtPrecioUnitario1_Change()

    If Not IsNumeric(txtCantidad.Text) Or Not IsNumeric(txtDiasPlazo.Text) Or Not IsNumeric(txtValorNominalDcto.Text) Then Exit Sub
    
    If Not (CCur(txtCantidad.Text) > 0 And CDbl(txtValorNominalDcto.Text) > 0 And txtPrecioUnitario1.Value > 0 And CInt(txtDiasPlazo.Text) > 0) Then Exit Sub

    
    lblSubTotal(0).Caption = CDbl(txtValorNominalDcto.Text) * CCur(txtCantidad.Text) * txtPrecioUnitario1.Value / 100

    'Aca calcula la TIR, si no es cambio directo
    If txtPrecioUnitario1.Tag = "0" Then
        txtTirBruta1.Tag = "1"
        txtTirBruta1.Text = ((CDbl(txtValorNominalDcto.Text) / (txtPrecioUnitario1.Value / 100 * CCur(txtCantidad.Text) * CDbl(txtValorNominalDcto.Text))) ^ (intBaseCalculo / CInt(txtDiasPlazo.Text)) - 1) * 100
    Else
        txtPrecioUnitario1.Tag = "0"
    End If

    lblPrecioResumen(0).Caption = CStr(txtPrecioUnitario1.Value)
    

End Sub

Private Sub txtTasa_Change()

    Call FormatoCajaTexto(txtTasa, Decimales_Tasa)
    
    Call txtPrecioUnitario1_Change '(0)
    
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



'Private Sub txtTirBruta_KeyPress(KeyAscii As Integer)
'
'    Call ValidaCajaTexto(KeyAscii, "M", txtTirBruta, Decimales_Tasa)
'
'End Sub

Private Sub txtTirBruta1_Change()
    
    If Not (txtTirBruta1.Value <> 0 And CInt(txtDiasPlazo.Text) > 0 And CCur(txtCantidad.Text) > 0 And CDbl(txtValorNominalDcto.Text) > 0) Then Exit Sub
    
    If txtTirBruta1.Tag = "0" Then 'indica cambio directo en la pantalla
        txtPrecioUnitario1.Tag = "1"
        txtPrecioUnitario1.Text = (CDbl(txtValorNominalDcto.Text) / ((1 + 0.01 * txtTirBruta1.Value) ^ (CInt(txtDiasPlazo.Text) / intBaseCalculo)) / (CDbl(txtValorNominalDcto.Text) * CDbl(txtCantidad.Text)) * 100)
'        txtMontoVencimiento1.Tag = "1"
'        txtMontoVencimiento1.Text = ((1 + 0.01 * CDbl(txtTirBruta1.Text)) ^ (CInt(txtDiasPlazo.Text) / 360)) * (CDbl(txtValorNominal.Text) * CDbl(txtCantidad.Text) * CDbl(txtPrecioUnitario(0).Text) / 100)
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

'    Dim curCantidad     As Currency, curSubTotal            As Currency
'    Dim dblPreUni       As Double, dblFactorDiarioNormal    As Double
  
    If Not IsNumeric(txtCantidad.Text) Then Exit Sub
        
    txtValorNominalDcto.Text = CStr(txtPorcenDctoValorNominal.Value / 100 * txtValorNominal.Value)
        
    If strCalcVcto = "D" Then
        txtMontoVencimiento1.Text = txtValorNominal.Value * CCur(txtCantidad.Text)
    Else
        txtMontoVencimiento1.Text = txtValorNominal.Value * CCur(txtCantidad.Text) + CalculoInteresCorrido(strCodGarantia, txtValorNominal.Value * CCur(txtCantidad.Text), dtpFechaEmision.Value, dtpFechaVencimiento.Value, strCodIndiceFinal, strCodTipoAjuste, strCodTipoTasa, strCodPeriodoPago, strCodIndiceInicial, strCodBaseAnual, intBaseCalculo)
    End If
    
    Call txtPrecioUnitario1_Change '(0)
    
    
'Inicio ACR: 02-06-2009
'    If Trim(txtCantidad.Text) = Valor_Caracter Then Exit Sub
'
'    If IsNumeric(txtCantidad.Text) Then
'       curCantidad = CCur(txtCantidad.Text)
'    Else
'       curCantidad = 0
'    End If
'
'    lblCantidadResumen.Caption = CStr(curCantidad)
'
'    If IsNumeric(txtPrecioUnitario(0).Text) Then
'        dblPreUni = CDbl(txtPrecioUnitario(0).Text) * 0.01
'    End If
'
'    curSubTotal = curCantidad * dblPreUni
'    lblSubTotal(0).Caption = curSubTotal
'
'    Call CalculoTotal(0)
'
'    If CCur(txtCantidad.Text) > 0 And cboTitulo.ListIndex > 0 And strIndCuponCero = Valor_Caracter Then
'       txtInteresCorrido(0).Text = CStr(CalculoInteresCorrido(strCodGarantia, CDbl(curCantidad), dtpFechaEmision.Value, dtpFechaLiquidacion.Value, strCodIndiceFinal, strCodTipoAjuste, strCodTipoTasa, strCodPeriodoPago, strCodIndiceInicial, strCodBaseAnual, intBaseCalculo))
'    ElseIf CCur(txtCantidad.Text) > 0 And Trim(txtTasa.Text) <> Valor_Caracter And Trim(txtDiasPlazo.Text) <> Valor_Caracter And strIndCuponCero = Valor_Caracter Then
'        If CCur(txtCantidad.Text) > 0 And CDbl(txtTasa.Text) > 0 And CInt(txtDiasPlazo.Text) And strIndCuponCero = Valor_Caracter Then
'            '*** Calculando factores ***
'            dblTasaCuponNormal = FactorAnualNormal(CDbl(txtTasa.Text), CInt(txtDiasPlazo.Text), intBaseCalculo, strCodTipoTasa, Valor_Indicador, Valor_Caracter, 0, CInt(txtDiasPlazo.Text), 1)
'            dblFactorDiarioNormal = FactorDiarioNormal(dblTasaCuponNormal, CInt(txtDiasPlazo.Text), strCodTipoTasa, Valor_Indicador, CInt(txtDiasPlazo.Text))
'
'            txtInteresCorrido(0).Text = CStr(CalculoInteresCorrido(strCodGarantia, CDbl(curCantidad), dtpFechaEmision.Value, dtpFechaLiquidacion.Value, strCodIndiceFinal, strCodTipoAjuste, strCodTipoTasa, strCodPeriodoPago, strCodIndiceInicial, intBaseCalculo, dblFactorDiarioNormal))
'        End If
'    Else
'       txtInteresCorrido(0).Text = "0"
'    End If
'Fin ACR: 02-06-2009

    
End Sub

Private Sub txtValorNominal_KeyPress(KeyAscii As Integer)

    'Call ValidaCajaTexto(KeyAscii, "M", txtValorNominal, Decimales_Monto)
    
End Sub

Private Sub txtValorNominalDcto_Change()

    Call FormatoCajaTexto(txtValorNominalDcto, Decimales_Monto)

End Sub
