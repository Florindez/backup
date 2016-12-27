VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{5D1B2F4C-4B16-4B89-95C7-87E9AF4DB6BC}#1.0#0"; "TAMControls2.ocx"
Begin VB.Form frmCobertura 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Operación de Cobertura con Monedas"
   ClientHeight    =   9345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9345
   ScaleWidth      =   14505
   Begin TAMControls2.ucBotonEdicion2 cmdSalir 
      Height          =   735
      Left            =   12120
      TabIndex        =   101
      Top             =   8400
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
      Left            =   1080
      TabIndex        =   100
      Top             =   8400
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
   Begin MSAdodcLib.Adodc adoConsulta 
      Height          =   375
      Left            =   7200
      Top             =   8760
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin TabDlg.SSTab tabCobertura 
      Height          =   8055
      Left            =   120
      TabIndex        =   47
      Top             =   120
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   14208
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
      TabPicture(0)   =   "frmCobertura.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tdgConsulta"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraCriterio"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Datos Negociación"
      TabPicture(1)   =   "frmCobertura.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdAccion"
      Tab(1).Control(1)=   "txtObservacion"
      Tab(1).Control(2)=   "fraDatosCobertura"
      Tab(1).Control(3)=   "fraDatosActivo"
      Tab(1).Control(4)=   "fraDatosBasicos"
      Tab(1).Control(5)=   "lblDescrip(41)"
      Tab(1).ControlCount=   6
      Begin TAMControls2.ucBotonEdicion2 cmdAccion 
         Height          =   735
         Left            =   -64080
         TabIndex        =   102
         Top             =   6960
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
      Begin VB.TextBox txtObservacion 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   525
         Left            =   -73080
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   35
         Top             =   6840
         Width           =   7920
      End
      Begin VB.Frame fraDatosCobertura 
         Caption         =   "Datos Cobertura"
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
         Left            =   -74760
         TabIndex        =   77
         Top             =   3720
         Width           =   13575
         Begin VB.CheckBox chkCobertura 
            Caption         =   "Cobertura Parcial"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   22
            Top             =   2060
            Width           =   1575
         End
         Begin VB.ComboBox cboCuenta 
            Height          =   315
            Left            =   6240
            Style           =   2  'Dropdown List
            TabIndex        =   21
            ToolTipText     =   "cboCuenta.Text"
            Top             =   1695
            Width           =   1980
         End
         Begin VB.TextBox txtNemonico 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   2040
            MaxLength       =   15
            TabIndex        =   18
            Top             =   1359
            Width           =   1980
         End
         Begin VB.ComboBox cboMonedaCobertura 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   2500
            Width           =   1980
         End
         Begin VB.TextBox txtMontoCoberturado 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   6240
            MaxLength       =   45
            TabIndex        =   24
            Top             =   2140
            Width           =   1980
         End
         Begin VB.TextBox txtDiasPlazo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2040
            TabIndex        =   15
            Top             =   693
            Width           =   1720
         End
         Begin VB.TextBox txtDescripOrden 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   2040
            MaxLength       =   45
            TabIndex        =   17
            Top             =   1026
            Width           =   6160
         End
         Begin VB.TextBox txtTipoCambioFuturo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   11280
            MaxLength       =   45
            TabIndex        =   29
            Top             =   1026
            Width           =   1935
         End
         Begin VB.TextBox txtDiferencial 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   11280
            MaxLength       =   45
            TabIndex        =   28
            Top             =   693
            Width           =   1935
         End
         Begin VB.TextBox txtTipoCambioSpot 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   11280
            MaxLength       =   45
            TabIndex        =   27
            Top             =   360
            Width           =   1935
         End
         Begin VB.TextBox txtPorcenCobertura 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   2040
            MaxLength       =   45
            TabIndex        =   23
            Top             =   2140
            Width           =   1980
         End
         Begin VB.TextBox txtTipoCambio 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   6240
            MaxLength       =   45
            TabIndex        =   26
            Top             =   2500
            Width           =   1980
         End
         Begin VB.ComboBox cboIndicador 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   1695
            Width           =   1980
         End
         Begin VB.CommandButton cmdRentabilidad 
            Caption         =   "Rentabilidad"
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
            Left            =   8760
            TabIndex        =   32
            Top             =   2400
            Width           =   1335
         End
         Begin MSComCtl2.DTPicker dtpFechaOrden 
            Height          =   285
            Left            =   2040
            TabIndex        =   13
            Top             =   360
            Width           =   1980
            _ExtentX        =   3493
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
            Format          =   175833089
            CurrentDate     =   38776
         End
         Begin MSComCtl2.DTPicker dtpFechaLiquidacion 
            Height          =   285
            Left            =   6240
            TabIndex        =   14
            Top             =   360
            Width           =   1980
            _ExtentX        =   3493
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
            Format          =   175833089
            CurrentDate     =   38776
         End
         Begin MSComCtl2.DTPicker dtpFechaPago 
            Height          =   285
            Left            =   6240
            TabIndex        =   19
            Top             =   1359
            Width           =   1980
            _ExtentX        =   3493
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
            Format          =   175833089
            CurrentDate     =   38776
         End
         Begin MSComCtl2.DTPicker dtpFechaVencimiento 
            Height          =   285
            Left            =   6240
            TabIndex        =   16
            Top             =   693
            Width           =   1980
            _ExtentX        =   3493
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
            Format          =   175833089
            CurrentDate     =   38776
         End
         Begin MSComCtl2.UpDown updDiasPlazo 
            Height          =   285
            Left            =   3740
            TabIndex        =   93
            Top             =   693
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtDiasPlazo"
            BuddyDispid     =   196616
            OrigLeft        =   3315
            OrigTop         =   1090
            OrigRight       =   3570
            OrigBottom      =   1375
            Max             =   360
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Moneda de Cobertura"
            BeginProperty Font 
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
            Index           =   45
            Left            =   360
            TabIndex        =   99
            Top             =   2440
            Width           =   1365
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
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
            Left            =   4680
            TabIndex        =   98
            Top             =   1710
            Width           =   615
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
            Index           =   44
            Left            =   360
            TabIndex        =   97
            Top             =   1379
            Width           =   855
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Base 365"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
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
            Left            =   12120
            TabIndex        =   96
            Top             =   2145
            Width           =   810
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Base 360"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
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
            Left            =   10680
            TabIndex        =   95
            Top             =   2175
            Width           =   810
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
            Index           =   40
            Left            =   360
            TabIndex        =   94
            Top             =   1050
            Width           =   1020
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
            Index           =   15
            Left            =   360
            TabIndex        =   91
            Top             =   720
            Width           =   1095
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
            Index           =   38
            Left            =   4680
            TabIndex        =   90
            Top             =   1379
            Width           =   450
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Cambio"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
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
            Left            =   4680
            TabIndex        =   89
            Top             =   2520
            Width           =   1335
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
            Index           =   31
            Left            =   4680
            TabIndex        =   88
            Top             =   375
            Width           =   990
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Operación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
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
            Left            =   360
            TabIndex        =   87
            Top             =   380
            Width           =   885
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Coberturar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
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
            TabIndex        =   86
            Top             =   1710
            Width           =   900
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
            Index           =   32
            Left            =   4680
            TabIndex        =   85
            Top             =   720
            Width           =   1050
         End
         Begin VB.Label lblDescrip 
            Caption         =   "Monto Coberturado"
            BeginProperty Font 
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
            Index           =   30
            Left            =   4680
            TabIndex        =   84
            Top             =   2080
            Width           =   1395
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
            Index           =   1
            Left            =   4020
            TabIndex        =   83
            Top             =   2140
            Width           =   150
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Cambio Spot"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
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
            Left            =   8820
            TabIndex        =   82
            Top             =   375
            Width           =   1515
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Cambio Futuro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
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
            Left            =   8820
            TabIndex        =   81
            Top             =   1050
            Width           =   1665
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Diferencial"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
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
            Left            =   8820
            TabIndex        =   80
            Top             =   720
            Width           =   930
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
            Index           =   36
            Left            =   8820
            TabIndex        =   79
            Top             =   1380
            Width           =   720
         End
         Begin VB.Label lblMontoContado 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   11280
            TabIndex        =   30
            Tag             =   "0.00"
            Top             =   1359
            Width           =   1935
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Futuro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
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
            Left            =   8820
            TabIndex        =   78
            Top             =   1710
            Width           =   555
         End
         Begin VB.Label lblMontoFuturo 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   11280
            TabIndex        =   31
            Tag             =   "0.00"
            Top             =   1695
            Width           =   1935
         End
         Begin VB.Label lblRentabilidad365 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00000000"
            Height          =   285
            Left            =   11835
            TabIndex        =   34
            Tag             =   "0.00"
            Top             =   2500
            Width           =   1410
         End
         Begin VB.Label lblRentabilidad360 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00000000"
            Height          =   285
            Left            =   10395
            TabIndex        =   33
            Tag             =   "0.00"
            Top             =   2500
            Width           =   1410
         End
      End
      Begin VB.Frame fraDatosActivo 
         Caption         =   "Datos del Activo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   -74760
         TabIndex        =   65
         Top             =   2140
         Width           =   13575
         Begin VB.Label lblMonedaActivo 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nuevos Soles"
            Height          =   285
            Left            =   2085
            TabIndex        =   37
            Tag             =   "0.00"
            Top             =   700
            Width           =   1365
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Tasa"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
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
            Left            =   3720
            TabIndex        =   76
            Top             =   1060
            Width           =   1140
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
            Index           =   18
            Left            =   360
            TabIndex        =   75
            Top             =   1060
            Width           =   975
         End
         Begin VB.Label lblTipoTasa 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Efectiva"
            Height          =   285
            Left            =   5160
            TabIndex        =   41
            Tag             =   "0.00"
            Top             =   1040
            Width           =   1365
         End
         Begin VB.Label lblBaseAnual 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "360"
            Height          =   285
            Left            =   2085
            TabIndex        =   38
            Tag             =   "0.00"
            Top             =   1040
            Width           =   1365
         End
         Begin VB.Label lblPlazo 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            Height          =   285
            Left            =   8805
            TabIndex        =   43
            Tag             =   "0.00"
            Top             =   700
            Width           =   1365
         End
         Begin VB.Label lblFechaVencimiento 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "01/01/2002"
            Height          =   285
            Left            =   11880
            TabIndex        =   45
            Tag             =   "0.00"
            Top             =   360
            Width           =   1365
         End
         Begin VB.Label lblFechaEmision 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "01/01/2002"
            Height          =   285
            Left            =   8805
            TabIndex        =   42
            Tag             =   "0.00"
            Top             =   360
            Width           =   1365
         End
         Begin VB.Label lblTasaFacial 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.000000"
            Height          =   285
            Left            =   5160
            TabIndex        =   40
            Tag             =   "0.00"
            Top             =   700
            Width           =   1365
         End
         Begin VB.Label lblValorNominal 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   5160
            TabIndex        =   39
            Tag             =   "0.00"
            Top             =   360
            Width           =   1365
         End
         Begin VB.Label lblMontoMFL1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8805
            TabIndex        =   44
            Tag             =   "0.00"
            Top             =   1040
            Width           =   1365
         End
         Begin VB.Label lblMontoMFL2 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   11880
            TabIndex        =   46
            Tag             =   "0.00"
            Top             =   1040
            Width           =   1365
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto (MFL2)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
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
            Left            =   10440
            TabIndex        =   74
            Top             =   1060
            Width           =   1185
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Monto (MFL1)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
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
            Left            =   6840
            TabIndex        =   73
            Top             =   1060
            Width           =   1185
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
            Index           =   22
            Left            =   6840
            TabIndex        =   72
            Top             =   380
            Width           =   660
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
            Index           =   25
            Left            =   10440
            TabIndex        =   71
            Top             =   380
            Width           =   1050
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
            Index           =   23
            Left            =   6840
            TabIndex        =   70
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tasa"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
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
            Left            =   3720
            TabIndex        =   69
            Top             =   720
            Width           =   435
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
            Index           =   17
            Left            =   360
            TabIndex        =   68
            Top             =   720
            Width           =   690
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
            Index           =   19
            Left            =   3720
            TabIndex        =   67
            Top             =   380
            Width           =   1185
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
            Index           =   16
            Left            =   360
            TabIndex        =   66
            Top             =   380
            Width           =   780
         End
         Begin VB.Label lblAnalitica 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "??-??????"
            Height          =   285
            Left            =   2085
            TabIndex        =   36
            Tag             =   "0.00"
            Top             =   360
            Width           =   1365
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
         Height          =   1660
         Left            =   -74760
         TabIndex        =   58
         Top             =   480
         Width           =   13575
         Begin VB.ComboBox cboMonedaCoberturada 
            Height          =   315
            Left            =   8805
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   1095
            Width           =   4440
         End
         Begin VB.ComboBox cboTitulo 
            Height          =   315
            Left            =   8805
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   730
            Width           =   4440
         End
         Begin VB.ComboBox cboEmisor 
            Height          =   315
            Left            =   8805
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   360
            Width           =   4455
         End
         Begin VB.ComboBox cboTipoInstrumentoOrden 
            Height          =   315
            Left            =   2085
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1100
            Width           =   4440
         End
         Begin VB.ComboBox cboFondoOrden 
            Height          =   315
            Left            =   2085
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   360
            Width           =   4440
         End
         Begin VB.ComboBox cboTipoCobertura 
            Height          =   315
            Left            =   2085
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   730
            Width           =   4455
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Activo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
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
            Left            =   6840
            TabIndex        =   64
            Top             =   750
            Width           =   555
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Entidad Financiera"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
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
            Left            =   6840
            TabIndex        =   63
            Top             =   380
            Width           =   1605
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Activo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
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
            TabIndex        =   62
            Top             =   1120
            Width           =   990
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
            Index           =   9
            Left            =   360
            TabIndex        =   61
            Top             =   380
            Width           =   540
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Cobertura"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
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
            TabIndex        =   60
            Top             =   750
            Width           =   1275
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Moneda a Coberturar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
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
            Left            =   6840
            TabIndex        =   59
            Top             =   1125
            Width           =   1800
         End
      End
      Begin VB.Frame fraCriterio 
         Caption         =   "Criterios de Búsqueda"
         Height          =   2175
         Left            =   240
         TabIndex        =   48
         Top             =   480
         Width           =   13575
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
            Left            =   11960
            Picture         =   "frmCobertura.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Enviar a BackOffice"
            Top             =   1230
            Width           =   1200
         End
         Begin VB.ComboBox cboEstado 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   840
            Width           =   4185
         End
         Begin VB.ComboBox cboFondo 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   360
            Width           =   4185
         End
         Begin MSComCtl2.DTPicker dtpFechaOrdenDesde 
            Height          =   285
            Left            =   9360
            TabIndex        =   2
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
            Format          =   175833089
            CurrentDate     =   38785
         End
         Begin MSComCtl2.DTPicker dtpFechaOrdenHasta 
            Height          =   285
            Left            =   11715
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
            Format          =   175833089
            CurrentDate     =   38785
         End
         Begin MSComCtl2.DTPicker dtpFechaLiquidacionDesde 
            Height          =   285
            Left            =   9360
            TabIndex        =   4
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
            Format          =   175833089
            CurrentDate     =   38785
         End
         Begin MSComCtl2.DTPicker dtpFechaLiquidacionHasta 
            Height          =   285
            Left            =   11715
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
            Format          =   175833089
            CurrentDate     =   38785
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Estado"
            Height          =   195
            Index           =   2
            Left            =   480
            TabIndex        =   56
            Top             =   855
            Width           =   495
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            Height          =   195
            Index           =   7
            Left            =   11040
            TabIndex        =   55
            Top             =   375
            Width           =   420
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            Height          =   195
            Index           =   5
            Left            =   8640
            TabIndex        =   54
            Top             =   375
            Width           =   465
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fondo"
            Height          =   195
            Index           =   0
            Left            =   480
            TabIndex        =   53
            Top             =   380
            Width           =   450
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Orden"
            Height          =   195
            Index           =   3
            Left            =   6960
            TabIndex        =   52
            Top             =   375
            Width           =   930
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Liquidación"
            Height          =   195
            Index           =   4
            Left            =   6960
            TabIndex        =   51
            Top             =   795
            Width           =   1305
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            Height          =   195
            Index           =   6
            Left            =   8640
            TabIndex        =   50
            Top             =   795
            Width           =   465
         End
         Begin VB.Label lblDescrip 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            Height          =   195
            Index           =   8
            Left            =   11040
            TabIndex        =   49
            Top             =   795
            Width           =   420
         End
      End
      Begin TrueOleDBGrid60.TDBGrid tdgConsulta 
         Bindings        =   "frmCobertura.frx":0593
         Height          =   5055
         Left            =   240
         OleObjectBlob   =   "frmCobertura.frx":05AD
         TabIndex        =   57
         Top             =   2760
         Width           =   13545
      End
      Begin VB.Label lblDescrip 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   41
         Left            =   -74400
         TabIndex        =   92
         Top             =   7020
         Width           =   1065
      End
   End
End
Attribute VB_Name = "frmCobertura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Operaciones de Cobertura"
Option Explicit

Dim arrFondo()              As String, arrTipoInstrumentoOrden()        As String
Dim arrTipoOrden()          As String, arrEmisor()                      As String
Dim arrOrigen()             As String, arrTitulo()                      As String
Dim arrIndicador()          As String, arrTipoCobertura()               As String
Dim arrFondoOrden()         As String, arrEstado()                      As String
Dim arrMonedaCobertura()    As String, arrCuenta()                      As String
Dim arrMonedaCoberturada()  As String

Dim strCodFondo             As String, strCodTipoInstrumentoOrden       As String
Dim strCodTipoOrden         As String, strCodEmisor                     As String
Dim strCodOrigen            As String, strCodTitulo                     As String
Dim strCodIndicador         As String, strCodTipoCobertura              As String
Dim strCodFile              As String, strCodDetalleFile                As String
Dim strSQL                  As String, strCodEstado                     As String
Dim strCodFondoOrden        As String, strCodMoneda                     As String
Dim strEstado               As String, strEstadoOrden                   As String
Dim strCodMonedaCobertura   As String, strCodAnalitica                  As String
Dim strCodTipoTasa          As String, strCodigosFile                   As String
Dim strCodMonedaTitulo      As String, strCodBaseAnual                  As String
Dim strCodRiesgo            As String, strCodReportado                  As String
Dim strCodGarantia          As String, strNemotecnico                   As String
Dim strNumOrden             As String, strCodAnaliticaCuenta            As String
Dim strCodFileCuenta        As String
Dim dblTipoCambio           As Double


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

Public Sub Adicionar()

    If Not EsDiaUtil(gdatFechaActual) Then
        MsgBox "No se puede negociar en un día no útil !", vbCritical, Me.Caption
        Exit Sub
    End If
    
    frmMainMdi.stbMdi.Panels(3).Text = "Adicionar operación de cobertura..."
                
    strEstado = Reg_Adicion
    LlenarFormulario strEstado
    cmdOpcion.Visible = False
    With tabCobertura
        .TabEnabled(0) = False
        .Tab = 1
    End With
    
End Sub
Private Sub Limpiar()

    Dim adoRegistro         As ADODB.Recordset
    Dim strFecha            As String
    Dim intRegistro         As Integer
    
    Set adoRegistro = New ADODB.Recordset
    txtDescripOrden.Text = Valor_Caracter
    With adoComm
        .CommandText = "SELECT DescripFile,DescripInicial FROM InversionFile WHERE CodFile='013'"
        Set adoRegistro = .Execute
            
        If Not adoRegistro.EOF Then
            strFecha = Format(Day(gdatFechaActual), "00") & Format(Month(gdatFechaActual), "00") & Format(Year(gdatFechaActual), "0000")
            txtNemonico.Text = Trim(adoRegistro("DescripInicial")) & strFecha
            txtDescripOrden.Text = "Cobertura " & strNemotecnico & " - " & Trim(txtNemonico.Text)
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    strCodRiesgo = "00"
    strCodReportado = Valor_Caracter
    
    cboTipoInstrumentoOrden.ListIndex = -1
    If cboTipoInstrumentoOrden.ListCount > 0 Then cboTipoInstrumentoOrden.ListIndex = 0
    
    cboMonedaCoberturada.ListIndex = -1
    If cboMonedaCoberturada.ListCount > 0 Then cboMonedaCoberturada.ListIndex = 0
    
    cboMonedaCobertura.ListIndex = -1
    If cboMonedaCobertura.ListCount > 0 Then cboMonedaCobertura.ListIndex = 0
    
    lblAnalitica.Caption = Valor_Caracter
    lblValorNominal.Caption = "0"
    lblMonedaActivo.Caption = Valor_Caracter
    lblBaseAnual.Caption = Valor_Caracter
    lblTipoTasa.Caption = Valor_Caracter
    lblTasaFacial.Caption = "0"
    lblFechaEmision.Caption = Valor_Caracter: lblFechaVencimiento.Caption = Valor_Caracter
    lblPlazo.Caption = "0"
    lblMontoMFL1.Caption = "0": lblMontoMFL2.Caption = "0"
    
    dtpFechaOrden.Value = gdatFechaActual
    dtpFechaLiquidacion.Value = dtpFechaOrden.Value
    
    txtDiasPlazo.Text = "0"
    txtMontoCoberturado.Text = "0"
    txtDiferencial.Text = "0"
    txtTipoCambioSpot.Text = "0": txtTipoCambioFuturo.Text = "0"
    lblMontoContado.Caption = "0": lblMontoFuturo.Caption = "0"
    lblRentabilidad360.Caption = "0": lblRentabilidad365.Caption = "0"
    
    chkCobertura.Value = vbChecked
    chkCobertura.Value = vbUnchecked
    
    intRegistro = ObtenerItemLista(arrIndicador(), Codigo_IndCobertura_Amortizacion)
    If intRegistro >= 0 Then cboIndicador.ListIndex = intRegistro
    
    cboCuenta.ListIndex = -1
    If cboCuenta.ListCount > 0 Then cboCuenta.ListIndex = 0
    
    strCodTipoOrden = Codigo_Orden_Compra
    strCodOrigen = Codigo_Negociacion_Local
            
End Sub

Public Sub Salir()

    Unload Me
    
End Sub
Public Sub Cancelar()

    cmdOpcion.Visible = True
    With tabCobertura
        .TabEnabled(0) = True
        .Tab = 0
    End With
    Call Buscar
    
End Sub
Public Sub Grabar()

    Dim adoRegistro         As ADODB.Recordset
    Dim strFechaOrden       As String, strFechaLiquidacion      As String
    Dim strFechaEmision     As String, strFechaVencimiento      As String
    Dim strFechaPago        As String
    Dim strMensaje          As String, strIndTitulo             As String
    Dim dblMontoCobertura   As Double
    Dim intRegistro         As Integer
    
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
                "Fecha de Operación" & Space(4) & ">" & Space(2) & CStr(dtpFechaOrden.Value) & Chr(vbKeyReturn) & _
                "Fecha de Liquidación" & Space(3) & ">" & Space(2) & CStr(dtpFechaLiquidacion.Value) & Chr(vbKeyReturn) & Chr(vbKeyReturn) & _
                "Fecha de Vencimiento" & Space(1) & ">" & Space(2) & lblFechaVencimiento.Caption & Chr(vbKeyReturn) & _
                "Fecha de Pago" & Space(12) & ">" & Space(2) & CStr(dtpFechaPago.Value) & Chr(vbKeyReturn) & Chr(vbKeyReturn) & _
                "Monto a Coberturar" & Space(4) & ">" & Space(2) & txtMontoCoberturado.Text & Chr(vbKeyReturn) & Chr(vbKeyReturn) & _
                "Monto Contado" & Space(11) & ">" & Space(2) & lblMontoContado.Caption & Chr(vbKeyReturn) & _
                "Monto Futuro" & Space(14) & ">" & Space(2) & lblMontoFuturo.Caption & Chr(vbKeyReturn) & Chr(vbKeyReturn) & _
                Chr(vbKeyReturn) & Chr(vbKeyReturn) & _
                "¿ Seguro de continuar ?"

            If MsgBox(strMensaje, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
               Me.Refresh: Exit Sub
            End If

            Me.MousePointer = vbHourglass
            
            strFechaOrden = Convertyyyymmdd(dtpFechaOrden.Value)
            strFechaLiquidacion = Convertyyyymmdd(dtpFechaLiquidacion.Value)
            If strCodTipoCobertura = Codigo_Tipo_Cobertura_Independiente Then
                strFechaEmision = Convertyyyymmdd(dtpFechaOrden.Value)
            Else
                strFechaEmision = Convertyyyymmdd(lblFechaEmision.Caption)
            End If
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
                    strCodBaseAnual = Codigo_Base_Actual_Actual
                    strCodRiesgo = "00" ' Sin Clasificacion
                    strCodReportado = Valor_Caracter
                    strCodFile = Left(Trim(lblAnalitica.Caption), 3)
                ElseIf strCodTipoOrden = Codigo_Orden_Compra Then
'                    If chkTitulo.Value Then
'                        strIndTitulo = "X"
'                    Else
                        strCodTipoOrden = Codigo_Orden_Compromiso
                        strCodFile = "013"
                        strCodAnalitica = NumAleatorio(8)
                        strCodTitulo = NumAleatorio(15)
'                    End If
                Else
                    strIndTitulo = Valor_Indicador
'                    strCodTitulo = strCodGarantia
'                    strCodGarantia = Valor_Caracter
'                    strCodMoneda = lblMoneda.Tag
                    strFechaVencimiento = Convertyyyymmdd(Valor_Fecha)
                    strCodReportado = Valor_Caracter
                End If
                dblMontoCobertura = CDbl(txtMontoCoberturado.Text)
                '.CommandText = "BEGIN TRAN ProcOrden"
                'adoConn.Execute .CommandText
                
                On Error GoTo Ctrl_Error
                
'                .CommandText = "{ call up_IVAdicInversionOrden('" & strCodFondoOrden & "','" & _
'                    gstrCodAdministradora & "','','" & strFechaOrden & "','" & _
'                    strCodTitulo & "','" & Trim(txtNemonico.Text) & "','" & gstrPeriodoActual & "','" & gstrMesActual & "','" & _
'                    "','" & strEstadoOrden & "','" & strCodAnalitica & "','" & strCodFile & "','" & _
'                    strCodAnalitica & "','','','" & strCodTipoOrden & "','" & _
'                    "','','" & strCodOrigen & "','" & Trim(txtDescripOrden.Text) & "','" & strCodEmisor & "','" & _
'                    "','" & strCodGarantia & "','','" & strFechaPago & "','" & strFechaVencimiento & "','" & strFechaLiquidacion & "','" & _
'                    strFechaEmision & "','" & strCodMonedaTitulo & "'," & CDec(lblMontoFuturo.Caption) & "," & CDec(txtTipoCambio.Text) & "," & _
'                    "1," & CDec(txtPorcenCobertura.Text) & "," & CDec(dblMontoCobertura) & "," & _
'                    "0,0,0,0,0,0,0,0," & CDec(dblMontoCobertura) & ",0," & _
'                    "0,0,0,0,0,0,0,0,0,0," & _
'                    CDec(lblMontoFuturo.Caption) & "," & CInt(txtDiasPlazo.Text) & ",'','','','" & strCodReportado & "','" & strCodEmisor & "','" & strCodEmisor & "',0,'','','" & strIndTitulo & "','" & _
'                    strCodTipoTasa & "','" & strCodBaseAnual & "'," & CDec(lblTasaFacial.Caption) & "," & CDec(lblRentabilidad360.Caption) & "," & CDec(lblRentabilidad365.Caption) & ",'" & _
'                    strCodRiesgo & "','','" & Trim(txtObservacion.Text) & "','" & gstrLogin & "') }"
'                adoConn.Execute .CommandText
                
                .CommandText = "{ call up_IVAdicInversionOrden('" & strCodFondoOrden & "','" & _
                    gstrCodAdministradora & "','','" & strFechaOrden & "','" & _
                    strCodTitulo & "','" & Trim(txtNemonico.Text) & "','" & gstrPeriodoActual & "','" & gstrMesActual & "','" & _
                    "','" & strEstadoOrden & "','" & strCodAnalitica & "','" & strCodFile & "','" & _
                    strCodAnalitica & "','','','" & strCodTipoOrden & "','" & _
                    "','','" & strCodOrigen & "','" & Trim(txtDescripOrden.Text) & "','" & strCodEmisor & "','" & _
                    "','" & strCodGarantia & "','','" & strFechaPago & "','" & strFechaVencimiento & "','" & strFechaLiquidacion & "','" & _
                    strFechaEmision & "','" & strCodMonedaTitulo & "','" & strCodMonedaTitulo & "','" & strCodMonedaTitulo & "'," & CDec(lblMontoFuturo.Caption) & "," & CDec(txtTipoCambio.Text) & "," & _
                    CDec(txtTipoCambio.Text) & ",1,100,1," & CDec(txtPorcenCobertura.Text) & "," & CDec(dblMontoCobertura) & "," & _
                    "0,0,0,0,0,0,0,0,0,0,0," & CDec(dblMontoCobertura) & "," & CDec(dblMontoCobertura) & ",0," & _
                    "0,0,0,0,0,0,0,0,0,0,0,0," & _
                    CDec(lblMontoFuturo.Caption) & "," & CInt(txtDiasPlazo.Text) & ",'','','','" & strCodReportado & "','" & strCodEmisor & "','" & strCodEmisor & "','','','',0,'','','" & strIndTitulo & "','" & _
                    strCodTipoTasa & "','" & strCodBaseAnual & "'," & CDec(lblTasaFacial.Caption) & "," & CDec(lblRentabilidad360.Caption) & "," & CDec(lblRentabilidad365.Caption) & ",'" & _
                    strCodRiesgo & "','','" & Trim(txtObservacion.Text) & "','" & gstrLogin & "') }"
                adoConn.Execute .CommandText
                
                
                .CommandText = "SELECT NumOrden FROM InversionOrden WHERE CodTitulo='" & strCodTitulo & "' AND " & _
                    "CodFondo='" & strCodFondoOrden & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
                Set adoRegistro = .Execute
                
                If Not adoRegistro.EOF Then
                    strNumOrden = adoRegistro("NumOrden")
                End If
                adoRegistro.Close: Set adoRegistro = Nothing
                
                .CommandText = "{ call up_IVAdicInversionCobertura('" & strCodFondoOrden & "','" & _
                    gstrCodAdministradora & "','" & strCodTitulo & "','','" & strNumOrden & "','" & strCodFile & "','" & _
                    strCodAnalitica & "','" & strCodTipoCobertura & "','" & strCodIndicador & "','" & strFechaOrden & "','" & _
                    strFechaLiquidacion & "','" & strFechaVencimiento & "','" & strCodEmisor & "','" & _
                    strCodMonedaCobertura & "','" & strCodMonedaTitulo & "','" & _
                    Trim(txtDescripOrden.Text) & "',''," & CInt(txtDiasPlazo.Text) & "," & _
                    CDec(txtPorcenCobertura.Text) & "," & CDec(lblRentabilidad360.Caption) & "," & CDec(lblRentabilidad365.Caption) & "," & _
                    CDec(txtTipoCambio.Text) & "," & CDec(txtTipoCambioSpot.Text) & "," & CDec(txtTipoCambioFuturo.Text) & "," & _
                    CDec(txtMontoCoberturado.Text) & "," & CDec(lblMontoContado.Caption) & "," & _
                    CDec(lblMontoFuturo.Caption) & ",'" & strCodFileCuenta & "','" & strCodAnaliticaCuenta & "','" & strEstadoOrden & "') }"
                adoConn.Execute .CommandText
                
                '.CommandText = "COMMIT TRAN ProcOrden"
                'adoConn.Execute .CommandText
                                                                                                      
            End With
                                                                                    
            Me.MousePointer = vbDefault
        
            MsgBox Mensaje_Adicion_Exitosa, vbExclamation
            
            frmMainMdi.stbMdi.Panels(3).Text = "Acción"
            
            cmdOpcion.Visible = True
            With tabCobertura
                .TabEnabled(0) = True
                .Tab = 0
            End With
            Call Buscar
        End If
    End If
    Exit Sub
        
Ctrl_Error:
    'adoComm.CommandText = "ROLLBACK TRAN ProcOrden"
    'adoConn.Execute adoComm.CommandText
    'adoConn.Errors.Item(0).Description
    MsgBox adoConn.Errors.Item(0).Description & vbNewLine & vbNewLine & Mensaje_Proceso_NoExitoso, vbCritical
    Me.MousePointer = vbDefault
    
End Sub

Private Function TodoOK() As Boolean
        
    Dim adoRegistro         As ADODB.Recordset
    Dim strFechaDesde       As String, strFechaHasta            As String
    
    TodoOK = False
          
    If cboFondoOrden.ListIndex < 0 Then
        MsgBox "Debe seleccionar el Fondo.", vbCritical, Me.Caption
        If cboFondoOrden.Enabled Then cboFondoOrden.SetFocus
        Exit Function
    End If
    
    If cboTipoCobertura.ListIndex < 0 Then
        MsgBox "Debe seleccionar el Tipo de Cobertura.", vbCritical, Me.Caption
        If cboTipoCobertura.Enabled Then cboTipoCobertura.SetFocus
        Exit Function
    End If
    
    If strCodTipoCobertura = Codigo_Tipo_Cobertura_Sintetico Then
        If cboTipoInstrumentoOrden.ListIndex <= 0 Then
            MsgBox "Debe seleccionar el Tipo de Instrumento de Corto Plazo.", vbCritical, Me.Caption
            If cboTipoInstrumentoOrden.Enabled Then cboTipoInstrumentoOrden.SetFocus
            Exit Function
        End If
                
        If cboTitulo.ListIndex < 0 Then
            MsgBox "Debe seleccionar el Título a coberturar.", vbCritical, Me.Caption
            If cboTitulo.Enabled Then cboTitulo.SetFocus
            Exit Function
        End If
    End If
  
    If cboEmisor.ListIndex < 0 Then
        MsgBox "Debe seleccionar el Emisor o Contraparte.", vbCritical, Me.Caption
        If cboEmisor.Enabled Then cboEmisor.SetFocus
        Exit Function
    End If
  
    Set adoRegistro = New ADODB.Recordset
        
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
        
    If cboIndicador.ListIndex < 0 Then
        MsgBox "Debe seleccionar el Indicador de Cobertura.", vbCritical, Me.Caption
        If cboIndicador.Enabled Then cboIndicador.SetFocus
        Exit Function
    End If
    
    If chkCobertura.Value Then
        If CDbl(txtPorcenCobertura.Text) = 0 Then
            MsgBox "Debe indicar el Porcentaje a coberturar.", vbCritical, Me.Caption
            If txtPorcenCobertura.Enabled Then txtPorcenCobertura.SetFocus
            Exit Function
        End If
    End If
    
    If CDbl(txtTipoCambioSpot.Text) = 0 Then
        MsgBox "Debe indicar el Tipo de Cambio Spot.", vbCritical, Me.Caption
        If txtTipoCambioSpot.Enabled Then txtTipoCambioSpot.SetFocus
        Exit Function
    End If
    
    If CDbl(txtDiferencial.Text) = 0 Then
        MsgBox "Debe indicar el Diferencial.", vbCritical, Me.Caption
        If txtDiferencial.Enabled Then txtDiferencial.SetFocus
        Exit Function
    End If
    
    If CDbl(lblRentabilidad360.Caption) = 0 Or CDbl(lblRentabilidad365.Caption) = 0 Then
        MsgBox "Por favor hallar la Rentabilidad (360).", vbCritical, Me.Caption
        If cmdRentabilidad.Enabled Then cmdRentabilidad.SetFocus
        Exit Function
    End If
    
    '*** Si todo paso OK ***
    TodoOK = True
  
End Function
Public Sub Imprimir()

End Sub

Public Sub SubImprimir(Index As Integer)

'    Dim frmReporte              As frmVisorReporte
'    Dim aReportParamS(), aReportParamF(), aReportParamFn()
'    Dim strFechaDesde           As String, strFechaHasta        As String
'    Dim strSeleccionRegistro    As String
'
'    If tabRFCortoPlazo.Tab = 1 Then Exit Sub
'
'    Select Case Index
'        Case 1
'            gstrNameRepo = "InversionOrden"
'
'            strSeleccionRegistro = "{InversionOrden.FechaOrden} IN 'Fch1' TO 'Fch2'"
'            gstrSelFrml = strSeleccionRegistro
'            frmRangoFecha.Show vbModal
'
'            If gstrSelFrml <> "0" Then
'                Set frmReporte = New frmVisorReporte
'
'                ReDim aReportParamS(5)
'                ReDim aReportParamFn(5)
'                ReDim aReportParamF(5)
'
'                aReportParamFn(0) = "Usuario"
'                aReportParamFn(1) = "FechaDesde"
'                aReportParamFn(2) = "FechaHasta"
'                aReportParamFn(3) = "Hora"
'                aReportParamFn(4) = "Fondo"
'                aReportParamFn(5) = "NombreEmpresa"
'
'                aReportParamF(0) = gstrLogin
'                aReportParamF(1) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10)
'                aReportParamF(2) = Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10)
'                aReportParamF(3) = Format(Time(), "hh:mm:ss")
'                aReportParamF(4) = Trim(cboFondo.Text)
'                aReportParamF(5) = gstrNombreEmpresa & Space(1)
'
'                aReportParamS(0) = strCodFondo
'                aReportParamS(1) = gstrCodAdministradora
'                aReportParamS(2) = Convertyyyymmdd(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10))
'                aReportParamS(3) = Convertyyyymmdd(DateAdd("d", 1, CVDate(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10))))
'                aReportParamS(4) = strCodMoneda
'                aReportParamS(5) = strCodTipoInstrumento
'            End If
'        Case 2
'            gstrNameRepo = "PapeletaInversion"
'
'            strSeleccionRegistro = "{InversionOrden.FechaOrden} IN 'Fch1' TO 'Fch2'"
'            gstrSelFrml = strSeleccionRegistro
'            frmRangoFecha.Show vbModal
'
'            If gstrSelFrml <> "0" Then
'                Set frmReporte = New frmVisorReporte
'
'                ReDim aReportParamS(5)
'                ReDim aReportParamFn(1)
'                ReDim aReportParamF(1)
'
'                aReportParamFn(0) = "Fondo"
'                aReportParamFn(1) = "NombreEmpresa"
'
'                aReportParamF(0) = Trim(cboFondo.Text)
'                aReportParamF(1) = gstrNombreEmpresa & Space(1)
'
'                aReportParamS(0) = strCodFondo
'                aReportParamS(1) = gstrCodAdministradora
'                aReportParamS(2) = Convertyyyymmdd(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch1"), 10))
'                aReportParamS(3) = Convertyyyymmdd(DateAdd("d", 1, CVDate(Mid(gstrSelFrml, InStr(strSeleccionRegistro, "Fch2") + 6, 10))))
'                aReportParamS(4) = strCodMoneda
'                aReportParamS(5) = strCodTipoInstrumento
'            End If
'
'    End Select
'
'    If gstrSelFrml = "0" Then Exit Sub
'
'    gstrSelFrml = Valor_Caracter
'    frmReporte.strReportPath = gstrRptPath & gstrNameRepo & ".RPT"
'
'    Call frmReporte.SetReportParam(aReportParamS(), aReportParamF(), aReportParamFn())
'
'    frmReporte.Caption = "Reporte - (" & gstrNameRepo & ")"
'    frmReporte.Show vbModal
'
'    Set frmReporte = Nothing
'
'    Screen.MousePointer = vbNormal
    
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
    
    strSQL = "SELECT IOR.NumOrden,FechaOrden,IOR.FechaLiquidacion,IOR.CodTitulo,Nemotecnico,EstadoOrden,IOR.CodFile,IOR.CodAnalitica,TipoOrden,IOR.CodMoneda," & _
        "(RTRIM(DescripTipoOperacion) + SPACE(1) + DescripOrden) DescripOrden,CantOrden,ValorNominal,PrecioUnitarioMFL1,MontoTotalMFL1,MON.CodSigno DescripMonedaActivo,MON0.CodSigno DescripMoneda " & _
        "FROM InversionOrden IOR JOIN InversionCobertura ICO ON(ICO.CodTitulo=IOR.CodTitulo AND ICO.CodFondo=IOR.CodFondo AND ICO.CodAdministradora=IOR.CodAdministradora AND ICO.NumOrden=IOR.NumOrden) " & _
        "JOIN TipoOperacionNegociacion TON ON(TON.CodTipoOperacion=IOR.TipoOrden) " & _
        "JOIN Moneda MON ON(MON.CodMoneda=IOR.CodMoneda) JOIN Moneda MON0 ON(MON0.CodMoneda=ICO.CodMonedaCobertura) " & _
        "WHERE IOR.CodAdministradora='" & gstrCodAdministradora & "' AND IOR.CodFondo='" & strCodFondo & "' "

    If Not IsNull(dtpFechaOrdenDesde.Value) Or Not IsNull(dtpFechaOrdenHasta.Value) Then
        strSQL = strSQL & "AND (FechaOrden >='" & strFechaOrdenDesde & "' AND FechaOrden <'" & strFechaOrdenHasta & "') "
    End If

    If Not IsNull(dtpFechaLiquidacionDesde.Value) Or Not IsNull(dtpFechaLiquidacionHasta.Value) Then
        strSQL = strSQL & "AND (IOR.FechaLiquidacion >='" & strFechaLiquidacionDesde & "' AND IOR.FechaLiquidacion <'" & strFechaLiquidacionHasta & "') "
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
Public Sub Eliminar()

    If strEstado = Reg_Consulta Or strEstado = Reg_Edicion Then
        Dim strMensaje  As String
        
        strMensaje = "Se procederá a eliminar la ORDEN " & tdgConsulta.Columns(0) & " por la " & _
            tdgConsulta.Columns(2) & vbNewLine & vbNewLine & vbNewLine & "¿ Seguro de continuar ?"
        
        If MsgBox(strMensaje, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
    
            '*** Anular Orden ***
            adoComm.CommandText = "UPDATE InversionOrden SET EstadoOrden='" & Estado_Orden_Anulada & "' WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodTitulo='" & Trim(tdgConsulta.Columns(1)) & "' AND NumOrden='" & Trim(tdgConsulta.Columns(0)) & "'"
            adoConn.Execute adoComm.CommandText
            
            '*** Anular Cobertura corresponde ***
            adoComm.CommandText = "UPDATE InversionCobertura SET Estado='" & Estado_Orden_Anulada & "' WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodTitulo='" & Trim(tdgConsulta.Columns(1)) & "'"
            adoConn.Execute adoComm.CommandText
            
            '*** Anular Título si corresponde ***
            adoComm.CommandText = "UPDATE InstrumentoInversion SET IndVigente='' WHERE CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "' AND " & _
                "CodTitulo='" & Trim(tdgConsulta.Columns(1)) & "'"
            adoConn.Execute adoComm.CommandText
            
            MsgBox Mensaje_Eliminacion_Exitosa, vbExclamation, Me.Caption
            
            tabCobertura.Tab = 0
            Call Buscar
            
            Exit Sub
        End If
    End If
    
End Sub

Public Sub Modificar()

    If strEstado = Reg_Defecto Then Exit Sub
    
    If Trim(tdgConsulta.Columns(6)) <> Estado_Orden_Ingresada Then
        MsgBox "Solo se pueden confirmar las Ordenes con estado INGRESADA.", vbCritical, Me.Caption
        Exit Sub
    End If
    
    If strEstado = Reg_Consulta Then
        strEstado = Reg_Edicion
        LlenarFormulario strEstado
        cmdOpcion.Visible = False
        With tabCobertura
            .TabEnabled(0) = False
            .Tab = 1
        End With
        'Call Habilita
    End If
    
End Sub

Private Sub LlenarFormulario(strModo As String)

    Dim adoRegistro     As ADODB.Recordset
    Dim strFecha        As String
    Dim intRegistro     As Integer
    
    Select Case strModo
        Case Reg_Adicion
            Set adoRegistro = New ADODB.Recordset
            txtDescripOrden.Text = Valor_Caracter
            With adoComm
                .CommandText = "SELECT DescripFile,DescripInicial FROM InversionFile WHERE CodFile='013'"
                Set adoRegistro = .Execute
                    
                If Not adoRegistro.EOF Then
                    strFecha = Format(Day(gdatFechaActual), "00") & Format(Month(gdatFechaActual), "00") & Format(Year(gdatFechaActual), "0000")
                    txtNemonico.Text = Trim(adoRegistro("DescripInicial")) & strFecha
                    txtDescripOrden.Text = "Cobertura " & strNemotecnico & " - " & Trim(txtNemonico.Text)
                End If
                adoRegistro.Close: Set adoRegistro = Nothing
            End With
            strCodRiesgo = "00"
            strCodReportado = Valor_Caracter
            intRegistro = ObtenerItemLista(arrFondoOrden(), strCodFondo)
            If intRegistro >= 0 Then cboFondoOrden.ListIndex = intRegistro
        
            cboTipoInstrumentoOrden.ListIndex = -1
            If cboTipoInstrumentoOrden.ListCount > 0 Then cboTipoInstrumentoOrden.ListIndex = 0
                                                                                                    
            cboEmisor.ListIndex = -1
            If cboEmisor.ListCount > 0 Then cboEmisor.ListIndex = 0
            
            intRegistro = ObtenerItemLista(arrTipoCobertura(), Codigo_Tipo_Cobertura_Sintetico)
            If intRegistro >= 0 Then cboTipoCobertura.ListIndex = intRegistro
            
            cboMonedaCoberturada.ListIndex = -1
            If cboMonedaCoberturada.ListCount > 0 Then cboMonedaCoberturada.ListIndex = 0
            
            cboMonedaCobertura.ListIndex = -1
            If cboMonedaCobertura.ListCount > 0 Then cboMonedaCobertura.ListIndex = 0
            
            lblAnalitica.Caption = Valor_Caracter
            lblValorNominal.Caption = "0"
            lblMonedaActivo.Caption = Valor_Caracter
            lblBaseAnual.Caption = Valor_Caracter
            lblTipoTasa.Caption = Valor_Caracter
            lblTasaFacial.Caption = "0"
            lblFechaEmision.Caption = Valor_Caracter: lblFechaVencimiento.Caption = Valor_Caracter
            lblPlazo.Caption = "0"
            lblMontoMFL1.Caption = "0": lblMontoMFL2.Caption = "0"
            
            dtpFechaOrden.Value = gdatFechaActual
            dtpFechaLiquidacion.Value = dtpFechaOrden.Value

            txtDiasPlazo.Text = "0"
            txtMontoCoberturado.Text = "0"
            txtDiferencial.Text = "0"
            txtTipoCambioSpot.Text = "0": txtTipoCambioFuturo.Text = "0"
            lblMontoContado.Caption = "0": lblMontoFuturo.Caption = "0"
            lblRentabilidad360.Caption = "0": lblRentabilidad365.Caption = "0"
            
            chkCobertura.Value = vbChecked
            chkCobertura.Value = vbUnchecked
            
            intRegistro = ObtenerItemLista(arrIndicador(), Codigo_IndCobertura_Amortizacion)
            If intRegistro >= 0 Then cboIndicador.ListIndex = intRegistro
            
            cboCuenta.ListIndex = -1
            If cboCuenta.ListCount > 0 Then cboCuenta.ListIndex = 0
            
            strCodTipoOrden = Codigo_Orden_Compra
            strCodOrigen = Codigo_Negociacion_Local
            cboFondoOrden.SetFocus
    End Select
    
End Sub

Private Sub CargarReportes()

    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Visible = True
    frmMainMdi.tlbMdi.Buttons("Reportes").ButtonMenus("Repo1").Text = "Operaciones de Cobertura"
    
End Sub


Private Sub cboCuenta_Click()

    strCodFileCuenta = Valor_Caracter: strCodAnaliticaCuenta = Valor_Caracter
    If cboCuenta.ListIndex < 0 Then Exit Sub
    
    strCodFileCuenta = Left(Trim(arrCuenta(cboCuenta.ListIndex)), 3)
    strCodAnaliticaCuenta = Right(Trim(arrCuenta(cboCuenta.ListIndex)), 8)
    cboCuenta.ToolTipText = cboCuenta.Text
    
End Sub


Private Sub cboEmisor_Click()

    strCodEmisor = Valor_Caracter
    If cboEmisor.ListIndex < 0 Then Exit Sub
    
    strCodEmisor = Trim(arrEmisor(cboEmisor.ListIndex))
    
End Sub




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
            strCodMoneda = Trim(adoRegistro("CodMoneda"))
            dblTipoCambio = CDbl(adoRegistro("ValorTipoCambio"))
            txtTipoCambio.Text = CStr(dblTipoCambio)
                        
            frmMainMdi.txtFechaSistema.Text = CStr(gdatFechaActual)
            Me.Refresh
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
    End With
    
    '*** Tipo de Instrumento ***
    strSQL = "SELECT FIF.CodFile CODIGO,DescripFile DESCRIP " & _
        "FROM FondoInversionFile FIF JOIN InversionFile IVF ON(IVF.CodFile=FIF.CodFile) " & _
        "WHERE IndInstrumento='X' AND IndVigente='X' AND IndCobertura='X' AND " & _
        "CodFondo='" & strCodFondoOrden & "' AND CodAdministradora='" & gstrCodAdministradora & "' ORDER BY DescripFile"
    CargarControlLista strSQL, cboTipoInstrumentoOrden, arrTipoInstrumentoOrden(), Sel_Defecto

End Sub


Private Sub cboIndicador_Click()

    strCodIndicador = Valor_Caracter
    If cboIndicador.ListIndex < 0 Then Exit Sub
    
    strCodIndicador = Trim(arrIndicador(cboIndicador.ListIndex))
    
    If strCodIndicador = Codigo_IndCobertura_Amortizacion Then txtMontoCoberturado.Text = lblMontoMFL2.Caption
    If strCodIndicador = Codigo_IndCobertura_Principal Then txtMontoCoberturado.Text = lblMontoMFL1.Caption
    
    lblRentabilidad360.Caption = "0": lblRentabilidad365.Caption = "0"
    If CDbl(txtDiferencial.Text) > 0 Then txtDiferencial_Change
    
End Sub

Private Sub cboMonedaCobertura_Click()

    strCodMonedaCobertura = Valor_Caracter
    If cboMonedaCobertura.ListIndex < 0 Then Exit Sub
    
    strCodMonedaCobertura = arrMonedaCobertura(cboMonedaCobertura.ListIndex)
    
    lblDescrip(36).Caption = "Contado " & ObtenerSignoMoneda(strCodMonedaCobertura)
    lblDescrip(37).Caption = "Futuro " & ObtenerSignoMoneda(strCodMonedaCobertura)
    
End Sub

Private Sub cboMonedaCoberturada_Click()

    strCodMonedaTitulo = Valor_Caracter
    If cboMonedaCoberturada.ListIndex < 0 Then Exit Sub
    
    strCodMonedaTitulo = arrMonedaCoberturada(cboMonedaCoberturada.ListIndex)
    
    lblDescrip(30).Caption = "Monto Coberturado " & ObtenerSignoMoneda(strCodMonedaTitulo)
    
    '*** Cargar cuentas en la moneda del monto coberturado ***
    strSQL = "SELECT (CodFile + CodAnalitica) CODIGO,(RTRIM(DescripCuenta) + SPACE(1) + NumCuenta) DESCRIP FROM BancoCuenta " & _
        "WHERE CodMoneda='" & strCodMonedaTitulo & "' AND " & _
        "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
    CargarControlLista strSQL, cboCuenta, arrCuenta(), Sel_Defecto
    
    If cboCuenta.ListCount > 0 Then cboCuenta.ListIndex = 0
    
    '*** Moneda de Cobertura ***
    strSQL = "SELECT CodMoneda CODIGO,DescripMoneda DESCRIP FROM Moneda WHERE CodMoneda<>'" & strCodMonedaTitulo & "' AND CodSigno<>'' ORDER BY DescripMoneda"
    CargarControlLista strSQL, cboMonedaCobertura, arrMonedaCobertura(), Sel_Defecto
    
    If cboMonedaCobertura.ListCount > 0 Then cboMonedaCobertura.ListIndex = 0
        
End Sub


Private Sub Habilita()

    Call Limpiar
    
    cboTipoInstrumentoOrden.Enabled = True
    cboTitulo.Enabled = True
    cboMonedaCoberturada.Enabled = False
    chkCobertura.Enabled = True
    txtMontoCoberturado.Enabled = False

End Sub

Private Sub Deshabilita()

    Call Limpiar
    
    cboTipoInstrumentoOrden.Enabled = False
    cboTitulo.Enabled = False
    cboMonedaCoberturada.Enabled = True
    chkCobertura.Enabled = False
    txtMontoCoberturado.Enabled = True

End Sub
Private Sub cboTipoCobertura_Click()

    strCodTipoCobertura = Valor_Caracter
    If cboTipoCobertura.ListIndex < 0 Then Exit Sub
    
    strCodTipoCobertura = Trim(arrTipoCobertura(cboTipoCobertura.ListIndex))
    
    If strCodTipoCobertura = Codigo_Tipo_Cobertura_Sintetico Then
        Call Habilita
        cboTitulo_Click
    Else
        Call Deshabilita
        strCodFile = "013"
    End If
    
End Sub

Private Sub cboTipoInstrumentoOrden_Click()

    strCodTipoInstrumentoOrden = Valor_Caracter
    If cboTipoInstrumentoOrden.ListIndex < 0 Then Exit Sub
    
    strCodTipoInstrumentoOrden = Trim(arrTipoInstrumentoOrden(cboTipoInstrumentoOrden.ListIndex))
    
    strSQL = "SELECT InstrumentoInversion.CodTitulo CODIGO," & _
                "(RTRIM(InstrumentoInversion.CodTitulo) + ' ' + RTRIM(InstrumentoInversion.Nemotecnico) + ' ' + RTRIM(InstrumentoInversion.DescripTitulo)) DESCRIP FROM InstrumentoInversion,InversionKardex " & _
                "WHERE SaldoFinal > 0 AND IndUltimoMovimiento='X' AND InstrumentoInversion.CodFile=InversionKardex.CodFile AND " & _
                "InstrumentoInversion.CodAnalitica=InversionKardex.CodAnalitica AND InversionKardex.CodFile='" & strCodTipoInstrumentoOrden & "' AND " & _
                "InstrumentoInversion.CodFondo='" & strCodFondoOrden & "' AND InversionKardex.CodFondo='" & strCodFondoOrden & "' " & _
                "ORDER BY InstrumentoInversion.Nemotecnico"
    CargarControlLista strSQL, cboTitulo, arrTitulo(), Sel_Defecto
        
    If cboTitulo.ListCount > 0 Then cboTitulo.ListIndex = 0
    
End Sub

Private Sub cboTitulo_Click()

    Dim strFechaOperacion   As String, strFechaSiguiente    As String
    Dim adoRegistro         As ADODB.Recordset
    Dim dblNominal          As Double
    Dim intRegistro         As Integer
    
    strCodTitulo = Valor_Caracter: strCodAnalitica = Valor_Caracter
    lblAnalitica = strCodTipoInstrumentoOrden & "-????????": lblValorNominal.Caption = "0"
    strCodBaseAnual = Valor_Caracter: strNemotecnico = Valor_Caracter
    If cboTitulo.ListIndex < 0 Then Exit Sub
    
    strCodTitulo = Trim(arrTitulo(cboTitulo.ListIndex))
    
    With adoComm
        Set adoRegistro = New ADODB.Recordset

        .CommandText = "SELECT CodTitulo,CodAnalitica,ValorNominal,CodMoneda,CodEmisor,CodGrupo,FechaEmision,FechaVencimiento," & _
            "TasaInteres,CodRiesgo,CodSubRiesgo,CodTipoTasa,BaseAnual,Nemotecnico,DiasPlazo,CodFile,CodDetalleFile " & _
            "FROM InstrumentoInversion WHERE CodTitulo='" & strCodTitulo & "'"
        Set adoRegistro = .Execute

        If Not adoRegistro.EOF Then
            strCodGarantia = adoRegistro("CodTitulo")
            strCodFile = adoRegistro("CodFile")
            strCodDetalleFile = adoRegistro("CodDetalleFile")
            strCodTipoTasa = adoRegistro("CodTipoTasa")
            strCodAnalitica = Trim(adoRegistro("CodAnalitica"))
            lblAnalitica.Caption = strCodTipoInstrumentoOrden & "-" & strCodAnalitica
            
            strCodMonedaTitulo = adoRegistro("CodMoneda")
            lblMonedaActivo.Caption = ObtenerDescripcionMoneda(strCodMonedaTitulo)
            lblDescrip(30).Caption = "Monto Coberturado" & Space(1) & ObtenerSignoMoneda(strCodMonedaTitulo)
            lblFechaEmision.Caption = adoRegistro("FechaEmision")
            lblFechaVencimiento.Caption = adoRegistro("FechaVencimiento")
            dtpFechaVencimiento.Value = adoRegistro("FechaVencimiento")
            lblPlazo.Caption = CStr(adoRegistro("DiasPlazo"))
            txtDiasPlazo.Text = CStr(adoRegistro("DiasPlazo"))
            
            lblBaseAnual.Caption = "360"
            strCodBaseAnual = adoRegistro("BaseAnual")
            If strCodBaseAnual = Codigo_Base_Actual_Actual Then lblBaseAnual.Caption = "365"
            If strCodBaseAnual = Codigo_Base_Actual_365 Then lblBaseAnual.Caption = "365"
            If strCodBaseAnual = Codigo_Base_Actual_360 Then lblBaseAnual.Caption = "360"
            If strCodBaseAnual = Codigo_Base_30_360 Then lblBaseAnual.Caption = "360"
            If strCodBaseAnual = Codigo_Base_30_365 Then lblBaseAnual.Caption = "365"
            
            lblTasaFacial.Caption = CStr(adoRegistro("TasaInteres"))
            dblNominal = adoRegistro("ValorNominal")
            strNemotecnico = Trim(adoRegistro("Nemotecnico"))
            txtDescripOrden.Text = "Cobertura " & strNemotecnico & " - " & Trim(txtNemonico.Text)
            
            intRegistro = ObtenerItemLista(arrMonedaCoberturada(), strCodMonedaTitulo)
            If intRegistro >= 0 Then cboMonedaCoberturada.ListIndex = intRegistro
        End If
        adoRegistro.Close
        
        '*** Moneda de Cobertura ***
        strSQL = "SELECT CodMoneda CODIGO,DescripMoneda DESCRIP FROM Moneda WHERE CodMoneda<>'" & strCodMonedaTitulo & "' AND CodSigno<>'' ORDER BY DescripMoneda"
        CargarControlLista strSQL, cboMonedaCobertura, arrMonedaCobertura(), Sel_Defecto
        
        If cboMonedaCobertura.ListCount > 0 Then cboMonedaCobertura.ListIndex = 0
        
        .CommandText = "SELECT SaldoFinal,ValorPromedio,NumOperacion FROM InversionKardex " & _
            "WHERE CodTitulo='" & strCodTitulo & "' AND CodFondo='" & strCodFondoOrden & "' AND " & _
            "CodAdministradora='" & gstrCodAdministradora & "' AND IndUltimoMovimiento='X' AND SaldoFinal > 0"
        Set adoRegistro = .Execute

        If Not adoRegistro.EOF Then
            dblNominal = dblNominal * adoRegistro("SaldoFinal")
            lblValorNominal.Caption = CStr(dblNominal)
            strNumOperacion = adoRegistro("NumOperacion")
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
            
        .CommandText = "SELECT DescripParametro FROM AuxiliarParametro " & _
            "WHERE CodTipoParametro='NATTAS' AND CodParametro='" & strCodTipoTasa & "'"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            lblTipoTasa.Caption = Trim(adoRegistro("DescripParametro"))
        End If
        adoRegistro.Close
        
        .CommandText = "SELECT MontoTotalMFL1,MontoVencimiento,TasaInteres FROM InversionOperacion " & _
            "WHERE CodTitulo='" & strCodTitulo & "' AND NumOperacion='" & strNumOperacion & "' AND " & _
            "CodFondo='" & strCodFondo & "' AND CodAdministradora='" & gstrCodAdministradora & "'"
        Set adoRegistro = .Execute
        
        If Not adoRegistro.EOF Then
            lblMontoMFL2.Caption = CStr(adoRegistro("MontoVencimiento"))
            lblTasaFacial.Caption = CStr(adoRegistro("TasaInteres"))
        End If
        adoRegistro.Close: Set adoRegistro = Nothing
        
        '*** Obtener las cuentas de inversión ***
        Call ObtenerCuentasInversion(strCodFile, strCodDetalleFile, strCodMonedaTitulo)
            
        strFechaOperacion = Convertyyyymmdd(dtpFechaOrden.Value)
        strFechaSiguiente = Convertyyyymmdd(DateAdd("d", 1, dtpFechaOrden.Value))
        
        '*** Obtener Saldos ***
        curCtaInversion = ObtenerSaldoFinalCuenta(strCodFondo, gstrCodAdministradora, strCodFile, strCodAnalitica, strFechaOperacion, strFechaSiguiente, strCtaInversion, strCodMonedaTitulo)
        lblMontoMFL1.Caption = CStr(curCtaInversion)
        
        chkCobertura.Value = vbChecked
        chkCobertura.Value = vbUnchecked
        
        '*** Validar Limites ***
'        If Not PosicionLimites() Then Exit Sub
    End With
    
End Sub

Private Sub chkCobertura_Click()

    If chkCobertura.Value Then
        txtPorcenCobertura.Text = "0"
        txtMontoCoberturado.Text = "0"
        txtPorcenCobertura.Enabled = True
    Else
        txtPorcenCobertura.Text = "100"
        txtPorcenCobertura.Enabled = False
        cboIndicador_Click
        txtDiferencial_Change
    End If
    
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
                "WHERE NumOrden='" & Trim(tdgConsulta.Columns(0)) & "' AND CodFondo='" & strCodFondo & "' AND " & _
                "CodAdministradora='" & gstrCodAdministradora & "' AND EstadoOrden='" & Estado_Orden_Ingresada & "'"
        ElseIf strCodEstado = Estado_Orden_Enviada Then
            adoComm.CommandText = "UPDATE InversionOrden SET EstadoOrden='" & Estado_Orden_Ingresada & "'," & _
                "UsuarioEdicion='" & gstrLogin & "',FechaEdicion='" & strFechaDesde & Space(1) & Format(Time, "hh:mm") & "' " & _
                "WHERE NumOrden='" & Trim(tdgConsulta.Columns(0)) & "' AND CodFondo='" & strCodFondo & "' AND " & _
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

Private Sub cmdRentabilidad_Click()

    If CCur(lblMontoContado.Caption) = 0 Then Exit Sub
    
    If strCodTipoCobertura = Codigo_Tipo_Cobertura_Sintetico Then
        lblRentabilidad360.Caption = CStr((((CCur(lblMontoFuturo.Caption) / CCur(lblMontoContado.Caption)) ^ (360 / CInt(txtDiasPlazo.Text))) - 1) * 100)
        lblRentabilidad365.Caption = CStr((((CCur(lblMontoFuturo.Caption) / CCur(lblMontoContado.Caption)) ^ (365 / CInt(txtDiasPlazo.Text))) - 1) * 100)
    Else
        lblRentabilidad360.Caption = CStr(((((CCur(lblMontoFuturo.Caption) * CDec(txtTipoCambioFuturo.Text)) / (CCur(lblMontoFuturo.Caption) * CDec(txtTipoCambioSpot.Text))) ^ (360 / CInt(txtDiasPlazo.Text))) - 1) * 100)
        lblRentabilidad365.Caption = CStr(((((CCur(lblMontoFuturo.Caption) * CDec(txtTipoCambioFuturo.Text)) / (CCur(lblMontoFuturo.Caption) * CDec(txtTipoCambioSpot.Text))) ^ (365 / CInt(txtDiasPlazo.Text))) - 1) * 100)
    End If
    
End Sub

Private Sub dtpFechaLiquidacionDesde_CloseUp()

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
    
    If dtpFechaVencimiento.Value < dtpFechaLiquidacion.Value Then
        dtpFechaVencimiento.Value = dtpFechaLiquidacion.Value
    End If
    
    If strCodTitulo <> Valor_Caracter Then
        If dtpFechaVencimiento.Value < CVDate(lblFechaEmision.Caption) Then
            dtpFechaVencimiento.Value = CVDate(lblFechaEmision.Caption)
        End If
        
        If CVDate(lblFechaVencimiento.Caption) < dtpFechaVencimiento.Value Then
            dtpFechaVencimiento.Value = CVDate(lblFechaVencimiento.Caption)
        End If
    End If
    
    txtDiasPlazo.Text = CStr(DateDiff("d", dtpFechaLiquidacion.Value, dtpFechaVencimiento.Value))
'    Call CalculoTotal(0)
    dtpFechaPago.Value = dtpFechaVencimiento.Value
'    lblFechaVencimiento.Caption = CStr(dtpFechaVencimiento.Value)
'    lblFechaCupon.Caption = CStr(dtpFechaVencimiento.Value)
    
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
Private Sub CargarListas()

    Dim intRegistro As Integer

    '*** Fondos ***
    strSQL = "{ call up_ACSelDatosParametro(74,'" & gstrCodAdministradora & "','" & gstrCodFondoContable & "') }"
    CargarControlLista strSQL, cboFondo, arrFondo(), Valor_Caracter
    CargarControlLista strSQL, cboFondoOrden, arrFondoOrden(), Valor_Caracter
    
    If cboFondo.ListCount > 0 Then cboFondo.ListIndex = 0
        
    '*** Estados de la Orden ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='ESTORD' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboEstado, arrEstado(), Sel_Todos
    
    intRegistro = ObtenerItemLista(arrEstado(), Estado_Orden_Ingresada)
    If intRegistro >= 0 Then cboEstado.ListIndex = intRegistro
                
    '*** Emisor - Banco ***
    strSQL = "SELECT CodPersona CODIGO,DescripPersona DESCRIP FROM InstitucionPersona WHERE TipoPersona='" & Codigo_Tipo_Persona_Emisor & "' AND IndVigente='X' AND IndBanco='X' ORDER BY DescripPersona"
    CargarControlLista strSQL, cboEmisor, arrEmisor(), Sel_Defecto

    '*** Tipos de Cobertura ***
    strSQL = "SELECT CodParametro CODIGO, DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='TIPCOB' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboTipoCobertura, arrTipoCobertura(), Sel_Defecto

    '*** Moneda ***
    strSQL = "{ call up_ACSelDatos(2) }"
    CargarControlLista strSQL, cboMonedaCoberturada, arrMonedaCoberturada(), Sel_Defecto
    CargarControlLista strSQL, cboMonedaCobertura, arrMonedaCobertura(), Sel_Defecto
    
    '*** Indicadores de Cobertura ***
    strSQL = "SELECT CodParametro CODIGO,DescripParametro DESCRIP FROM AuxiliarParametro WHERE CodTipoParametro='INDCOB' ORDER BY DescripParametro"
    CargarControlLista strSQL, cboIndicador, arrIndicador(), Sel_Defecto

End Sub
Private Sub InicializarValores()

    Dim adoRegistro As ADODB.Recordset
    
    strEstado = Reg_Defecto
    tabCobertura.Tab = 0

    dtpFechaOrdenDesde.Value = gdatFechaActual
    dtpFechaOrdenHasta.Value = dtpFechaOrdenDesde.Value
    dtpFechaLiquidacionDesde.Value = Null
    dtpFechaLiquidacionHasta.Value = dtpFechaLiquidacionDesde.Value
            
    Set adoRegistro = New ADODB.Recordset
    With adoComm
        .CommandText = "SELECT CodFile FROM InversionFile " & _
            "WHERE TipoValor='" & Codigo_Valor_RentaVariable & "' AND IndInstrumento='X' AND IndVigente='X' " & _
            "ORDER BY DescripFile"
        Set adoRegistro = .Execute
                
        strCodigosFile = Valor_Caracter
        Do While Not adoRegistro.EOF
            If strCodigosFile <> Valor_Caracter Then strCodigosFile = strCodigosFile & ",'"
            
            strCodigosFile = strCodigosFile & Trim(adoRegistro("CodFile")) & "'"
        
            adoRegistro.MoveNext
        Loop
        adoRegistro.Close: Set adoRegistro = Nothing
        
        strCodigosFile = "('" & strCodigosFile & ")"
    End With
    
    Set cmdOpcion.FormularioActivo = Me
    Set cmdSalir.FormularioActivo = Me
    Set cmdAccion.FormularioActivo = Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmCobertura = Nothing
    Call OcultarReportes
    frmMainMdi.stbMdi.Panels(3).Text = "Acción"
    
End Sub

Private Sub lblMontoContado_Change()

    Call FormatoMillarEtiqueta(lblMontoContado, Decimales_Monto)
    
End Sub

Private Sub lblMontoFuturo_Change()

    Call FormatoMillarEtiqueta(lblMontoFuturo, Decimales_Monto)
    
End Sub

Private Sub lblMontoMFL1_Change()

    Call FormatoMillarEtiqueta(lblMontoMFL1, Decimales_Monto)
    
End Sub

Private Sub lblMontoMFL2_Change()

    Call FormatoMillarEtiqueta(lblMontoMFL2, Decimales_Monto)
    
End Sub

Private Sub lblRentabilidad360_Change()

    Call FormatoMillarEtiqueta(lblRentabilidad360, Decimales_Tasa)
    
End Sub

Private Sub lblRentabilidad365_Change()

    Call FormatoMillarEtiqueta(lblRentabilidad365, Decimales_Tasa)
    
End Sub

Private Sub lblValorNominal_Change()

    Call FormatoMillarEtiqueta(lblValorNominal, Decimales_Monto)
    
End Sub

Private Sub tdgConsulta_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)

    If ColIndex = 4 Then
        Call DarFormatoValor(Value, Decimales_Monto)
    End If
    
    If ColIndex = 7 Then
        Call DarFormatoValor(Value, Decimales_Monto)
    End If
    
End Sub

Private Sub txtDiasPlazo_Change()
   
    Call FormatoCajaTexto(txtDiasPlazo, 0)
    
    If IsNumeric(txtDiasPlazo.Text) Then
        dtpFechaVencimiento.Value = DateAdd("d", CInt(txtDiasPlazo.Text), CVDate(dtpFechaLiquidacion.Value))
    Else
        dtpFechaVencimiento.Value = dtpFechaLiquidacion.Value
    End If

    dtpFechaVencimiento_Change
    dtpFechaPago.Value = dtpFechaVencimiento.Value
    dtpFechaPago_Change
    
'    If CInt(txtDiasPlazo.Text) > 0 Then tabCobertura.TabEnabled(2) = True
    
End Sub

Private Sub txtDiasPlazo_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "N")
    
End Sub

Private Sub txtDiferencial_Change()

    Dim intBaseAnual As Integer
    
    Call FormatoCajaTexto(txtDiferencial, Decimales_Tasa)
        
    If Trim(txtDiferencial.Text) = Valor_Caracter Or Trim(txtTipoCambioSpot.Text) = Valor_Caracter Then Exit Sub
    
    intBaseAnual = 360
    If strCodBaseAnual = Codigo_Base_Actual_Actual Then intBaseAnual = 365
    If strCodBaseAnual = Codigo_Base_Actual_365 Then intBaseAnual = 365
    If strCodBaseAnual = Codigo_Base_30_365 Then intBaseAnual = 365
    If strCodBaseAnual = Codigo_Base_Actual_360 Then intBaseAnual = 360
    If strCodBaseAnual = Codigo_Base_30_360 Then intBaseAnual = 360
    
    If CDbl(txtDiferencial.Text) > 0 And CInt(txtDiasPlazo.Text) > 0 And CDbl(txtTipoCambioSpot.Text) > 0 Then
        txtTipoCambioFuturo.Text = ((1 + CDbl(txtDiferencial.Text) / 100) ^ (CInt(txtDiasPlazo.Text) / intBaseAnual)) * CDbl(txtTipoCambioSpot.Text)
    End If
    
End Sub


Private Sub txtDiferencial_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtDiferencial, Decimales_Tasa)
    
End Sub


Private Sub txtMontoCoberturado_Change()

    Call FormatoCajaTexto(txtMontoCoberturado, Decimales_Monto)
    
End Sub

Private Sub txtMontoCoberturado_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtMontoCoberturado, Decimales_Monto)
    
End Sub


Private Sub txtNemonico_Change()

    txtDescripOrden.Text = "Cobertura " & strNemotecnico & " - " & Trim(txtNemonico.Text)

End Sub

Private Sub txtNemonico_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
End Sub


Private Sub txtPorcenCobertura_Change()

    If strCodIndicador = Valor_Caracter Then Exit Sub
    
    Call FormatoCajaTexto(txtPorcenCobertura, Decimales_Tasa2)
    
    If strCodIndicador = Codigo_IndCobertura_Principal Then txtMontoCoberturado.Text = CStr(CCur(lblMontoMFL1.Caption) * CDbl(txtPorcenCobertura.Text) * 0.01)
    If strCodIndicador = Codigo_IndCobertura_Amortizacion Then txtMontoCoberturado.Text = CStr(CCur(lblMontoMFL2.Caption) * CDbl(txtPorcenCobertura.Text) * 0.01)
    
    txtDiferencial_Change
    
End Sub


Private Sub txtPorcenCobertura_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtPorcenCobertura, Decimales_Tasa2)
    
End Sub


Private Sub txtTipoCambio_Change()

    Call FormatoCajaTexto(txtTipoCambio, Decimales_TipoCambio)
    
End Sub

Private Sub txtTipoCambio_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtTipoCambio, Decimales_TipoCambio)
    
End Sub


Private Sub txtTipoCambioFuturo_Change()

    Call FormatoCajaTexto(txtTipoCambioFuturo, Decimales_TipoCambio)
    
    If (CCur(lblValorNominal.Caption) > 0 Or CCur(txtMontoCoberturado.Text) > 0) And CDbl(txtTipoCambioFuturo.Text) > 0 Then
        If strCodTipoCobertura = Codigo_Tipo_Cobertura_Sintetico Then
            If strCodMonedaTitulo = Codigo_Moneda_Local Then
                lblMontoContado.Caption = CStr((CCur(lblMontoMFL1.Caption) * CDbl(txtPorcenCobertura.Text) * 0.01) / CDbl(txtTipoCambio.Text))
                lblMontoFuturo.Caption = CStr(CCur(txtMontoCoberturado.Text) / CDbl(txtTipoCambioFuturo.Text))
            Else
                lblMontoContado.Caption = CStr((CCur(lblMontoMFL1.Caption) * CDbl(txtPorcenCobertura.Text) * 0.01) * CDbl(txtTipoCambio.Text))
                lblMontoFuturo.Caption = CStr(CCur(txtMontoCoberturado.Text) * CDbl(txtTipoCambioFuturo.Text))
            End If
        Else
            If strCodMonedaTitulo = Codigo_Moneda_Local Then
                Dim curValorContado     As Currency, strValorContado    As String
                
                lblMontoContado.Caption = CStr((CCur(txtMontoCoberturado.Text) * CDbl(txtPorcenCobertura.Text) * 0.01) / CDbl(txtTipoCambioFuturo.Text))
                curValorContado = CCur(lblMontoContado.Caption) * CDbl(txtTipoCambioSpot.Text)
                strValorContado = CStr(curValorContado)
                Call DarFormatoValor(strValorContado, Decimales_Monto)
                lblMontoContado.ToolTipText = strValorContado
                lblMontoFuturo.Caption = CStr(lblMontoContado.Caption)
            Else
                lblMontoContado.Caption = CStr((CCur(txtMontoCoberturado.Text) * CDbl(txtPorcenCobertura.Text) * 0.01) * CDbl(txtTipoCambioFuturo.Text))
                lblMontoFuturo.Caption = CStr(lblMontoContado.Caption)
            End If

        End If
    End If
    
End Sub

Private Sub txtTipoCambioFuturo_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtTipoCambioFuturo, Decimales_TipoCambio)
    
End Sub

Private Sub txtTipoCambioSpot_Change()

    Dim intBaseAnual As Integer
    
    Call FormatoCajaTexto(txtTipoCambioSpot, Decimales_TipoCambio)
    
    intBaseAnual = 360
    If strCodBaseAnual = Codigo_Base_Actual_Actual Then intBaseAnual = 365
    If strCodBaseAnual = Codigo_Base_Actual_365 Then intBaseAnual = 365
    If strCodBaseAnual = Codigo_Base_30_365 Then intBaseAnual = 365
    
    txtTipoCambioFuturo.Text = CStr(((1 + CDbl(txtDiferencial.Text) / 100) ^ (CInt(txtDiasPlazo.Text) / intBaseAnual)) * CDbl(txtTipoCambioSpot.Text))
    
End Sub

Private Sub txtTipoCambioSpot_KeyPress(KeyAscii As Integer)

    Call ValidaCajaTexto(KeyAscii, "M", txtTipoCambioSpot, Decimales_TipoCambio)
    
End Sub
